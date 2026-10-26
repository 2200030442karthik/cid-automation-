"""
Sunsets (unlives) CIDs on Simplex.

For each CID, opens the B2B SkillUp page and sets
Course Basics > "Part of other product offering" to None (it is
normally Simplilearn SkillUp while the course is live), then Submit.
Nothing else on the page is changed. The page is reloaded afterwards
to confirm Simplex saved None.

Usage:
    python sunsettled.py
"""

import re

from playwright.sync_api import sync_playwright

from main import (
    connect_to_chrome,
    find_select_with_option_text,
)

from logo_update import (
    click_skillup_submit,
    open_course_page,
    parse_cid_input,
    skillup_url,
)


SKILLUP_PATH = "/admin/course/course-slh/courseId/"

# Every course's "Part of other product offering" dropdown has this
# option, so it identifies the right SELECT on the page.
LIVE_PRODUCT_OPTION = "Simplilearn SkillUp"


def is_none_option(text):
    # Matches "None", "-- None --", "none" etc.
    return re.sub(r"[^a-z]", "", text.lower()) == "none"


def find_product_offering_select(page):

    select = find_select_with_option_text(page, LIVE_PRODUCT_OPTION)

    options = select.locator("option")

    for i in range(options.count()):

        option = options.nth(i)

        if is_none_option(option.inner_text()):
            return select, option.get_attribute("value")

    available = ", ".join(
        options.nth(i).inner_text().strip()
        for i in range(options.count())
    )

    raise Exception(
        "'Part of other product offering' has no None option.\n"
        f"Options found: {available}"
    )


def selected_text(select):
    return select.locator("option:checked").inner_text().strip()


def sunset_course(page, course_id):

    url = skillup_url(course_id)
    print("Opening:", url)

    open_course_page(page, url, SKILLUP_PATH, course_id)

    product_select, none_value = find_product_offering_select(page)

    current = selected_text(product_select)
    print("Part of other product offering (current):", current)

    if is_none_option(current):
        print("Already None - nothing to submit.")
        return "already None"

    print("Setting Part of other product offering = None...")
    product_select.select_option(value=none_value)

    if not is_none_option(selected_text(product_select)):
        raise Exception(
            "Part of other product offering could not be set to None."
        )

    print("Part of other product offering: PASS - None")

    click_skillup_submit(page)

    # Reload and confirm Simplex actually kept None.
    print("Reloading B2B SkillUp page to verify the save...")
    open_course_page(page, url, SKILLUP_PATH, course_id)

    product_select, _ = find_product_offering_select(page)
    saved = selected_text(product_select)

    if not is_none_option(saved):
        raise Exception(
            "Save verification failed: Part of other product offering "
            f"is '{saved}', not None."
        )

    print("B2B SkillUp: SAVE VERIFIED - None (course unlived)")

    return f"unlived (was {current})"


# ============================================================
# MAIN
# ============================================================

def main():

    print("\n========================================")
    print("SIMPLEX SUNSET - UNLIVE COURSES")
    print("========================================")

    mode = input(
        "\nSunset 1 CID or a bunch of CIDs?\n"
        "1. Single CID\n"
        "2. Bunch of CIDs\n"
        "Enter 1 or 2: "
    ).strip()

    if mode not in ("1", "2"):
        raise Exception("Invalid selection. Enter 1 or 2.")

    if mode == "1":

        cids = parse_cid_input(input("\nEnter CID: "))

        if len(cids) != 1:
            raise Exception(
                "Single CID mode takes exactly one CID. "
                "Use option 2 for several."
            )

    else:

        cids = parse_cid_input(input(
            "\nEnter CIDs separated by comma or space "
            "(e.g. 8071, 8072, 8073): "
        ))

    print(f"\nCIDs to UNLIVE ({len(cids)}): {', '.join(cids)}")

    confirm = input(
        "These courses will be taken off SkillUp. "
        "Type YES to continue: "
    ).strip()

    if confirm != "YES":
        print("\nCancelled - nothing was changed.")
        input("\nPress ENTER to close...")
        return

    with sync_playwright() as p:

        page = connect_to_chrome(p)

        results = []

        for position, course_id in enumerate(cids, start=1):

            print("\n\n################################################")
            print(f"CID {position}/{len(cids)} - {course_id}")
            print("################################################")

            result = {"course_id": course_id, "status": "-"}

            try:

                result["status"] = sunset_course(page, course_id)

            except Exception as error:

                print(f"\nCID {course_id} FAILED:")
                print(error)

                result["error"] = str(error)

            results.append(result)

        # ====================================================
        # SUMMARY
        # ====================================================

        failures = [r for r in results if "error" in r]

        print("\n\n========================================")
        print("SUNSET SUMMARY")
        print("========================================")
        print(f"\nTotal CIDs : {len(results)}")
        print(f"Succeeded  : {len(results) - len(failures)}")
        print(f"Failed     : {len(failures)}")

        for r in results:

            if "error" in r:
                print(f"\n  CID {r['course_id']}: FAILED")
                print(f"    Error: {r['error'].splitlines()[0]}")
            else:
                print(f"\n  CID {r['course_id']}: OK ({r['status']})")

        print("\n========================================")

        input("\nPress ENTER to close...")


if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        print("\n==================================================")
        print("SUNSET STOPPED")
        print("==================================================")
        print("\nERROR:")
        print(error)
        print("\n==================================================")

        input("\nPress ENTER to close...")
