"""
Adds the Simplilearn_New accreditation logo to CIDs that already exist.

For each CID:
  1. Basic page   - Accreditation Heading = Simplilearn_New and
                    Accreditation Logos gets 541 - Simplilearn_New,
                    then Save.
  2. B2B SkillUp  - Accreditation Configuration = Edit for SLH and
                    Logos gets 541 - Simplilearn_New, then Submit.
Both pages are reloaded afterwards to confirm Simplex saved the values.

Usage:
    python logo_update.py
"""

import re

from playwright.sync_api import sync_playwright

from main import (
    SIMPLEX_BASE,
    B2B_SKILLUP_PATH,
    SIMPLILEARN_NEW_LOGO_ID,
    SIMPLILEARN_NEW_LOGO_NAME,
    accreditor_logo_selected,
    add_accreditor_logo,
    check_skills_validation_checkbox,
    connect_to_chrome,
)


# Simplex allows at most this many logos in the Basic page's
# "Accreditation Logos for Old Course, SEM, Enterprise Page" list.
MAX_BASIC_ACCREDITATION_LOGOS = 2

SKILLS_VALIDATION_TEXT = "I have reviewed and validated the skills listed."


def basic_url(course_id):
    return f"{SIMPLEX_BASE}/admin/course/basic/courseId/{course_id}"


def skillup_url(course_id):
    return SIMPLEX_BASE + B2B_SKILLUP_PATH.format(course_id)


def open_course_page(page, url, expected_path, course_id):

    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_timeout(1500)

    print("Current URL:", page.url)

    if expected_path not in page.url or f"/courseId/{course_id}" not in page.url:
        raise Exception(
            f"Could not open CID {course_id} - check the CID exists and "
            "that you are still logged into Simplex.\n"
            f"Expected: {url}\n"
            f"Current URL: {page.url}"
        )


def parse_cid_input(raw):

    cids = [part for part in re.split(r"[\s,;]+", raw.strip()) if part]

    if not cids:
        raise Exception("No CID was entered.")

    invalid = [cid for cid in cids if not cid.isdigit()]

    if invalid:
        raise Exception(
            "CIDs must be numbers only.\n"
            f"Invalid: {', '.join(invalid)}"
        )

    # Keep the order the user typed, drop repeats.
    return list(dict.fromkeys(cids))


# ============================================================
# BASIC PAGE
# ============================================================

def update_basic_logo(page, course_id):

    print("\n----------------------------------------")
    print("BASIC PAGE - LOGOS")
    print("----------------------------------------")

    url = basic_url(course_id)
    print("Opening:", url)

    open_course_page(page, url, "/admin/course/basic/courseId/", course_id)

    heading = page.locator("#accreditor_title")

    if heading.count() == 0:
        raise Exception(
            "Accreditation Heading field (#accreditor_title) not found."
        )

    if page.locator("#accreditor_id").count() == 0:
        raise Exception(
            "Accreditation Logos list (#accreditor_id) not found."
        )

    heading_done = heading.input_value().strip() == SIMPLILEARN_NEW_LOGO_NAME
    logo_done = accreditor_logo_selected(
        page, "accreditor_id", SIMPLILEARN_NEW_LOGO_ID
    )

    if heading_done and logo_done:
        print(
            f"Accreditation Heading and Logo already {SIMPLILEARN_NEW_LOGO_NAME}"
            " - nothing to save."
        )
        return "already set"

    if not logo_done:

        existing_logos = page.locator(
            "ul.accreditor_id_sortable_list li[data-cid]"
        ).count()

        if existing_logos >= MAX_BASIC_ACCREDITATION_LOGOS:
            raise Exception(
                f"Accreditation Logos already has {existing_logos} logos "
                f"(max {MAX_BASIC_ACCREDITATION_LOGOS} allowed), so "
                f"{SIMPLILEARN_NEW_LOGO_NAME} cannot be added. Remove one "
                "manually and run this CID again. Nothing was saved."
            )

    print(f"Setting Accreditation Heading = {SIMPLILEARN_NEW_LOGO_NAME}...")
    heading.fill(SIMPLILEARN_NEW_LOGO_NAME)

    if heading.input_value() != SIMPLILEARN_NEW_LOGO_NAME:
        raise Exception("Accreditation Heading could not be filled.")

    print("Accreditation Heading: PASS")

    print(
        f"Selecting {SIMPLILEARN_NEW_LOGO_ID} - {SIMPLILEARN_NEW_LOGO_NAME} "
        "in Accreditation Logos..."
    )
    add_accreditor_logo(page, "accreditor_id", SIMPLILEARN_NEW_LOGO_ID)
    print("Accreditation Logo: PASS")

    print("\nClicking Save...")

    submit = page.locator("#submit")

    if submit.count() == 0:
        raise Exception("Basic page Save button (#submit) was not found.")

    submit.click()
    page.wait_for_timeout(2000)

    print("Current URL:", page.url)

    # Reload and confirm Simplex actually kept both values.
    print("Reloading Basic page to verify the save...")
    open_course_page(page, url, "/admin/course/basic/courseId/", course_id)

    if page.locator("#accreditor_title").input_value().strip() != \
            SIMPLILEARN_NEW_LOGO_NAME:
        raise Exception(
            "Basic page save verification failed: Accreditation Heading "
            f"is not {SIMPLILEARN_NEW_LOGO_NAME}."
        )

    if not accreditor_logo_selected(
        page, "accreditor_id", SIMPLILEARN_NEW_LOGO_ID
    ):
        raise Exception(
            "Basic page save verification failed: "
            f"{SIMPLILEARN_NEW_LOGO_NAME} logo is not saved."
        )

    print("Basic page: SAVE VERIFIED")

    return "updated"


# ============================================================
# B2B SKILLUP PAGE
# ============================================================

def click_skillup_submit(page):
    """Click the B2B SkillUp Submit and fail if Simplex rejects it."""

    print("\nClicking Submit...")

    submit_candidates = [
        page.get_by_role("button", name="Submit", exact=True),
        page.get_by_role("link", name="Submit", exact=True),
        page.locator("input[type='submit']"),
        page.locator("button[type='submit']"),
        page.locator("#submit"),
    ]

    submit = None

    for candidate in submit_candidates:

        if candidate.count() > 0 and candidate.first.is_visible():
            submit = candidate.first
            break

    if submit is None:
        raise Exception(
            "Submit button could not be found on B2B SkillUp page."
        )

    submit.click()
    page.wait_for_timeout(2200)

    error_text = page.get_by_text(
        "Unable to save changes. Please check your inputs.",
        exact=False
    )

    if error_text.count() > 0 and error_text.first.is_visible():
        raise Exception(
            "B2B SkillUp was NOT saved by Simplex.\n"
            "Simplex displayed: Unable to save changes. Please check your "
            "inputs.\nThe browser has been left on the B2B SkillUp page "
            "for inspection."
        )

    print("Current URL:", page.url)


def update_skillup_logo(page, course_id):

    print("\n----------------------------------------")
    print("B2B SKILLUP - ACCREDITATION CONFIGURATION")
    print("----------------------------------------")

    url = skillup_url(course_id)
    print("Opening:", url)

    open_course_page(page, url, "/admin/course/course-slh/courseId/", course_id)

    # Selected by id: Course Images has its own "Edit for SLH" radio.
    accreditation_edit = page.locator("#slh_accreditation_source-slh")

    if accreditation_edit.count() == 0:
        raise Exception(
            "Accreditation Configuration: Edit for SLH radio "
            "(#slh_accreditation_source-slh) was not found."
        )

    if accreditation_edit.is_checked() and accreditor_logo_selected(
        page, "slh_accreditor_id", SIMPLILEARN_NEW_LOGO_ID
    ):
        print(
            f"Edit for SLH and {SIMPLILEARN_NEW_LOGO_NAME} already set"
            " - nothing to submit."
        )
        return "already set"

    print("Selecting Edit for SLH...")
    accreditation_edit.check()

    if not accreditation_edit.is_checked():
        raise Exception(
            "Accreditation Configuration: Edit for SLH could not be selected."
        )

    # Simplex keeps the Logos list disabled until Edit for SLH is chosen.
    page.wait_for_selector(
        "#slh_accreditor_id:not([disabled])",
        timeout=5000
    )

    print("Accreditation Configuration - Edit for SLH: PASS")

    print(
        f"Selecting {SIMPLILEARN_NEW_LOGO_ID} - {SIMPLILEARN_NEW_LOGO_NAME} "
        "in Logos..."
    )
    add_accreditor_logo(page, "slh_accreditor_id", SIMPLILEARN_NEW_LOGO_ID)
    print("Accreditation Logo: PASS")

    # Simplex rejects the Submit when "I have reviewed and validated the
    # skills listed." is unticked - tick it only if it isn't already.
    if page.get_by_text(SKILLS_VALIDATION_TEXT, exact=False).count() > 0:

        print(f"Checking '{SKILLS_VALIDATION_TEXT}'...")
        check_skills_validation_checkbox(page)
        print("Skills Validation: PASS - Checked")

    else:

        print("Skills validation checkbox not on this page - skipping.")

    click_skillup_submit(page)

    # Reload and confirm Simplex actually kept both values.
    print("Reloading B2B SkillUp page to verify the save...")
    open_course_page(page, url, "/admin/course/course-slh/courseId/", course_id)

    if not page.locator("#slh_accreditation_source-slh").is_checked():
        raise Exception(
            "B2B SkillUp save verification failed: Accreditation "
            "Configuration is not Edit for SLH."
        )

    if not accreditor_logo_selected(
        page, "slh_accreditor_id", SIMPLILEARN_NEW_LOGO_ID
    ):
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"{SIMPLILEARN_NEW_LOGO_NAME} logo is not saved."
        )

    print("B2B SkillUp: SAVE VERIFIED")

    return "updated"


# ============================================================
# MAIN
# ============================================================

def main():

    print("\n========================================")
    print("SIMPLEX LOGO UPDATE - SIMPLILEARN_NEW")
    print("========================================")

    mode = input(
        "\nUpdate 1 CID or a bunch of CIDs?\n"
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

    print(f"\nCIDs to update ({len(cids)}): {', '.join(cids)}")

    with sync_playwright() as p:

        page = connect_to_chrome(p)

        results = []

        for position, course_id in enumerate(cids, start=1):

            print("\n\n################################################")
            print(f"CID {position}/{len(cids)} - {course_id}")
            print("################################################")

            result = {"course_id": course_id, "basic": "-", "b2b": "-"}

            try:

                result["basic"] = update_basic_logo(page, course_id)
                result["b2b"] = update_skillup_logo(page, course_id)

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
        print("LOGO UPDATE SUMMARY")
        print("========================================")
        print(f"\nTotal CIDs : {len(results)}")
        print(f"Succeeded  : {len(results) - len(failures)}")
        print(f"Failed     : {len(failures)}")

        for r in results:

            status = "FAILED" if "error" in r else "OK"

            print(
                f"\n  CID {r['course_id']}: {status} "
                f"(Basic: {r['basic']}, B2B SkillUp: {r['b2b']})"
            )

            if "error" in r:
                print(f"    Error: {r['error'].splitlines()[0]}")

        print("\n========================================")

        input("\nPress ENTER to close...")


if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        print("\n==================================================")
        print("LOGO UPDATE STOPPED")
        print("==================================================")
        print("\nERROR:")
        print(error)
        print("\n==================================================")

        input("\nPress ENTER to close...")
