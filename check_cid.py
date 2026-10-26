"""
READ-ONLY checker for CIDs already created by main.py.

Opens each course's Simplex pages in a NEW tab of your logged-in Chrome,
reads the saved values and compares them with what main.py sets. The
only thing it types is the course URL into the SEO *search* box (to
look the SEO record up); it never ticks a box or clicks Save/Next/Submit.

Usage:
    python check_cid.py 8071
    python check_cid.py 8071 8072 8073
"""

import sys
import math

from playwright.sync_api import sync_playwright

from main import (
    SIMPLEX_BASE,
    SEO_LIST_URL,
    LEARNING_RATING,
    DUMMY_SEO_THUMBNAIL,
    DUMMY_COURSE_LOGO,
    SIMPLILEARN_NEW_LOGO_ID,
    SIMPLILEARN_NEW_LOGO_NAME,
    accreditor_logo_selected,
    create_simplex_seo_url,
    normalize_course_name_for_match,
    find_course_overview_field,
    get_visible_skill_fields,
    find_section_container_by_heading,
    find_select_with_option_text,
    find_duration_type_selects,
    find_input_near_text_generic,
)


class Report:

    def __init__(self):
        self.rows = []

    def check(self, label, ok, detail=""):
        self.rows.append(("PASS" if ok else "FAIL", label, detail))

    def info(self, label, detail):
        self.rows.append(("INFO", label, detail))

    def error(self, label, error):
        self.rows.append(("ERROR", label, str(error).splitlines()[0]))

    def print(self, course_id):

        print("\n========================================")
        print(f"CID {course_id} - CHECK REPORT")
        print("========================================")

        for status, label, detail in self.rows:
            print(f"[{status:5}] {label}" + (f": {detail}" if detail else ""))

        failed = [r for r in self.rows if r[0] in ("FAIL", "ERROR")]

        print("----------------------------------------")
        print(
            "RESULT: ALL GOOD" if not failed
            else f"RESULT: {len(failed)} item(s) need a look"
        )


def selected_text(select):
    return select.locator("option:checked").first.inner_text().strip()


def value_of(page, selector):
    loc = page.locator(selector)
    return loc.first.input_value().strip() if loc.count() else None


def checked(page, selector):
    loc = page.locator(selector)
    return loc.count() > 0 and loc.first.is_checked()


def radio_checked_in_section(page, heading, radio_text):
    """Read-only: is the radio labelled radio_text in this section checked?"""

    section = find_section_container_by_heading(
        page, heading, required_selector="input[type='radio']"
    )

    labels = section.locator("label").filter(has_text=radio_text)

    for i in range(labels.count()):

        label = labels.nth(i)
        target_id = label.get_attribute("for")

        if target_id:
            radio = section.locator(f'[id="{target_id}"]')
            if radio.count() and radio.first.get_attribute("type") == "radio":
                return radio.first.is_checked()

        radio = label.locator("input[type='radio']")
        if radio.count():
            return radio.first.is_checked()

    raise Exception(f"'{radio_text}' not found in '{heading}'")


def skills_checkbox_checked(page):
    """Read-only version of check_skills_validation_checkbox."""

    text = page.get_by_text(
        "I have reviewed and validated the skills listed.", exact=False
    )

    for i in range(min(text.count(), 5)):
        for level in range(1, 7):
            container = text.nth(i).locator("xpath=" + "/.." * level)
            boxes = container.locator("input[type='checkbox']")
            if boxes.count() == 1:
                return boxes.first.is_checked()

    raise Exception("Skills validation checkbox not found")


def open_page(page, url, must_contain):

    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_timeout(1800)

    if must_contain not in page.url:
        raise Exception(f"Unexpected page: {page.url}")


def check_basic(page, course_id, r, facts):

    open_page(
        page,
        f"{SIMPLEX_BASE}/admin/course/basic/courseId/{course_id}",
        f"/courseId/{course_id}",
    )

    name = value_of(page, "#name")
    facts["name"] = name
    r.info("Course Name", name)

    r.check("Display Name = Course Name", value_of(page, "#displayName") == name)
    r.check("Tags = Course Name", value_of(page, "#tags") == name)
    r.check(
        "Learning Rating", value_of(page, "#rating") == LEARNING_RATING,
        value_of(page, "#rating"),
    )

    expected_url = create_simplex_seo_url(name or "")
    facts["seo_url"] = expected_url
    # Simplex only shows the Course URL field when a course is created;
    # the edit page has no #url, so the SEO check below covers it.
    if page.locator("#url").count():
        r.check(
            "Course URL", value_of(page, "#url") == expected_url,
            value_of(page, "#url"),
        )
    else:
        r.info("Course URL", "not shown on edit page - checked via SEO record")
    r.check(
        "Course Logo (dummy)",
        value_of(page, "#image_home_page") == DUMMY_COURSE_LOGO,
    )
    r.check(
        f"Accreditation Heading = {SIMPLILEARN_NEW_LOGO_NAME}",
        value_of(page, "#accreditor_title") == SIMPLILEARN_NEW_LOGO_NAME,
        value_of(page, "#accreditor_title"),
    )
    r.check(
        f"Accreditation Logo = {SIMPLILEARN_NEW_LOGO_NAME}",
        accreditor_logo_selected(page, "accreditor_id", SIMPLILEARN_NEW_LOGO_ID),
    )

    level = page.locator("input[type='radio'][id^='level-']:checked")
    level_id = level.first.get_attribute("id") if level.count() else None
    facts["simplex_level"] = level_id.replace("level-", "") if level_id else None
    r.check("Course Level selected", bool(level_id), facts["simplex_level"])

    r.check("Online Self Learning", checked(page, "#trainingTypes-osl2"))
    r.check("Access Days 30", checked(page, "#trainingTypes-oslAccessDays-1"))

    weeks = value_of(page, "#durations")
    facts["weeks"] = weeks
    r.check("Duration (weeks) set", bool(weeks), weeks)
    r.check(
        "Duration Type = week",
        selected_text(page.locator("#duration_type")).lower().startswith("week"),
    )

    r.check("Hide From Website", checked(page, "#hideFromSearch"))
    r.check(
        "Course Available For = B2B Only",
        page.locator("#course_available_for").input_value() == "b2b_only",
    )

    primary_cat = selected_text(page.locator("#primary_label_id"))
    other_cats = page.locator("#label_id option:checked").all_inner_texts()
    r.info("Primary Category", primary_cat)
    r.check(
        "Other Category includes Primary Category",
        primary_cat.strip() in [c.strip() for c in other_cats],
        ", ".join(c.strip() for c in other_cats),
    )

    eid_select = page.locator("#primary_eLearning_id")
    eid_value = eid_select.input_value()
    eid_name = selected_text(eid_select)
    r.info("Primary EID", f"{eid_value} - {eid_name}")
    r.check(
        "EID course name matches Course Name",
        normalize_course_name_for_match(eid_name)
        == normalize_course_name_for_match(name),
    )

    overview = find_course_overview_field(page).input_value().strip()
    r.check(
        "Course Overview 300-350 chars",
        300 <= len(overview) <= 350,
        f"{len(overview)} chars",
    )
    r.info("Course Overview", overview)


def check_seo(page, r, facts):

    # Same Add More -> Search lookup main.py uses; the form is only
    # read, never submitted.
    page.goto(SEO_LIST_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    page.get_by_role("link", name="Add More").first.click()
    page.wait_for_timeout(1000)
    page.locator("#url").fill(facts["seo_url"])
    page.locator("#search").click()
    page.wait_for_timeout(1500)

    name = facts["name"]

    for label, selector in [
        ("SEO Meta Title", "#title"),
        ("SEO Meta Description", "#description"),
        ("SEO Meta Keyword", "#keyword"),
        ("SEO H1 Tag", "#h1Tag"),
    ]:
        r.check(f"{label} = Course Name", value_of(page, selector) == name,
                value_of(page, selector))

    r.check(
        "SEO Thumbnail (dummy)",
        value_of(page, "#thumb_image") == DUMMY_SEO_THUMBNAIL,
    )
    r.check("SEO No Index", checked(page, "#noindex"))


def check_skills(page, course_id, r):

    open_page(
        page,
        f"{SIMPLEX_BASE}/admin/course/course-moredetail/courseId/{course_id}",
        "/course-moredetail/courseId/",
    )

    fields = get_visible_skill_fields(page)
    skills = [fields.nth(i).input_value().strip() for i in range(fields.count())]

    r.check(
        "Skills Covered: exactly 5, none empty",
        len(skills) == 5 and all(skills),
        f"{len(skills)} found",
    )

    for i, skill in enumerate(skills, start=1):
        r.info(f"Skill {i}", skill)


def check_b2b(page, course_id, r, facts):

    open_page(
        page,
        f"{SIMPLEX_BASE}/admin/course/course-slh/courseId/{course_id}",
        "/course-slh/courseId/",
    )

    r.check(
        "Product offering = Simplilearn SkillUp",
        selected_text(find_select_with_option_text(page, "Simplilearn SkillUp"))
        .lower() == "simplilearn skillup",
    )

    lms_level = selected_text(find_select_with_option_text(page, "Beginner"))
    expected_simplex = {
        "beginner": "Foundational",
        "intermediate": "Foundational",
        "advanced": "Advanced",
    }.get(lms_level.lower())
    r.check(
        "LMS Level consistent with Basic Course Level",
        expected_simplex == facts["simplex_level"],
        f"LMS {lms_level} / Basic {facts['simplex_level']}",
    )

    r.check(
        "Course Info = Same as B2C",
        radio_checked_in_section(page, "Course Info", "Same as B2C"),
    )

    slh = find_input_near_text_generic(page, "SLH - Duration", ("input",))
    slh_plus = (
        find_input_near_text_generic(page, "SLH+ - Duration", ("input",))
        or find_input_near_text_generic(page, "SLH+ Duration", ("input",))
    )
    slh_hours = slh.input_value().strip() if slh else ""
    slh_plus_hours = slh_plus.input_value().strip() if slh_plus else ""

    r.check("SLH Duration set", slh_hours.isdigit(), slh_hours)
    r.check("SLH+ Duration = SLH Duration", slh_plus_hours == slh_hours,
            slh_plus_hours)

    if slh_hours.isdigit() and str(facts["weeks"]).isdigit():
        r.check(
            "SLH hours consistent with Basic weeks",
            math.ceil(int(slh_hours) / 5) == int(facts["weeks"]),
            f"{slh_hours} h -> {math.ceil(int(slh_hours) / 5)} week(s), "
            f"Basic has {facts['weeks']}",
        )

    type_1, type_2 = find_duration_type_selects(page)
    r.check("SLH Duration Type = Hours", selected_text(type_1).lower() == "hours")
    r.check("SLH+ Duration Type = Hours", selected_text(type_2).lower() == "hours")

    r.check(
        "Program Tag = Newly Launched",
        selected_text(find_select_with_option_text(page, "Newly Launched"))
        .lower() == "newly launched",
    )

    r.check("Skills validation checkbox", skills_checkbox_checked(page))
    r.check(
        "Course Images = Edit for SLH",
        radio_checked_in_section(page, "Course Images", "Edit for SLH"),
    )
    r.check(
        "Accreditation = Edit for SLH",
        checked(page, "#slh_accreditation_source-slh"),
    )
    r.check(
        f"Accreditation Logo = {SIMPLILEARN_NEW_LOGO_NAME}",
        accreditor_logo_selected(
            page, "slh_accreditor_id", SIMPLILEARN_NEW_LOGO_ID
        ),
    )


def check_course(page, course_id):

    r = Report()
    facts = {"name": None, "seo_url": None, "simplex_level": None, "weeks": None}

    for label, step in [
        ("Basic page", lambda: check_basic(page, course_id, r, facts)),
        ("SEO", lambda: check_seo(page, r, facts)),
        ("Skills Covered", lambda: check_skills(page, course_id, r)),
        ("B2B SkillUp", lambda: check_b2b(page, course_id, r, facts)),
    ]:
        try:
            step()
        except Exception as error:
            r.error(label, error)

    r.print(course_id)


def main():

    course_ids = [arg.strip() for arg in sys.argv[1:] if arg.strip()]

    if not course_ids or not all(cid.isdigit() for cid in course_ids):
        print("Usage: python check_cid.py <courseId> [<courseId> ...]")
        sys.exit(1)

    with sync_playwright() as p:

        browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        page = browser.contexts[0].new_page()

        try:
            for course_id in course_ids:
                check_course(page, course_id)
        finally:
            page.close()


if __name__ == "__main__":
    main()
