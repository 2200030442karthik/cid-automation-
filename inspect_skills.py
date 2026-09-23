from playwright.sync_api import sync_playwright


ADDITIONAL_CONTENT_URL = (
    "https://i9simplex.simplilearn.com/admin/course/"
    "course-moredetail/courseId/8007"
)


with sync_playwright() as p:

    print("=" * 50)
    print("SIMPLEX SKILLS INSPECTION")
    print("=" * 50)

    print("\nConnecting to existing Chrome...")

    browser = p.chromium.connect_over_cdp(
        "http://127.0.0.1:9222"
    )

    print("Connected to Chrome.")

    context = browser.contexts[0]

    if context.pages:
        page = context.pages[0]
    else:
        page = context.new_page()

    print("\nOpening Additional Content page...")

    page.goto(
        ADDITIONAL_CONTENT_URL,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(2000)

    print("Current URL:")
    print(page.url)

    print("\n" + "=" * 50)
    print("INPUT ELEMENTS")
    print("=" * 50)

    inputs = page.locator("input")

    print(f"Total input elements: {inputs.count()}")

    for i in range(inputs.count()):

        element = inputs.nth(i)

        try:
            print(
                f"\nINPUT {i + 1}"
                f"\n  type  : {element.get_attribute('type')}"
                f"\n  id    : {element.get_attribute('id')}"
                f"\n  name  : {element.get_attribute('name')}"
                f"\n  value : {element.input_value()}"
                f"\n  class : {element.get_attribute('class')}"
            )
        except Exception as e:
            print(f"  Could not inspect: {e}")

    print("\n" + "=" * 50)
    print("BUTTONS")
    print("=" * 50)

    buttons = page.locator(
        "button, input[type='button'], input[type='submit'], a"
    )

    print(f"Total buttons/links: {buttons.count()}")

    for i in range(buttons.count()):

        element = buttons.nth(i)

        try:
            text = element.inner_text().strip()

            value = element.get_attribute("value")

            if text or value:
                print(
                    f"\nBUTTON/LINK {i + 1}"
                    f"\n  text  : {text}"
                    f"\n  value : {value}"
                    f"\n  id    : {element.get_attribute('id')}"
                    f"\n  class : {element.get_attribute('class')}"
                )

        except Exception:
            pass

    print("\n" + "=" * 50)
    print("SKILLS SECTION TEXT")
    print("=" * 50)

    body_text = page.locator("body").inner_text()

    lines = body_text.splitlines()

    for i, line in enumerate(lines):

        if "Skills Covered" in line:

            print("\nRelevant page area:")

            start = max(0, i - 3)
            end = min(len(lines), i + 15)

            for x in lines[start:end]:
                print(x)

    print("\n" + "=" * 50)
    print("INSPECTION COMPLETE")
    print("=" * 50)

    input("\nPress ENTER to close...")