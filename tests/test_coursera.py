def click_About(baseurl, page):
    page.goto(f"{baseurl}#about", wait_until="domcontentloaded")
    try:
        btn = page.get_by_role("button", name="View all skills")
        btn.scroll_into_view_if_needed()
        btn.wait_for(state="visible", timeout=10000)
        btn.click(delay=300)
    except Exception as e:
        print(f"Could not click 'View all skills' button: {e}")

def click_Modules(baseurl, page):
    page.goto(f"{baseurl}#modules", wait_until="domcontentloaded")

def click_Course(baseurl, page):
    page.goto(f"{baseurl}#courses", wait_until="domcontentloaded")

def test_corsera(browser):
    page = browser.new_page()
    baseurl = "https://www.coursera.org/specializations/content-creation-and-copywriting"
    try:
        click_About(baseurl, page)
    except Exception as e:
        print(f"Error in About navigation: {e}")

    try:
        click_Modules(baseurl, page)
    except Exception as e:
        print(f"Error in Modules navigation: {e}")
    try:
        click_Course(baseurl, page)
    except Exception as e:
        print(f"Error in Course navigation: {e}")

    page.wait_for_timeout(3000)
    return page