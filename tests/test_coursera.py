def test_corsera(browser):
    page = browser.new_page()
    page.goto("https://www.coursera.org/")
    page.type("#search-autocomplete-input", "Machine Learning")
    page.keyboard.press("Enter")
    page.wait_for_timeout(10000)


    return page
