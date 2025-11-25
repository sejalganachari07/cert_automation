def click_About(baseurl, page):
    """Navigate to About section and expand all content"""
    page.goto(f"{baseurl}#about", wait_until="domcontentloaded")
    page.wait_for_timeout(2000)  # Wait for content to load
    
    # Click "View all skills" button
    try:
        btn = page.get_by_role("button", name="View all skills")
        btn.scroll_into_view_if_needed()
        btn.wait_for(state="visible", timeout=10000)
        btn.click(delay=300)
        print("✓ Clicked 'View all skills' button")
        page.wait_for_timeout(1000)
    except Exception as e:
        print(f"Could not click 'View all skills' button: {e}")
    
    # Click all "Read more" buttons in About section
    click_all_read_more_buttons(page, section_name="About")

def click_all_read_more_buttons(page, section_name=""):
    """Click all 'Read more' buttons on the current view"""
    try:
        # Find all "Read more" buttons
        read_more_buttons = page.get_by_role("button", name="Read more").all()
        
        if read_more_buttons:
            print(f"Found {len(read_more_buttons)} 'Read more' button(s) in {section_name} section")
            
            for idx, btn in enumerate(read_more_buttons, 1):
                try:
                    # Scroll button into view
                    btn.scroll_into_view_if_needed()
                    page.wait_for_timeout(500)
                    
                    # Check if button is visible and enabled
                    if btn.is_visible() and btn.is_enabled():
                        btn.click(delay=300)
                        print(f"✓ Clicked 'Read more' button {idx}/{len(read_more_buttons)}")
                        page.wait_for_timeout(800)
                    else:
                        print(f"⊘ 'Read more' button {idx} not clickable (hidden or disabled)")
                        
                except Exception as e:
                    print(f"✗ Error clicking 'Read more' button {idx}: {e}")
        else:
            print(f"No 'Read more' buttons found in {section_name} section")
            
    except Exception as e:
        print(f"Error finding 'Read more' buttons: {e}")

def click_Modules(baseurl, page):
    """Navigate to Modules section and expand content"""
    page.goto(f"{baseurl}#modules", wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    
    # Click all "Read more" buttons in Modules section
    click_all_read_more_buttons(page, section_name="Modules")

def click_Course(baseurl, page):
    """Navigate to Courses section and expand content"""
    page.goto(f"{baseurl}#courses", wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    
    # Click all "Read more" buttons in Courses section
    click_all_read_more_buttons(page, section_name="Courses")

def expand_all_sections(baseurl, page):
    """Comprehensive function to expand all collapsible content across all sections"""
    sections = [
        ("Modules", f"{baseurl}#modules")
    ]
    
    for section_name, url in sections:
        print(f"\n{'='*50}")
        print(f"Processing {section_name} section...")
        print('='*50)
        
        try:
            page.goto(url, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            
            # Special handling for About section (has "View all skills")
            if section_name == "About":
                try:
                    btn = page.get_by_role("button", name="View all skills")
                    btn.scroll_into_view_if_needed()
                    btn.wait_for(state="visible", timeout=5000)
                    btn.click(delay=300)
                    print("✓ Clicked 'View all skills' button")
                    page.wait_for_timeout(1000)
                except Exception as e:
                    print(f"'View all skills' button not found or not clickable: {e}")
            
            # Click all "Read more" buttons in current section
            click_all_read_more_buttons(page, section_name=section_name)
            
        except Exception as e:
            print(f"✗ Error processing {section_name} section: {e}")

def test_coursera(browser):
    """Main test function"""
    page = browser.new_page()
    baseurl = "https://www.coursera.org/specializations/content-creation-and-copywriting"
    
    print("Starting Coursera page expansion test...")
    print(f"URL: {baseurl}\n")
    
    # Option 1: Use individual functions
    
    click_About(baseurl, page)

    click_Modules(baseurl, page)
        
    
    # Option 2: Or use the comprehensive function (commented out to avoid duplication)
    # expand_all_sections(baseurl, page)
    
    page.wait_for_timeout(3000)
    print("\n" + "="*50)
    print("All sections processed. Page ready for scraping.")
    print("="*50)
    

