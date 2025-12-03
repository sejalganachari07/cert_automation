from playwright.sync_api import sync_playwright
import re
import os
import sys
import pandas as pd

try:
    sys.stdout.reconfigure(errors="ignore")
except Exception:
    pass


def wait(page, ms: int = 300):
    """Minimal wait wrapper."""
    page.wait_for_timeout(ms)


def safe_click(page, locator, timeout: int = 2000) -> bool:
    """Click safely with scroll, returns True on success."""
    try:
        if not locator or not locator.is_visible(timeout=timeout):
            return False
        locator.scroll_into_view_if_needed()
        wait(page, 200)
        locator.click(timeout=timeout)
        wait(page, 200)
        return True
    except Exception:
        return False


def setup_ad_blocker(page):
    """Set up continuous ad blocking using MutationObserver - runs automatically."""
    page.evaluate("""
        () => {
            // Cleanup function
            window.removeAds = () => {
                // Remove promotional/Black Friday overlays
                const promoSelectors = [
                    '[class*="black-friday"]', '[class*="promo"]', '[class*="sale"]',
                    '[id*="black-friday"]', '[class*="discount"]', '[class*="banner"]'
                ];
                promoSelectors.forEach(sel => {
                    document.querySelectorAll(sel).forEach(el => el.remove());
                });
                
                // Remove modals (but not FAQ)
                document.querySelectorAll('[role="dialog"], [role="alertdialog"], [class*="modal"]').forEach(el => {
                    const text = (el.textContent || '').toLowerCase();
                    if (!text.includes('frequently asked') && !text.includes('faq')) {
                        el.remove();
                    }
                });
                
                // Remove high z-index overlays (ads, popups)
                document.querySelectorAll('*').forEach(el => {
                    const style = window.getComputedStyle(el);
                    const zIndex = parseInt(style.zIndex);
                    if ((style.position === 'fixed' || style.position === 'absolute') && zIndex > 999) {
                        const text = (el.textContent || '').toLowerCase();
                        const classList = (el.className || '').toLowerCase();
                        // Remove if it looks like an ad/overlay (not FAQ)
                        if (!text.includes('frequently asked') && !text.includes('faq') &&
                            (classList.includes('overlay') || classList.includes('backdrop') || 
                             classList.includes('modal') || classList.includes('popup') ||
                             zIndex > 9999)) {
                            el.remove();
                        }
                    }
                });
                
                // Accept cookies immediately
                const cookieBtn = document.querySelector('#onetrust-accept-btn-handler, [id*="cookie-accept"]');
                if (cookieBtn) cookieBtn.click();
                
                // Reset scroll lock
                document.body.style.overflow = 'visible';
                document.documentElement.style.overflow = 'visible';
            };
            
            // Run cleanup immediately
            window.removeAds();
            
            // Set up MutationObserver to watch for new ads
            const observer = new MutationObserver((mutations) => {
                window.removeAds();
            });
            
            // Observe document for any DOM changes
            observer.observe(document.body, {
                childList: true,
                subtree: true
            });
            
            // Also run cleanup every 2 seconds as backup
            setInterval(window.removeAds, 2000);
            
            console.log('✅ Ad blocker activated');
        }
    """)


def block_unwanted_elements(page):
    """Permanently block FAQ accordions and unwanted buttons with CSS."""
    page.add_style_tag(content="""
        /* Block FAQ accordions - CRITICAL */
        button[aria-label*='frequently asked' i],
        button[aria-label*='FAQ' i],
        [data-testid*='faq' i],
        [data-e2e*='faq' i],
        div[class*='faq' i] button[aria-expanded],
        section[class*='faq' i] button[aria-expanded] {
            pointer-events: none !important;
            opacity: 0.3 !important;
            cursor: not-allowed !important;
        }
        
        /* Block Explore buttons */
        button[data-testid*='explore' i],
        button[aria-label*='Explore' i] {
            pointer-events: none !important;
            display: none !important;
        }
        
        /* Aggressively hide promotional modals */
        [class*="black-friday" i],
        [class*="promo-modal" i],
        [class*="sale-modal" i],
        [class*="discount-modal" i],
        [data-track*="promo" i],
        [data-track*="sale" i] {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            z-index: -9999 !important;
            pointer-events: none !important;
        }
        
        /* Hide cookie banners */
        #onetrust-banner-sdk,
        [id*="cookie-banner" i],
        [class*="cookie-banner" i] {
            display: none !important;
        }
    """)


def close_popups(page):
    """Manual popup cleanup - call when needed."""
    try:
        # Press Escape key
        page.keyboard.press("Escape")
        wait(page, 100)
        
        # Trigger manual cleanup
        page.evaluate("if (window.removeAds) window.removeAds();")
    except Exception:
        pass


def is_faq_element(element) -> bool:
    """Check if an element is FAQ-related."""
    try:
        aria_label = (element.get_attribute("aria-label") or "").lower()
        data_e2e = (element.get_attribute("data-e2e") or "").lower()
        btn_class = (element.get_attribute("class") or "").lower()
        btn_text = (element.text_content() or "").lower()
        
        faq_keywords = ['faq', 'frequently asked', 'question']
        return any(kw in text for kw in faq_keywords 
                  for text in [aria_label, data_e2e, btn_class, btn_text])
    except:
        return False


def click_read_more_buttons(page):
    """Click only valid Read more buttons, excluding FAQ/Explore."""
    try:
        buttons = page.locator('button:has-text("Read more")').all()
        clicked = 0
        
        for btn in buttons:
            if not btn.is_visible(timeout=500):
                continue
                
            aria_label = (btn.get_attribute("aria-label") or "").lower()
            skip_keywords = ['explore', 'frequently asked', 'faq', 'offered by', 'partner']
            
            if any(kw in aria_label for kw in skip_keywords):
                continue
            
            if safe_click(page, btn, timeout=1000):
                clicked += 1
        
        if clicked > 0:
            print(f"    ✓ Clicked {clicked} Read more button(s)")
    except Exception:
        pass


def process_about_section(page, base_url):
    """Process About section - expand skills and read more."""
    print("\n" + "="*70)
    print("📍 STEP 1: ABOUT SECTION")
    print("="*70)
    
    try:
        page.goto(f"{base_url}#about", wait_until="load")
        wait(page, 800)  # Let ad blocker work
        
        # Expand skills
        skills_btn = page.locator('button:has-text("View all skills")').first
        if safe_click(page, skills_btn, timeout=2000):
            print("    ✓ Expanded skills")
        
        # Click read more
        click_read_more_buttons(page)
        print("  ✅ About section complete")
    except Exception as e:
        print(f"  ⚠️ About section error: {str(e)[:50]}")
    
    print("="*70)


def process_modules_section(page, base_url):
    """Process Modules - expand module accordions (NOT FAQ)."""
    print("\n" + "="*70)
    print("📍 STEP 2: MODULES SECTION")
    print("="*70)
    
    try:
        page.goto(f"{base_url}#modules", wait_until="load")
        wait(page, 800)  # Let ad blocker work
        
        # Find all accordion buttons
        all_accordions = page.locator('button[aria-expanded]').all()
        
        # Filter out FAQ accordions
        module_buttons = [btn for btn in all_accordions if not is_faq_element(btn)]
        
        if not module_buttons:
            # Try courses section
            page.goto(f"{base_url}#courses", wait_until="load")
            wait(page, 800)
            all_accordions = page.locator('button[aria-expanded]').all()
            module_buttons = [btn for btn in all_accordions if not is_faq_element(btn)]
        
        if module_buttons:
            total = len(module_buttons)
            print(f"  📊 Found {total} module(s) (FAQ excluded)")
            
            for idx, btn in enumerate(module_buttons, 1):
                try:
                    is_expanded = btn.get_attribute("aria-expanded")
                    if is_expanded == "true":
                        continue
                    
                    # Double-check not FAQ before clicking
                    if is_faq_element(btn):
                        print(f"    [{idx}/{total}] ⊘ Skipped FAQ")
                        continue
                    
                    btn.scroll_into_view_if_needed()
                    wait(page, 200)
                    
                    if safe_click(page, btn, timeout=1500):
                        print(f"    [{idx}/{total}] ✓ Expanded")
                except Exception:
                    pass
            
            print("  ✅ All modules processed")
        
        # Scroll through content
        for _ in range(3):
            page.evaluate("window.scrollBy({top: 400, behavior: 'smooth'})")
            wait(page, 300)
        
        # Click read more in modules
        click_read_more_buttons(page)
        
    except Exception as e:
        print(f"  ⚠️ Modules error: {str(e)[:50]}")
    
    print("="*70)


def scroll_to_bottom(page):
    """Scroll to bottom to load all content."""
    print("\n" + "="*70)
    print("📍 STEP 3: SCROLL TO BOTTOM")
    print("="*70)
    
    try:
        last_height = page.evaluate("document.body.scrollHeight")
        scroll_count = 0
        max_scrolls = 30
        
        while scroll_count < max_scrolls:
            page.evaluate("window.scrollBy(0, window.innerHeight * 0.8)")
            wait(page, 300)
            
            scroll_count += 1
            new_height = page.evaluate("document.body.scrollHeight")
            current_pos = page.evaluate("window.pageYOffset + window.innerHeight")
            
            if current_pos >= new_height - 100:
                print(f"    ✓ Reached bottom after {scroll_count} scrolls")
                break
            
            if new_height == last_height:
                break
            
            last_height = new_height
        
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        print("  ✅ Scroll complete")
    except Exception as e:
        print(f"  ⚠️ Scroll error: {str(e)[:50]}")
    
    print("="*70)


def prepare_for_pdf(page):
    """Final preparation - remove overlays."""
    print("\n" + "="*70)
    print("📍 STEP 4: PREPARE FOR PDF")
    print("="*70)
    
    try:
        # Manual cleanup before PDF
        close_popups(page)
        page.evaluate("window.scrollTo({top: 0, behavior: 'smooth'})")
        wait(page, 500)
        
        page.evaluate("""
            () => {
                // Remove dialogs
                document.querySelectorAll('[role="dialog"], [role="alertdialog"]').forEach(el => el.remove());
                
                // Hide fixed headers/navs
                document.querySelectorAll('header, nav, footer').forEach(el => {
                    const style = window.getComputedStyle(el);
                    if (style.position === 'fixed' || style.position === 'sticky') {
                        el.style.display = 'none';
                    }
                });
                
                // Show all content
                document.querySelectorAll('main, article, section').forEach(el => {
                    el.style.display = 'block';
                    el.style.visibility = 'visible';
                    el.style.overflow = 'visible';
                    el.style.maxHeight = 'none';
                });
                
                document.body.style.overflow = 'visible';
            }
        """)
        
        print("  ✅ Page prepared")
    except Exception as e:
        print(f"  ⚠️ Preparation error: {str(e)[:50]}")
    
    print("="*70)


def generate_pdf(page, base_url, output_dir=".", custom_name=None):
    """Generate PDF with selectable text."""
    print("\n" + "="*70)
    print("📍 STEP 5: GENERATE PDF")
    print("="*70)
    
    try:
        page.emulate_media(media="print")
        
        # Extract course name
        course_name = "Coursera_Course"
        try:
            title = page.locator('h1').first.text_content().strip()
            course_name = re.sub(r'[<>:"/\\|?*]', '_', title)
        except:
            pass
        
        # Create filename
        url_slug = base_url.split("/")[-1].split("?")[0]
        if custom_name:
            safe_custom = re.sub(r'[<>:"/\\|?*]', '_', str(custom_name).strip())
            filename = f"{safe_custom}_{url_slug}.pdf"
        else:
            filename = f"{course_name}_{url_slug}.pdf"
        
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)[:200]
        os.makedirs(output_dir, exist_ok=True)
        full_path = os.path.join(output_dir, filename)
        
        print(f"  💾 Saving: {filename}")
        
        # Prepare page for PDF rendering
        page.set_viewport_size({"width": 1200, "height": 800})
        page.wait_for_selector('main, article', timeout=5000)
        
        # Scroll to load lazy content
        for _ in range(10):
            page.evaluate("window.scrollBy(0, window.innerHeight)")
            wait(page, 300)
        
        # Remove fixed elements that break PDF
        page.evaluate("""
            document.querySelectorAll("*").forEach(el => {
                const style = getComputedStyle(el);
                if (style.position === "fixed" || style.position === "sticky") {
                    el.style.position = "static";
                }
            });
            
            document.querySelectorAll('main, article, section').forEach(el => {
                el.style.display = 'block';
                el.style.visibility = 'visible';
                el.style.opacity = '1';
            });
        """)
        
        wait(page, 1000)
        
        # Generate PDF
        page.pdf(
            path=full_path,
            format="A4",
            print_background=True,
            prefer_css_page_size=False,
            margin={"top": "0.4in", "bottom": "0.4in", "left": "0.5in", "right": "0.5in"},
            scale=0.90
        )
        
        print(f"  ✅ PDF SAVED: {full_path}")
        print("="*70)
        return full_path
    
    except Exception as e:
        print(f"  ❌ PDF generation failed: {str(e)}")
        print("="*70)
        return None


def detect_excel_columns(df):
    """Detect URL and name columns flexibly."""
    def find_col(names):
        lower_map = {str(c).strip().lower(): c for c in df.columns}
        for name in names:
            if name in lower_map:
                return lower_map[name]
        return None
    
    url_col = find_col(["url", "course_url", "course url", "link", "course_link"])
    name_col = find_col(["name", "course_name", "course name", "title", "course_title"])
    return url_col, name_col


def main():
    """Main execution: batch process URLs from Excel."""
    excel_path = "courses.xlsx"
    output_dir = "pdfs"
    
    print("\n" + "="*70)
    print("🚀 COURSERA SCRAPER - BATCH MODE")
    print("="*70)
    print(f"📄 Excel: {excel_path}")
    print(f"📂 Output: {output_dir}")
    print("="*70)
    
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return
    
    try:
        df = pd.read_excel(excel_path)
    except PermissionError:
        print(f"❌ Cannot open Excel. Please close it and try again.")
        return
    
    if df.empty:
        print("❌ Excel file is empty")
        return
    
    url_col, name_col = detect_excel_columns(df)
    if not url_col:
        print("❌ Could not find URL column (expected: 'url', 'course_url', 'link')")
        return
    if not name_col:
        print("❌ Could not find name column (expected: 'name', 'course_name', 'title')")
        return
    
    print(f"✅ Columns - URL: '{url_col}', Name: '{name_col}'")
    print(f"🧮 Total rows: {len(df)}")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="msedge",
            headless=False,
            args=['--disable-blink-features=AutomationControlled', '--no-sandbox']
        )
        
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        page.on("popup", lambda popup: popup.close())
        
        try:
            url_idx = df.columns.get_loc(url_col)
            name_idx = df.columns.get_loc(name_col)
            
            for idx, row in enumerate(df.itertuples(index=False, name=None)):
                base_url = str(row[url_idx]).strip()
                if not base_url or base_url.lower() == "nan":
                    print(f"\n[Row {idx+1}] ⚠️ Skipping empty URL")
                    continue
                
                name_value = row[name_idx]
                custom_name = name_value if pd.notna(name_value) else None
                
                print("\n" + "="*70)
                print(f"▶️  Row {idx + 1}/{len(df)}")
                print(f"📍 URL: {base_url}")
                if custom_name:
                    print(f"🏷  Name: {custom_name}")
                print("="*70)
                
                # Load page
                try:
                    page.goto(base_url, wait_until="domcontentloaded")
                    wait(page, 1000)
                except Exception as e:
                    print(f"❌ Navigation failed: {e}")
                    continue
                
                # Set up automatic ad blocker FIRST
                setup_ad_blocker(page)
                block_unwanted_elements(page)
                wait(page, 1000)  # Let initial ads get blocked
                
                # Process page
                process_about_section(page, base_url)
                process_modules_section(page, base_url)
                scroll_to_bottom(page)
                prepare_for_pdf(page)
                
                # Generate PDF
                pdf_file = generate_pdf(page, base_url, output_dir, custom_name)
                
                if pdf_file:
                    print("\n🎉 SUCCESS!")
                    print(f"📄 {pdf_file}")
                
                wait(page, 1000)
        
        except Exception as e:
            print(f"\n❌ Critical error: {str(e)}")
            import traceback
            traceback.print_exc()
        
        finally:
            context.close()
            browser.close()
            print("\n✅ Browser closed")


if __name__ == "__main__":
    main()