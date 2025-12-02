from playwright.sync_api import sync_playwright
import re
import time
import os
import sys

import pandas as pd

# Avoid UnicodeEncodeError on Windows consoles when printing emoji/special chars
try:
    sys.stdout.reconfigure(errors="ignore")
except Exception:
    pass


def wait(page, ms: int = 500):
    """Small wrapper around Playwright timeout to keep calls consistent."""
    page.wait_for_timeout(ms)


def safe_click(page, locator, *, timeout: int = 2000, scroll: bool = True, force: bool = False) -> bool:
    """Click a locator safely without raising, returns True on success."""
    try:
        if locator is None:
            return False
        if not locator.is_visible(timeout=timeout):
            return False
        if scroll:
            locator.scroll_into_view_if_needed()
            wait(page, 300)
        locator.click(timeout=timeout, force=force)
        wait(page, 300)
        return True
    except Exception:
        return False


def clean_ads(page, times: int = 1, delay_ms: int = 400):
    """Run the aggressive popup cleaner multiple times with a delay."""
    for _ in range(times):
        close_ads_and_popups(page)


def close_ads_and_popups(page):
    """Aggressively close all ads, popups, and overlays including Black Friday ads"""
    try:
        # Press Escape key a few times to dismiss native dialogs
        for _ in range(3):
            page.keyboard.press("Escape")
            wait(page, 150)
        
        # First: Try to find and click close buttons with Playwright
        try:
            # Black Friday specific close buttons
            black_friday_close = page.locator(
                'button[aria-label*="Close"], '
                'button[data-testid*="close"], '
                '[class*="modal"] button:has-text("×"), '
                '[role="dialog"] button[aria-label*="Close"]'
            ).all()

            # Limit to first few to avoid clicking FAQ dialogs deeply nested
            for btn in black_friday_close[:5]:
                safe_click(page, btn, timeout=1000, force=True)
        except Exception:
            pass
        
        # Execute JavaScript to remove popups and ads
        page.evaluate("""
            () => {
                // Remove Black Friday / promotional overlays FIRST
                const blackFridaySelectors = [
                    '[class*="black-friday"]', '[class*="Black-Friday"]',
                    '[class*="blackfriday"]', '[class*="BlackFriday"]',
                    '[id*="black-friday"]', '[id*="BlackFriday"]',
                    '[class*="cyber-monday"]', '[class*="CyberMonday"]',
                    '[class*="promotion"]', '[class*="Promotion"]',
                    '[class*="promo"]', '[class*="Promo"]',
                    '[class*="sale-modal"]', '[class*="SaleModal"]',
                    '[class*="discount-modal"]', '[class*="DiscountModal"]',
                    '[data-track*="promo"]', '[data-track*="sale"]',
                    '[data-track*="black-friday"]'
                ];
                
                blackFridaySelectors.forEach(selector => {
                    try {
                        document.querySelectorAll(selector).forEach(el => {
                            el.remove();
                        });
                    } catch(e) {}
                });
                
                // Remove all modal dialogs
                const dialogs = document.querySelectorAll(
                    '[role="dialog"], [role="alertdialog"], ' +
                    '[class*="modal"], [class*="Modal"], ' +
                    '[id*="modal"], [id*="Modal"], ' +
                    '[class*="popup"], [class*="Popup"]'
                );
                dialogs.forEach(el => {
                    // Don't remove if it's part of FAQ
                    const text = el.textContent || '';
                    if (!text.toLowerCase().includes('frequently asked') && 
                        !text.toLowerCase().includes('faq')) {
                        el.remove();
                    }
                });
                
                // Click close buttons but NOT FAQ buttons
                const closeButtons = document.querySelectorAll(
                    '[data-testid*="close"]:not([data-testid*="faq"]), ' +
                    '[aria-label*="Close"]:not([aria-label*="FAQ"]):not([aria-label*="frequently"]), ' +
                    'button[class*="close"]:not([class*="faq"])'
                );
                closeButtons.forEach(btn => {
                    try:
                        const parent = btn.closest('[role="dialog"], [class*="modal"]');
                        if (parent) {
                            const parentText = parent.textContent || '';
                            if (!parentText.toLowerCase().includes('frequently asked') &&
                                !parentText.toLowerCase().includes('faq')) {
                                btn.click();
                            }
                        }
                    } catch(e) {}
                });
                
                // Remove all high z-index overlays and backdrops
                const allElements = document.querySelectorAll('*');
                allElements.forEach(el => {
                    const style = window.getComputedStyle(el);
                    if (style.position === 'fixed' || style.position === 'absolute') {
                        const zIndex = parseInt(style.zIndex);
                        // High z-index elements are likely popups
                        if (zIndex > 999 || (zIndex > 100 && (
                            el.className.toLowerCase().includes('overlay') ||
                            el.className.toLowerCase().includes('backdrop') ||
                            el.className.toLowerCase().includes('modal') ||
                            el.className.toLowerCase().includes('popup')
                        ))) {
                            // Don't remove FAQ elements
                            const elText = el.textContent || '';
                            const elClass = el.className || '';
                            if (!elText.toLowerCase().includes('frequently asked') &&
                                !elClass.toLowerCase().includes('faq')) {
                                el.remove();
                            }
                        }
                    }
                });
                
                // Accept cookie consent if present
                const cookieBtn = document.querySelector(
                    '#onetrust-accept-btn-handler, ' +
                    '[id*="cookie"] button, ' +
                    'button[id*="accept-cookie"]'
                );
                if (cookieBtn) cookieBtn.click();
                
                // Remove any ad containers
                const ads = document.querySelectorAll(
                    '[class*="ad-"], [class*="ad_"], ' +
                    '[id*="ad-"], [id*="ad_"], ' +
                    '[class*="advertisement"], [class*="Advertisement"], ' +
                    'iframe[src*="ads"], iframe[src*="doubleclick"]'
                );
                ads.forEach(ad => ad.remove());
                
                // Remove notification banners (but not FAQ)
                const notifications = document.querySelectorAll(
                    '[class*="notification"], [class*="Notification"], ' +
                    '[class*="banner"], [class*="Banner"]'
                );
                notifications.forEach(notif => {
                    const style = window.getComputedStyle(notif);
                    const notifText = notif.textContent || '';
                    if ((style.position === 'fixed' || style.position === 'sticky') &&
                        !notifText.toLowerCase().includes('frequently asked') &&
                        !notifText.toLowerCase().includes('faq')) {
                        notif.remove();
                    }
                });
                
                // Reset body overflow to prevent scroll lock
                document.body.style.overflow = 'visible';
                document.body.style.position = 'static';
                document.documentElement.style.overflow = 'visible';
            }
        """)
        return True
        
    except Exception as e:
        return False


def close_initial_popups(page):
    """Close only the Recommended Experience popup - ONCE at start"""
    try:
        print("  🔒 Closing initial popups and ads...")
        
        # Close ads and popups aggressively - multiple times
        clean_ads(page, times=5, delay_ms=600)
        
        # Look for the Recommended Experience popup specifically
        try:
            popup = page.locator("text=Recommended experience").first
            if popup.is_visible(timeout=3000):
                # Click OK button
                ok_btn = page.locator("button:has-text('OK')").first
                if safe_click(page, ok_btn, timeout=2000):
                    print("    ✓ Closed 'Recommended Experience' popup")
                    wait(page, 1000)
        except Exception:
            print("    ℹ️  No Recommended Experience popup found")
        
        # Block unwanted buttons permanently with CSS
        page.add_style_tag(content="""
            /* Block Explore button */
            button[data-testid*='explore'],
            button[aria-label*='Explore'],
            a[href*='/explore'],
            [data-track-component*='explore'] {
                pointer-events: none !important;
                opacity: 0.3 !important;
                display: none !important;
            }
            
            /* Block FAQ accordions - CRITICAL */
            button[data-e2e*='faq'],
            button[data-e2e*='FAQ'],
            button[aria-label*='frequently asked'],
            button[aria-label*='Frequently asked'],
            button[aria-label*='Frequently Asked'],
            [data-testid*='faq'],
            [data-testid*='FAQ'],
            div[class*='faq'] button[aria-expanded],
            section[class*='faq'] button[aria-expanded] {
                pointer-events: none !important;
                opacity: 0.3 !important;
                cursor: not-allowed !important;
            }
            
            /* Block difficulty level info button */
            button[aria-label*='Information about difficulty level'] {
                pointer-events: none !important;
                opacity: 0.3 !important;
            }
            
            /* Block Black Friday and promotional popups - AGGRESSIVE */
            [class*="black-friday"],
            [class*="Black-Friday"],
            [class*="blackfriday"],
            [class*="BlackFriday"],
            [id*="black-friday"],
            [id*="BlackFriday"],
            [class*="cyber-monday"],
            [class*="promotion-modal"],
            [class*="promo-modal"],
            [class*="sale-modal"],
            [data-track*="promo-modal"],
            [data-track*="black-friday"] {
                display: none !important;
                visibility: hidden !important;
                pointer-events: none !important;
                opacity: 0 !important;
                z-index: -9999 !important;
            }
        """)
        
        print("    ✓ Blocked unwanted buttons (Explore, FAQ, Difficulty info, Promo ads)")
        
        # Final cleanup after blocking
        clean_ads(page, times=3, delay_ms=500)
        
    except Exception as e:
        print(f"    ⚠️  Popup close warning: {str(e)[:50]}")


def scroll_and_wait(page, pixels=500):
    """Smooth scroll with wait and AGGRESSIVE ad cleanup"""
    page.evaluate(f"window.scrollBy({{top: {pixels}, behavior: 'smooth'}})")
    # Close any ads that appeared during scroll
    clean_ads(page, times=2, delay_ms=200)


def process_about_section(page, base_url):
    """Process About section - View skills FIRST, then Read more"""
    print("\n" + "="*70)
    print("📍 STEP 1: ABOUT SECTION")
    print("="*70)
    
    try:
        # Navigate to About
        page.goto(f"{base_url}#about", wait_until="load")
        
        # Close any ads that appeared - AGGRESSIVE
        clean_ads(page, times=3, delay_ms=500)
        
        # Scroll within About section first
        print("  📜 Initial scroll through About section...")
        for _ in range(2):
            scroll_and_wait(page, 50)
        
        # Close ads after scrolling
        clean_ads(page, times=1, delay_ms=50)
        
        # FIRST: Click "View all skills" button
        print("  🔍 STEP 1A: Looking for 'View all skills' button...")
        try:
            skills_btn = page.locator('button:has-text("View all skills")').first

            if skills_btn.is_visible(timeout=3000):
                # Close ads before clicking
                clean_ads(page, times=1, delay_ms=300)

                if safe_click(page, skills_btn, timeout=1000):
                    print("    ✅ Expanded 'View all skills'")
                    # Close ads immediately after expansion
                    clean_ads(page, times=2, delay_ms=500)
            else:
                print("    ℹ️  'View all skills' not found")
        except Exception as e:
            print(f"    ℹ️  'View all skills' not available: {str(e)[:40]}")
        
        # SECOND: Click Read more buttons
        print("  📖 STEP 1B: Clicking 'Read more' buttons...")
        click_read_more_buttons_in_section(page, "About")
        
        # Close ads after all clicks
        clean_ads(page, times=2, delay_ms=400)
        
        print("  ✅ About section complete")
        
    except Exception as e:
        print(f"  ❌ Error in About: {str(e)[:50]}")
    
    print("="*70)


def click_read_more_buttons_in_section(page, section_name=""):
    """Click ONLY valid Read more buttons - skip Explore/FAQ/Partner sections"""
    try:
        print(f"  📖 Looking for 'Read more' buttons in {section_name}...")
        
        # Find all Read more buttons
        all_buttons = page.locator('button:has-text("Read more")').all()
        
        if not all_buttons:
            print(f"    ℹ️  No 'Read more' buttons found")
            return
        
        print(f"    Found {len(all_buttons)} potential buttons, filtering...")
        
        clicked = 0
        for idx, btn in enumerate(all_buttons, 1):
            try:
                # Check if button is visible
                if not btn.is_visible(timeout=1000):
                    continue
                
                # Get context to filter out unwanted buttons
                aria_label = btn.get_attribute("aria-label") or ""
                
                # Skip unwanted buttons
                skip_keywords = [
                    "explore", "Explore", "EXPLORE",
                    "frequently asked", "FAQ", "faq",
                    "offered by", "partner", "Partner",
                    "Learn more about"
                ]
                
                should_skip = any(keyword.lower() in aria_label.lower() 
                                 for keyword in skip_keywords)
                
                if should_skip:
                    print(f"      ⊘ Skipped unwanted: {aria_label[:40]}")
                    continue
                
                # Valid button - click it
                if safe_click(page, btn, timeout=1500):
                    clicked += 1
                    print(f"      ✓ Clicked Read more {clicked}")
                
            except Exception as e:
                continue
        
        if clicked > 0:
            print(f"    ✅ Clicked {clicked} valid 'Read more' button(s)")
        else:
            print(f"    ℹ️  No valid buttons to click")
            
    except Exception as e:
        print(f"    ⚠️  Error: {str(e)[:50]}")


def process_modules_section(page, base_url):
    """Process Modules - expand ALL module accordions one by one (NOT FAQ)"""
    print("\n" + "="*70)
    print("📍 STEP 2: MODULES/COURSES SECTION")
    print("="*70)
    
    try:
        # Navigate to Modules
        page.goto(f"{base_url}#modules", wait_until="load")
        
        # Close any ads
        clean_ads(page, times=1, delay_ms=50    )
        
        print("  📦 Expanding module accordions sequentially (excluding FAQ)...")
        
        # Find all module accordion buttons with STRICT filtering
        all_accordions = page.locator('button[aria-expanded]').all()
        
        # Filter to ONLY module buttons, EXCLUDE FAQ
        module_buttons = []
        for btn in all_accordions:
            try:
                aria_label = (btn.get_attribute("aria-label") or "").lower()
                data_e2e = (btn.get_attribute("data-e2e") or "").lower()
                btn_class = (btn.get_attribute("class") or "").lower()
                btn_text = (btn.text_content() or "").lower()
                
                # Get parent section to check context
                parent_text = ""
                try:
                    parent = btn.locator('xpath=../..').first
                    parent_text = (parent.text_content() or "")[:200].lower()
                except:
                    pass
                
                # STRICT FAQ filtering - do not expand any FAQ / question accordions
                is_faq = (
                    'faq' in aria_label
                    or 'faq' in data_e2e
                    or 'faq' in btn_class
                    or 'faq' in btn_text
                    or 'frequently asked' in aria_label
                    or 'frequently asked' in btn_text
                    or 'frequently asked' in parent_text
                    or 'frequently asked questions' in parent_text
                    or 'questions' in btn_text
                )
                
                if not is_faq:
                    module_buttons.append(btn)
                else:
                    print(f"    ⊘ Filtered out FAQ button: {aria_label[:40] or data_e2e[:40]}")
            except:
                continue
        
        if not module_buttons:
            print("    ℹ️  No module accordions found, trying Courses section...")
            page.goto(f"{base_url}#courses", wait_until="load")
            clean_ads(page, times=1, delay_ms=50)
            
            # Try again with filtering
            all_accordions = page.locator('button[aria-expanded]').all()
            module_buttons = [
                btn for btn in all_accordions
                if 'faq' not in (btn.get_attribute('aria-label') or '').lower()
                and 'frequently' not in (btn.get_attribute('aria-label') or '').lower()
            ]
        
        if module_buttons:
            total = len(module_buttons)
            print(f"  📊 Found {total} valid module(s) to expand (FAQ excluded)")
            
            for idx, btn in enumerate(module_buttons, 1):
                try:
                    # Close ads before each click
                    clean_ads(page, times=1, delay_ms=250)
                    
                    # Check if already expanded
                    is_expanded = btn.get_attribute("aria-expanded")
                    
                    if is_expanded == "true":
                        print(f"    [{idx}/{total}] Already expanded, skipping")
                        continue
                    
                    # Double-check it's not FAQ before clicking
                    btn_label = (btn.get_attribute("aria-label") or "").lower()
                    btn_text_click = (btn.text_content() or "").lower()
                    if (
                        'faq' in btn_label
                        or 'frequently' in btn_label
                        or 'faq' in btn_text_click
                        or 'frequently' in btn_text_click
                        or 'question' in btn_text_click
                    ):
                        print(f"    [{idx}/{total}] ⊘ Skipped FAQ button")
                        continue
                    
                    # Scroll to module
                    print(f"    [{idx}/{total}] Scrolling to module...")
                    btn.scroll_into_view_if_needed()


                    # Click to expand
                    print(f"    [{idx}/{total}] Clicking to expand...")
                    if safe_click(page, btn, timeout=1800, scroll=False):
                        print(f"    [{idx}/{total}] ✅ Expanded")

            
                    clean_ads(page, times=1, delay_ms=300)
                    
                except Exception as e:
                    print(f"    [{idx}/{total}] ⚠️  Error: {str(e)[:40]}")
            
            print("  ✅ All modules processed")
        else:
            print("    ℹ️  No valid modules found")
        
        # Scroll through expanded modules
        print("  📜 Scrolling through expanded content...")
        for _ in range(3):
            scroll_and_wait(page, 450)
        
        # Click Read more in modules
        click_read_more_buttons_in_section(page, "Modules")
        
        # Final ad cleanup
        clean_ads(page, times=1, delay_ms=400)
        
        print("  ✅ Modules section complete")
        
    except Exception as e:
        print(f"  ❌ Error in Modules: {str(e)[:50]}")
    
    print("="*70)


def progressive_scroll_to_bottom(page):
    """Scroll to absolute bottom to load all lazy content"""
    print("\n" + "="*70)
    print("📍 STEP 3: SCROLL TO BOTTOM")
    print("="*70)
    print("  📜 Scrolling to load all remaining content...")
    
    try:
        last_height = page.evaluate("document.body.scrollHeight")
        scroll_count = 0
        max_scrolls = 50
        
        while scroll_count < max_scrolls:
            # Scroll by viewport height
            page.evaluate("window.scrollBy(0, window.innerHeight * 0.8)")
            
            # Close ads periodically
            if scroll_count % 5 == 0:
                clean_ads(page, times=1, delay_ms=300)
            
            scroll_count += 1
            
            # Check if at bottom
            new_height = page.evaluate("document.body.scrollHeight")
            current_pos = page.evaluate("window.pageYOffset + window.innerHeight")
            
            if scroll_count % 10 == 0:
                print(f"    → Scrolled {scroll_count} times...")
            
            # Check if reached bottom
            if current_pos >= new_height - 100:
                print(f"    ✅ Reached bottom after {scroll_count} scrolls")
                break
            
            # Check if no new content
            if new_height == last_height:
                break
            
            last_height = new_height
        
        # Ensure absolute bottom
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        
        # Final ad cleanup
        clean_ads(page, times=1, delay_ms=100)
        
        print("  ✅ Scroll complete")
        
    except Exception as e:
        print(f"  ⚠️  Scroll warning: {str(e)[:50]}")
    
    print("="*70)


def prepare_page_for_pdf(page):
    """Final preparation - expand all, remove overlays"""
    print("\n" + "="*70)
    print("📍 STEP 4: PREPARE FOR PDF")
    print("="*70)
    
    try:
        print("  🔧 Removing overlays and expanding content...")
        
        # Final aggressive ad/popup cleanup
        clean_ads(page, times=3, delay_ms=100)
        
        # Scroll to top
        page.evaluate("window.scrollTo({top: 0, behavior: 'smooth'})")
        
        # Execute cleanup script
        page.evaluate("""
            () => {
                // Remove all dialogs and modals
                document.querySelectorAll('[role="dialog"], [role="alertdialog"]').forEach(el => el.remove());
                
                // Remove fixed/sticky elements
                document.querySelectorAll('header, nav, footer').forEach(el => {
                    const style = window.getComputedStyle(el);
                    if (style.position === 'fixed' || style.position === 'sticky') {
                        el.style.display = 'none';
                    }
                });
                
                // Remove overlays
                document.querySelectorAll('[class*="overlay"], [class*="backdrop"]').forEach(el => el.remove());
                
                // Expand all collapsed content
                document.querySelectorAll('[aria-expanded="false"]').forEach(btn => {
                    btn.setAttribute('aria-expanded', 'true');
                });
                
                // Show all hidden content
                document.querySelectorAll('[aria-hidden="true"]').forEach(el => {
                    el.setAttribute('aria-hidden', 'false');
                    el.style.display = 'block';
                    el.style.visibility = 'visible';
                });
                
                // Ensure main content visible
                document.querySelectorAll('main, article, section').forEach(el => {
                    el.style.display = 'block';
                    el.style.visibility = 'visible';
                    el.style.overflow = 'visible';
                    el.style.maxHeight = 'none';
                });
                
                // Reset body
                document.body.style.overflow = 'visible';
                document.body.style.height = 'auto';
            }
        """)
                
        # One final scroll to ensure everything loaded
        print("  📜 Final scroll to ensure all content loaded...")
        page.evaluate("""
            () => {
                let pos = 0;
                const height = document.body.scrollHeight;
                while (pos < height) {
                    window.scrollBy(0, 500);
                    pos += 500;
                }
                window.scrollTo(0, 0);
            }
        """)
        
        print("  ✅ Page prepared")
        
    except Exception as e:
        print(f"  ⚠️  Preparation warning: {str(e)[:50]}")
    
    print("="*70)


def generate_pdf(page, base_url, output_dir=".", custom_name=None):
    """Generate PDF with selectable text.

    - `output_dir`: directory where the PDF will be saved.
    - `custom_name`: optional name (from Excel) to use in the filename.
    """
    print("\n" + "="*70)
    print("📍 STEP 5: GENERATE PDF")
    print("="*70)
    
    try:
        # Emulate print media
        print("  🖨️  Setting print mode...")
        page.emulate_media(media="print")
        
        # Extract course name (fallback if no custom name provided)
        course_name = "Coursera_Course"
        try:
            title = page.locator('h1').first.text_content().strip()
            course_name = re.sub(r'[<>:"/\\|?*]', '_', title)
            print(f"  📖 Course title: {course_name}")
        except:
            print("  ℹ️  Using default course title in filename")
        
        # Create filename using optional custom name from Excel
        url_slug = base_url.split("/")[-1].split("?")[0]
        if custom_name:
            safe_custom = re.sub(r'[<>:"/\\|?*]', '_', str(custom_name).strip())
            filename = f"{safe_custom}_{url_slug}.pdf"
        else:
            filename = f"{course_name}_{url_slug}.pdf"

        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)[:200]

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        full_path = os.path.join(output_dir, filename)
        
        print(f"  💾 Filename: {full_path}")
        print("Preparing page for PDF...")

        # --- FIX BLANK PDF (Ensure all content is visible) ---

        # 1) Force proper viewport for PDF
        page.set_viewport_size({"width": 1200, "height": 800})

        # 2) Wait for main content to be visible
        page.wait_for_selector('main, [data-testid*="main"], article, .content', timeout=10000)

        # 3) Scroll entire page to load lazy elements
        for _ in range(15):
            page.evaluate("window.scrollBy(0, window.innerHeight)")
            wait(page, 500)

        # 4) Remove fixed headers/overlays that ruin PDF rendering
        page.evaluate("""
        document.querySelectorAll("*").forEach(el => {
            const style = getComputedStyle(el);
            if (style.position === "fixed" || style.position === "sticky") {
                el.style.position = "static";
                el.style.top = "auto";
                el.style.zIndex = "0";
            }
        });

        // Ensure main content is visible
        document.querySelectorAll('main, article, section, .content').forEach(el => {
            el.style.display = 'block';
            el.style.visibility = 'visible';
            el.style.opacity = '1';
        });

        // Remove any remaining overlays
        document.querySelectorAll('[role="dialog"], [role="alertdialog"], .modal, .overlay').forEach(el => el.remove());
        """)

        # 5) Final wait for rendering
        wait(page, 3000)

        print("Saving PDF now...")

        page.pdf(
            path=full_path,
            format="A4",
            print_background=True,
            prefer_css_page_size=False,
            margin={
                "top": "0.4in",
                "bottom": "0.4in",
                "left": "0.5in",
                "right": "0.5in"
            },
            scale=0.90
        )
        
        print(f"\n  ✅ PDF SAVED: {full_path}")
        print("="*70)
        return full_path
        
    
    except Exception as e:
        print(f"  ❌ PDF generation failed: {str(e)}")
        import traceback
        traceback.print_exc()
        print("="*70)
        return None


def _detect_excel_columns(df):
    """Detect URL and name columns in the Excel file with flexible matching."""
    def _find_col(possible_names):
        lower_map = {str(c).strip().lower(): c for c in df.columns}
        for name in possible_names:
            if name in lower_map:
                return lower_map[name]
        return None

    # Allow a few common header variants (case-insensitive, with/without spaces)
    url_col = _find_col([
        "url",
        "course_url",
        "course url",
        "link",
        "course_link",
        "course link",
        "coursera_url",
        "coursera url",
    ])
    name_col = _find_col([
        "name",
        "course_name",
        "course name",
        "title",
        "course_title",
        "course title",
        "coursera course name",
    ])
    return url_col, name_col


def main():
    """Main execution flow: read URLs from Excel and generate PDFs in batch."""

    excel_path = "courses.xlsx"
    output_dir = "pdfs"

    print("\n" + "="*70)
    print("🚀 COURSERA SCRAPER - BATCH MODE FROM EXCEL")
    print("="*70)
    print(f"📄 Excel source: {excel_path}")
    print(f"📂 Output folder: {output_dir}")
    print("="*70)

    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return

    # Read Excel with URLs and names
    try:
        df = pd.read_excel(excel_path)
    except PermissionError as e:
        print(f"❌ Cannot open '{excel_path}': {e}")
        print("   Please close the Excel file (or any program using it) and run the script again.")
        return
    if df.empty:
        print("❌ Excel file has no rows.")
        return

    url_col, name_col = _detect_excel_columns(df)
    if not url_col:
        raise ValueError(
            "Could not detect URL column in Excel. "
            "Please name it one of: 'url', 'course_url', 'link'."
        )
    if not name_col:
        raise ValueError(
            "Could not detect course-name column in Excel. "
            "Please name it one of: 'name', 'course_name', 'course name', 'title'."
        )

    print(f"✅ Detected columns - URL: '{url_col}', Name: '{name_col}'")
    print(f"🧮 Total rows: {len(df)}")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge",
            headless=False,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-web-security',
                '--no-sandbox'
            ]
        )

        # Use a browser context; popups are closed per-page to avoid closing the main tab.
        context = browser.new_context(viewport={"width": 1920, "height": 1080})
        page = context.new_page()
        page.on("popup", lambda popup: popup.close())
        
        try:
            # Pre-compute column indexes for faster access in the loop
            url_idx = df.columns.get_loc(url_col)
            name_idx = df.columns.get_loc(name_col)

            for idx, row in enumerate(df.itertuples(index=False, name=None)):
                base_url = str(row[url_idx]).strip()
                if not isinstance(base_url, str) or not base_url or base_url.lower() == "nan":
                    print(f"\n[Row {idx}] ⚠️ Skipping row with empty URL")
                    continue

                name_value = row[name_idx]
                custom_name = name_value if pd.notna(name_value) else None

                print("\n" + "="*70)
                print(f"▶️  Processing row {idx + 1}/{len(df)}")
                print(f"📍 URL: {base_url}")
                if custom_name:
                    print(f"🏷  Name: {custom_name}")
                print("="*70)

                # Initial page load
                print("\n⏳ Loading page...")
                try:
                    page.goto(base_url, wait_until="domcontentloaded")
                except Exception as e:
                    print(f"\n❌ Navigation failed for URL '{base_url}': {e}")
                    print("   Skipping this row and continuing with the next one.")
                    continue
                page.wait_for_timeout(3000)
                
                # Close initial popups and block unwanted buttons
                close_initial_popups(page)
                page.wait_for_timeout(1000)
                
                # Additional aggressive cleanup after initial load
                print("  🧹 Additional cleanup after page load...")
                for i in range(3):
                    close_ads_and_popups(page)
                    page.wait_for_timeout(800)
                
                # Sequential flow
                process_about_section(page, base_url)
                page.wait_for_timeout(500)
                
                process_modules_section(page, base_url)
                page.wait_for_timeout(600)
                
                progressive_scroll_to_bottom(page)
                page.wait_for_timeout(600)
                
                prepare_page_for_pdf(page)
                
                pdf_file = generate_pdf(
                    page,
                    base_url,
                    output_dir=output_dir,
                    custom_name=custom_name,
                )
                
                if pdf_file:
                    print("\n" + "="*70)
                    print("🎉 SUCCESS!")
                    print(f"📄 {pdf_file}")
                    print("="*70)
                
                page.wait_for_timeout(3000)
            
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


