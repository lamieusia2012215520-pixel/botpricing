import re

def fix():
    with open(r"c:\He_thong_Bot\bot_msc.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Add MSCAccessDeniedException at the top
    if "class MSCAccessDeniedException" not in content:
        content = content.replace("import time\n", "import time\n\nclass MSCAccessDeniedException(Exception):\n    pass\n\n")

    # Update wait_for_search_results
    old_wait = """def wait_for_search_results(driver, timeout=60) -> str:
    print("Waiting for results...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                "iqsearchresult" in d.current_url
                or d.execute_script(\"""
                    var shadow = document.querySelector('mymsc-instantquote-app')?.shadowRoot;
                    if (!shadow) return false;
                    var h1 = shadow.querySelector('h1');
                    return h1 && h1.textContent.includes('No rates found');
                \""")
            )
        )
    except TimeoutException:
        print("Timeout waiting for results")
        return 'timeout'

    if "iqsearchresult" in driver.current_url:
        print("Results found")
        return 'has_results'
    else:
        print("No rates found")
        return 'no_results'"""
        
    new_wait = """def wait_for_search_results(driver, timeout=60) -> str:
    print("Waiting for results...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                "iqsearchresult" in d.current_url
                or "You do not have access to use this functionality" in d.page_source
                or d.execute_script(\"""
                    var shadow = document.querySelector('mymsc-instantquote-app')?.shadowRoot;
                    if (!shadow) return false;
                    var h1 = shadow.querySelector('h1');
                    return h1 && h1.textContent.includes('No rates found');
                \""")
            )
        )
    except TimeoutException:
        print("Timeout waiting for results")
        return 'timeout'

    if "You do not have access to use this functionality" in driver.page_source:
        raise MSCAccessDeniedException("Access denied during search")

    if "iqsearchresult" in driver.current_url:
        print("Results found")
        return 'has_results'
    else:
        print("No rates found")
        return 'no_results'"""
    
    content = content.replace(old_wait, new_wait)
    
    # Update main loop
    old_loop = """        try:
            try:
                driver.get(MYMSC_QUOTE_URL)
            except Exception as e:
                print(f"Page load timeout/error: {e}")
                try:
                    driver.execute_script("window.stop();")
                except:
                    pass
            wait_page_load(driver)

            if "instantquote" not in driver.current_url or "errorMessage" in driver.current_url:
                if not login_and_go_to_instant_quote(driver):
                    print(f"Login error at row {row_idx}")
                    continue

            final_data = solve_msc_quote(driver, pol, pod, country)

            if final_data:
                conts = final_data['containers']
                if "20DV" in conts:
                    sheet[f"F{row_idx}"].value = conts['20DV']['total'].get('formula') or conts['20DV']['total'].get('USD', "-")
                    sheet[f"M{row_idx}"].value = conts['20DV']['remark'] 
                else:
                    sheet[f"F{row_idx}"].value = "-"

                sheet[f"G{row_idx}"].value = (conts['40DV']['total'].get('formula') or conts['40DV']['total'].get('USD', "-")) if "40DV" in conts else "-"
                sheet[f"H{row_idx}"].value = (conts['40HC']['total'].get('formula') or conts['40HC']['total'].get('USD', "-")) if "40HC" in conts else "-"

                sheet[f"I{row_idx}"].value = final_data['etd']
                sheet[f"J{row_idx}"].value = final_data['tt_display']
                sheet[f"K{row_idx}"].value = final_data['validity']
                sheet[f"N{row_idx}"].value = final_data['freetime']
                
                sheet[f"O{row_idx}"].value = final_data['full_vessels']
                sheet[f"O{row_idx}"].alignment = openpyxl.styles.Alignment(wrapText=True)
                sheet[f"P{row_idx}"].value = final_data['ts_ports']

                print(f"Saved row {row_idx}")
            else:
                sheet[f"F{row_idx}"].value = "-"
                sheet[f"G{row_idx}"].value = "-"
                sheet[f"H{row_idx}"].value = "-"
                sheet[f"O{row_idx}"].value = "-"
                sheet[f"P{row_idx}"].value = "-"
                print(f"No rates for row {row_idx}, filled '-'")

        except Exception as e:
            print(f"[ERROR] Row {row_idx} thất bại: {e}")
            sheet[f"F{row_idx}"].value = "-"
            sheet[f"G{row_idx}"].value = "-"
            sheet[f"H{row_idx}"].value = "-"
            sheet[f"O{row_idx}"].value = "-"
            sheet[f"P{row_idx}"].value = "-"
            print(f"Filled '-' for row {row_idx} due to error")
            # Thử reconnect driver nếu tab bị đóng
            try:
                driver.current_url
            except:
                print("Driver lost connection, trying to reconnect...")
                try:
                    driver = connect_driver()
                except:
                    print("Reconnect failed, skipping remaining rows")
                    break"""
                    
    new_loop = """        while True:
            try:
                try:
                    driver.get(MYMSC_QUOTE_URL)
                except Exception as e:
                    print(f"Page load timeout/error: {e}")
                    try:
                        driver.execute_script("window.stop();")
                    except:
                        pass
                wait_page_load(driver)
                
                if "You do not have access to use this functionality" in driver.page_source:
                    raise MSCAccessDeniedException("Access denied on load")

                if "instantquote" not in driver.current_url or "errorMessage" in driver.current_url:
                    if not login_and_go_to_instant_quote(driver):
                        print(f"Login error at row {row_idx}")
                        raise Exception("Login error")

                final_data = solve_msc_quote(driver, pol, pod, country)

                if final_data:
                    conts = final_data['containers']
                    if "20DV" in conts:
                        sheet[f"F{row_idx}"].value = conts['20DV']['total'].get('formula') or conts['20DV']['total'].get('USD', "-")
                        sheet[f"M{row_idx}"].value = conts['20DV']['remark'] 
                    else:
                        sheet[f"F{row_idx}"].value = "-"

                    sheet[f"G{row_idx}"].value = (conts['40DV']['total'].get('formula') or conts['40DV']['total'].get('USD', "-")) if "40DV" in conts else "-"
                    sheet[f"H{row_idx}"].value = (conts['40HC']['total'].get('formula') or conts['40HC']['total'].get('USD', "-")) if "40HC" in conts else "-"

                    sheet[f"I{row_idx}"].value = final_data['etd']
                    sheet[f"J{row_idx}"].value = final_data['tt_display']
                    sheet[f"K{row_idx}"].value = final_data['validity']
                    sheet[f"N{row_idx}"].value = final_data['freetime']
                    
                    sheet[f"O{row_idx}"].value = final_data['full_vessels']
                    sheet[f"O{row_idx}"].alignment = openpyxl.styles.Alignment(wrapText=True)
                    sheet[f"P{row_idx}"].value = final_data['ts_ports']

                    print(f"Saved row {row_idx}")
                else:
                    sheet[f"F{row_idx}"].value = "-"
                    sheet[f"G{row_idx}"].value = "-"
                    sheet[f"H{row_idx}"].value = "-"
                    sheet[f"O{row_idx}"].value = "-"
                    sheet[f"P{row_idx}"].value = "-"
                    print(f"No rates for row {row_idx}, filled '-'")
                
                break # Success or handled properly, break the retry loop

            except MSCAccessDeniedException as e:
                print(f"ACCESS DENIED ERROR: {e}. Logging out and restarting row {row_idx}...")
                driver.delete_all_cookies()
                driver.get(MYMSC_QUOTE_URL)
                time.sleep(3)
                continue # Retry the loop for the same row
                
            except Exception as e:
                print(f"[ERROR] Row {row_idx} thất bại: {e}")
                sheet[f"F{row_idx}"].value = "-"
                sheet[f"G{row_idx}"].value = "-"
                sheet[f"H{row_idx}"].value = "-"
                sheet[f"O{row_idx}"].value = "-"
                sheet[f"P{row_idx}"].value = "-"
                print(f"Filled '-' for row {row_idx} due to error")
                # Thử reconnect driver nếu tab bị đóng
                try:
                    driver.current_url
                except:
                    print("Driver lost connection, trying to reconnect...")
                    try:
                        driver = connect_driver()
                    except:
                        print("Reconnect failed, skipping remaining rows")
                        break
                break # Break retry loop on normal error"""
                
    content = content.replace(old_loop, new_loop)
    
    with open(r"c:\He_thong_Bot\bot_msc.py", "w", encoding="utf-8") as f:
        f.write(content)

fix()
