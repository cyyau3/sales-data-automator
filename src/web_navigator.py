# src/web_navigator.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from logger_config import logger
from datetime import datetime
import pandas as pd
import numpy as np
import os
import time
import openpyxl
import traceback
from urls import URLConfig


class WebNavigator:
    def __init__(self, timeout=30):
        self.timeout = timeout

        # Initialize driver with options
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--lang=zh-TW')

        # Add incognito mode
        options.add_argument('--incognito')
        options.add_argument('--disable-cache')
        options.add_argument('--disable-application-cache')
        options.add_argument('--disable-offline-load-stale-cache')
        
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, self.timeout) 
        
    def login(self, username, password):
        """Login to UCD website"""
        try:
            # Navigate to main page
            logger.info("Navigating to main page...")
            self.driver.get(URLConfig.BASE_URL)
            
            # Wait for and click the login link in nav bar
            logger.info("Looking for login link...")
            login_link = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//a[contains(@href, '{URLConfig.LOGIN_PATH}')]"))
            )
            logger.info("Clicking login link...")
            login_link.click()
            
            # Wait for login form
            logger.info("Waiting for login form...")
            
            # Find form elements using their exact IDs
            username_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "user_name"))
            )
            password_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "user_password"))
            )
            
            # Clear and fill in credentials
            username_field.clear()
            logger.info("Entering username...")
            username_field.send_keys(username)
            
            password_field.clear()
            logger.info("Entering password...")
            password_field.send_keys(password)
            
            # Click the login button using its exact attributes
            login_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='B1'][type='submit'][value='確認登入']"))
            )
            logger.info("Clicking submit button...")
            login_button.click()
            
            # Wait for redirect to member page
            logger.info("Waiting for redirect to member page...")
            self.wait.until(
                EC.url_to_be(URLConfig.get_full_url(URLConfig.MEMBER_PATH))
            )
            
            logger.info("Successfully logged in")
                
        except TimeoutException as e:
            logger.error("Login form interaction failed - timeout")
            self.save_screenshot("login_failure")
            raise
        except Exception as e:
            # Mask the username in the error message
            logger.error(f"Login process failed for user: {username[:2]}***")
            self.save_screenshot("login_failure")
            raise

    def return_to_index(self):
        """Return to the member index page"""
        try:
            # Navigate directly to index page
            self.driver.get(URLConfig.get_full_url(URLConfig.MEMBER_PATH))
            
            # Additional wait for nav menu
            self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "nav"))
            )
            
            logger.info("Successfully returned to index page")
            
        except Exception as e:
            logger.error(f"Failed to return to index: {str(e)}")
            self.save_screenshot("return_to_index_error")
            raise

    def filter_month_generator(self, year=None, month=None):
        """Generate appropriate month and year for filtering"""
        try:
            current_date = datetime.now()
            
            # If no year/month provided, get previous month
            if year is None and month is None:
                if current_date.month == 1:
                    year = current_date.year - 1
                    month = 12
                else:
                    year = current_date.year
                    month = current_date.month - 1
            
            # Validate month
            if month < 1 or month > 12:
                raise ValueError(f"Invalid month value: {month}")
                
            # Return both separate and combined formats
            return {
                'year': year,
                'month': str(month).zfill(2),
                'combined': f"{year}{str(month).zfill(2)}"  # e.g., "202410"
            }
            
        except Exception as e:
            logger.error(f"Failed to generate filter dates: {str(e)}")
            raise

    def navigate_to_inventory(self):
        """Navigate to the inventory page after login"""
        try:
            # Use self.wait instead of creating new WebDriverWait
            nav_div = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "nav"))
            )
            
            inventory_link = self.driver.find_element(By.XPATH, "//a[contains(text(), '[606030] 庫存明細')]")
            inventory_link.click()
            
            # Use self.wait instead of creating new WebDriverWait
            self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "dataGrid"))
            )
            
            logger.info("Successfully navigated to inventory page")
            
        except TimeoutException:
            logger.error("Timeout waiting for inventory page elements")
            raise
        except Exception as e:
            logger.error(f"Failed to navigate to inventory page: {str(e)}")
            raise

    def extract_inventory_table(self):
        """Extract data from the inventory table"""
        try:
            # Wait for table to be present
            table = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "dataGrid"))
            )
            
            # Initialize lists to store data
            data = []
            
            # Get headers
            headers = []
            header_row = table.find_element(By.TAG_NAME, "thead").find_element(By.TAG_NAME, "tr")
            for th in header_row.find_elements(By.TAG_NAME, "th"):
                headers.append(th.text.strip())
            
            # Get table body rows
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            
            # Extract data from each row
            for row in rows:
                row_data = []
                for cell in row.find_elements(By.TAG_NAME, "td"):
                    row_data.append(cell.text.strip())
                if row_data:  # Only add non-empty rows
                    data.append(row_data)
            
            # Extract footer data
            tfoot = table.find_element(By.TAG_NAME, "tfoot")
            footer_row = tfoot.find_element(By.TAG_NAME, "tr")
            
            footer_data = []
            # Get text from pdtCode (總計)
            footer_data.append(footer_row.find_element(By.CLASS_NAME, "pdtCode").text.strip())
            
            # Get text and number from pdtName (共19種產品)
            pdt_name_text = footer_row.find_element(By.CLASS_NAME, "pdtName").text.strip()
            footer_data.append(pdt_name_text)
            
            # Get number from stockQuantity
            stock_qty = footer_row.find_element(By.CLASS_NAME, "stockQuantity").text.strip()
            footer_data.append(stock_qty)
            
            # Get number from stockAmount
            stock_amount = footer_row.find_element(By.CLASS_NAME, "stockAmount").text.strip()
            footer_data.append(stock_amount)
            
            # Add empty values for remaining columns (定價, 序號, 安全存量)
            footer_data.extend(['', '', ''])
            
            # Add footer data to main data
            data.append(footer_data)


            # Create DataFrame
            df = pd.DataFrame(data, columns=headers)
            
            # Convert numeric columns
            numeric_columns = ['庫存量', '庫存額', '定價', '安全存量']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            logger.info(f"Successfully extracted {len(df)} inventory records")
            return df
            
        except Exception as e:
            logger.error(f"Failed to extract inventory table: {str(e)}")
            raise

    def navigate_to_monthly_supply(self):
            """Navigate to the monthly supply report page"""
            try:
                nav_div = self.wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nav"))
                )
                
                # Update this XPath according to the actual menu item text
                supply_link = self.driver.find_element(By.XPATH, "//a[contains(text(), '[606031] 庫存月報表')]")
                supply_link.click()
                
                # Wait for the filter form to be present
                self.wait.until(
                    EC.presence_of_element_located((By.XPATH, "//form[@action='supp_summary.jsp']"))
                )
                
                logger.info("Successfully navigated to monthly supply page")
                
            except TimeoutException:
                logger.error("Timeout waiting for monthly supply page elements")
                raise
            except Exception as e:
                logger.error(f"Failed to navigate to monthly supply page: {str(e)}")
                raise

    def set_monthly_supply_filter(self, year=None, month=None):
        """Set filter for monthly supply report"""
        try:
            # Get date values
            date_values = self.filter_month_generator(year, month)
            
            # Select year
            year_select = self.wait.until(
                EC.presence_of_element_located((By.NAME, "p_year"))
            )
            year_dropdown = Select(year_select)
            year_dropdown.select_by_value(str(date_values['year']))

            # Select month
            month_select = self.wait.until(
                EC.presence_of_element_located((By.NAME, "p_period"))
            )
            month_dropdown = Select(month_select)
            month_dropdown.select_by_value(date_values['month'])

            # Submit
            submit_button = self.wait.until(
                EC.element_to_be_clickable((By.NAME, "B1"))
            )
            submit_button.click()

            # Wait for results
            self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "sortable"))
            )
            
            logger.info(f"Successfully set filter for {date_values['year']}/{date_values['month']}")

        except Exception as e:
            logger.error(f"Failed to set monthly supply filter: {str(e)}")
            raise

    def extract_monthly_supply_table(self):
        """Extract data from the monthly supply table"""
        try:
            # Extract title from p element
            title = self.driver.find_element(By.XPATH, "//p[contains(text(), '庫存銷售月報表')]").text
            logger.debug(f"Found title: {title}")

            # Wait for main table to be present
            main_table = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "sortable"))
            )
            
            # Get the main table data
            table_html = main_table.get_attribute('outerHTML')
            tables = pd.read_html(table_html)
            df = tables[0]
            
            # Convert columns as before
            numeric_columns = ['定價', '存量', '存額', '月進量', '退量', '進淨量', 
                            '出量', '退量', '出淨量', '年進量', '退量', '進淨量', 
                            '出量', '退量', '出淨量']
            
            date_columns = ['發書日']
            string_columns = ['貨物代碼', '書名', '系列編號']
            
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(',', '')
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            for col in string_columns:
                if col in df.columns:
                    df[col] = df[col].fillna('').astype(str)

            try:
                # Try to find the summary row using a more specific XPath
                summary_rows = self.driver.find_elements(
                    By.XPATH, "//tr[td[contains(text(), '合計') or contains(text(), '合  計')]]"
                )
                
                if summary_rows:
                    logger.debug(f"Found {len(summary_rows)} potential summary rows")
                    summary_row = summary_rows[-1]  # Take the last one if multiple found
                    cells = summary_row.find_elements(By.TAG_NAME, "td")
                    
                    # Log the actual cell contents for debugging
                    cell_texts = [cell.text.strip() for cell in cells]
                    logger.debug(f"Summary row cell contents: {cell_texts}")
                    
                    # Create summary data dictionary with empty values for first columns
                    summary_data = {
                        '貨物代碼': '',
                        '書名': '',
                        '發書日': pd.NaT,
                        '定價': None,
                        '系列編號': '合計'  # Put '合計' in 系列編號 column
                    }
                    
                    # Remove the first cell that contains '合計' and process the remaining cells sequentially
                    remaining_cells = cells[1:]  # Skip the first cell with '合計'
                    
                    # Map the values to the correct columns in order
                    columns_order = ['存量', '存額', '月進量', '退量', '進淨量', 
                                   '出量', '退量.1', '出淨量', '年進量', '退量.2', 
                                   '進淨量.1', '出量.1', '退量.3', '出淨量.1']
                    
                    for i, col_name in enumerate(columns_order):
                        if i < len(remaining_cells):
                            value = remaining_cells[i].text.strip().replace(',', '')
                            try:
                                summary_data[col_name] = float(value) if value else 0.0
                            except ValueError:
                                logger.warning(f"Could not convert value for {col_name}: {value}")
                                summary_data[col_name] = 0.0
                        else:
                            summary_data[col_name] = 0.0
                    
                    # Add summary row to DataFrame
                    summary_df = pd.DataFrame([summary_data])
                    df = pd.concat([df, summary_df], ignore_index=True)
                    logger.debug(f"Added summary row: {summary_data}")
                else:
                    logger.warning("No summary row found")

            except Exception as e:
                logger.warning(f"Failed to extract summary data: {str(e)}")
                logger.warning(f"Summary extraction error details: {traceback.format_exc()}")

            logger.info(f"Successfully extracted {len(df)} monthly supply records")
            return df, title
            
        except Exception as e:
            logger.error(f"Failed to extract monthly supply table: {str(e)}")
            raise

    def navigate_to_analysis_report(self):
        """Navigate to the analysis report page"""
        try:
            nav_div = self.wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "nav"))
            )
            
            # Update XPath to match the exact menu text
            analysis_link = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(text(), '[606062] 銷售資料綜合分析')]")
                )
            )
            analysis_link.click()
            
            # Wait for the filter form to load
            self.wait.until(
                EC.presence_of_element_located((By.NAME, "b_ym"))
            )
            
            logger.debug("Analysis report page loaded")  # Add debug logging
            logger.info("Successfully navigated to analysis report page")
            
        except TimeoutException:
            logger.error("Timeout waiting for analysis report page elements")
            self.save_screenshot("analysis_navigation_error")
            raise
        except Exception as e:
            logger.error(f"Failed to navigate to analysis report page: {str(e)}")
            self.save_screenshot("analysis_navigation_error")
            raise
    
    def set_analysis_report_filter(self, year=None, month=None, filter_type='customer'):
        """Set filter for analysis report
        Args:
            year: Optional year to filter
            month: Optional month to filter
            filter_type: 'customer' or 'product' to determine which checkboxes to select
        """
        try:
            # Get date values
            date_values = self.filter_month_generator(year, month)
            combined_date = date_values['combined']
            
            logger.debug(f"Setting analysis filter for {combined_date}, type: {filter_type}")
            
            # Select start and end dates (same month for our case)
            for field_name in ['b_ym', 'e_ym']:
                date_select = self.wait.until(
                    EC.presence_of_element_located((By.NAME, field_name))
                )
                date_dropdown = Select(date_select)
                date_dropdown.select_by_value(combined_date)
            
            # Clear existing selections
            checkboxes = self.driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            for checkbox in checkboxes:
                if checkbox.is_selected():
                    checkbox.click()
            
            # Select appropriate checkboxes based on filter type
            if filter_type == 'customer':
                # Wait and select customer-related checkboxes
                self.wait.until(
                    EC.element_to_be_clickable((By.NAME, "acc_code"))
                ).click()
                self.wait.until(
                    EC.element_to_be_clickable((By.NAME, "acc_cat1"))
                ).click()
            else:  # product
                # Wait and select product-related checkboxes
                self.wait.until(
                    EC.element_to_be_clickable((By.NAME, "stk_c"))
                ).click()
                self.wait.until(
                    EC.element_to_be_clickable((By.NAME, "acc_cat"))
                ).click()
            
            # Submit form
            submit_button = self.wait.until(
                EC.element_to_be_clickable((By.NAME, "B1"))
            )
            submit_button.click()
            
            # Wait for results table
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//table[@bgcolor='#008080']"))
            )
            
            logger.info(f"Successfully set analysis filter type: {filter_type} for {combined_date}")

        except Exception as e:
            logger.error(f"Failed to set analysis report filter: {str(e)}")
            self.save_screenshot("analysis_filter_error")
            raise

    def extract_analysis_table(self):
        """Extract data from analysis report table based on current filter"""
        try:
            # Wait for table to be present
            table = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//table[@bgcolor='#008080']"))
            )
            
            # Get headers
            headers = []
            header_row = table.find_element(By.TAG_NAME, "tr")
            for th in header_row.find_elements(By.TAG_NAME, "td"):
                headers.append(th.text.strip())
            
            # Initialize lists to store data
            data = []
            
            # Get all rows except header
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # Skip header row
            
            # Process regular rows
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = []
                
                # Check if this is the total row
                is_total = False
                if cells[0].get_attribute("bgcolor") == "#CCFF66":
                    is_total = True
                
                for cell in cells:
                    value = cell.text.strip()
                    # If it's a total row and the cell spans multiple columns
                    if is_total and cell.get_attribute("colspan"):
                        row_data.append("合計")  # Add '合計' for the first column
                        # Add empty strings for spanned columns
                        for _ in range(int(cell.get_attribute("colspan")) - 1):
                            row_data.append("")
                    else:
                        row_data.append(value)
                        
                if row_data:  # Only add non-empty rows
                    data.append(row_data)
            
            # Create DataFrame
            df = pd.DataFrame(data, columns=headers)
            
            # Convert numeric columns
            numeric_columns = ['出量', '退量', '淨量']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col].replace('', '0'), errors='coerce')
            
            # Convert 退率 (return rate) - remove % and convert to numeric
            if '退率' in df.columns:
                df['退率'] = df['退率'].replace('', '0')
                df['退率'] = df['退率'].str.rstrip('%').astype(float)
            
            logger.info(f"Successfully extracted {len(df)} analysis records")
            return df
                
        except Exception as e:
            logger.error(f"Failed to extract analysis table: {str(e)}")
            self.save_screenshot("analysis_table_extraction_error")
            raise


    def export_to_excel(self, df, report_type, title=None):
        """Export the DataFrame to Excel with report type specification and optional title"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{report_type}_{timestamp}.xlsx"
            
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(current_dir)
            exports_dir = os.path.join(project_root, 'exports')
            os.makedirs(exports_dir, exist_ok=True)
            
            excel_path = os.path.join(exports_dir, filename)
            
            # Create Excel writer object
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # Always write DataFrame, adjust startrow based on title presence
                start_row = 1 if title else 0
                df.to_excel(writer, sheet_name=report_type, index=False, startrow=start_row)
                
                if title:
                    # Get the worksheet and write title
                    worksheet = writer.sheets[report_type]
                    worksheet.cell(row=1, column=1, value=title)
            
            logger.info(f"Successfully exported {report_type} data to {excel_path}")
            return excel_path
            
        except Exception as e:
            logger.error(f"Failed to export to Excel: {str(e)}")
            raise

    def save_screenshot(self, prefix):
        """Save screenshot on failure"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"error_screenshots/{prefix}_{timestamp}.png"
            os.makedirs("error_screenshots", exist_ok=True)
            self.driver.save_screenshot(filename)
            logger.info(f"Screenshot saved as {filename}")
        except Exception as e:
            logger.error(f"Failed to save screenshot: {str(e)}")

    def logout_and_quit(self):
        """Logout from the website and close the browser"""
        try:
            # Check if already logged out
            if not self.is_logged_in():
                logger.info("Already logged out")
                return

            # Find and click the logout link in nav bar
            logger.info("Looking for logout link...")
            logout_link = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//a[contains(@href, '{URLConfig.LOGOUT_PATH}')][text()='會員登出']")
                )
            )
            logger.info("Clicking logout link...")
            logout_link.click()
            
            # Wait for logout message
            self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '您目前登出系統中')]"))
            )
            
            # Wait for redirect to homepage
            self.wait.until(
                lambda driver: driver.current_url == URLConfig.BASE_URL + "/index.jsp"
            )
            
            # Clear cookies and session storage
            self.driver.delete_all_cookies()
            self.driver.execute_script("window.localStorage.clear();")
            self.driver.execute_script("window.sessionStorage.clear();")
            
            logger.info("Successfully logged out")
        except Exception as e:
            logger.error(f"Logout failed: {str(e)}")
            raise  # Re-raise the exception to handle it at a higher level
        finally:
            self.close()

    def close(self):
        """Close the browser"""
        try:
            if self.driver:
                self.driver.quit()
                logger.info("Browser session closed")
        except Exception as e:
            logger.error(f"Error closing browser: {str(e)}")
        finally:
            self.driver = None

    def is_logged_in(self):
        """Check if user is currently logged in"""
        try:
            logout_link = self.driver.find_element(By.XPATH, "//a[contains(@href, '/user_menu/user_logout.jsp')][text()='會員登出']")
            return logout_link.is_displayed()
        except:
            return False

    def extract_analysis_table(self):
        """Extract data from analysis report table based on current filter"""
        try:
            # Wait for table to be present
            table = self.wait.until(
                EC.presence_of_element_located((By.XPATH, "//table[@bgcolor='#008080']"))
            )
            
            # Get headers
            headers = []
            header_row = table.find_element(By.TAG_NAME, "tr")
            for th in header_row.find_elements(By.TAG_NAME, "td"):
                headers.append(th.text.strip())
            
            # Initialize lists to store data
            data = []
            
            # Get all rows except header
            rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # Skip header row
            
            # Process regular rows
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = []
                
                # Check if this is the total row
                is_total = False
                if cells[0].get_attribute("bgcolor") == "#CCFF66":
                    is_total = True
                
                for cell in cells:
                    value = cell.text.strip()
                    # If it's a total row and the cell spans multiple columns
                    if is_total and cell.get_attribute("colspan"):
                        row_data.append("合計")  # Add '合計' for the first column
                        # Add empty strings for spanned columns
                        for _ in range(int(cell.get_attribute("colspan")) - 1):
                            row_data.append("")
                    else:
                        row_data.append(value)
                        
                if row_data:  # Only add non-empty rows
                    data.append(row_data)
            
            # Create DataFrame
            df = pd.DataFrame(data, columns=headers)
            
            # Convert numeric columns
            numeric_columns = ['出量', '退量', '淨量']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col].replace('', '0'), errors='coerce')
            
            # Convert 退率 (return rate) - remove % and convert to numeric
            if '退率' in df.columns:
                df['退率'] = df['退率'].replace('', '0')
                df['退率'] = df['退率'].str.rstrip('%').astype(float)
            
            logger.info(f"Successfully extracted {len(df)} analysis records")
            return df
                
        except Exception as e:
            logger.error(f"Failed to extract analysis table: {str(e)}")
            self.save_screenshot("analysis_table_extraction_error")
            raise