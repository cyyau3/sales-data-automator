from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from web_navigator import WebNavigator
from logger_config import logger
from datetime import datetime
import configparser
from pathlib import Path
import time

def load_config():
    try:
        config = configparser.ConfigParser()
        # Use Path for safer path handling
        current_dir = Path(__file__).parent
        project_root = current_dir.parent
        config_path = project_root / 'config' / 'config.ini'
        
        if not config_path.exists():
            raise FileNotFoundError(f"Config file not found at: {config_path}")
            
        config.read(config_path)
        
        return {
            'website_url': config['Credentials']['website_url'],
            'username': config['Credentials']['username'],
            'password': config['Credentials']['password'],
            'timeout': int(config['Settings']['timeout']),
            'browser': config['Settings']['browser']
        }
    except Exception as e:
        logger.error(f"Error loading config: {str(e)}")
        raise

def setup_driver():
    try:
        chrome_options = Options()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--lang=zh-TW')
        # Enhanced security options
        chrome_options.add_argument('--incognito')
        chrome_options.add_argument('--disable-cache')
        chrome_options.add_argument('--disable-application-cache')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--no-sandbox')
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except WebDriverException as e:
        logger.error(f"Failed to setup WebDriver: {str(e)}")
        raise

def perform_ucd_automation(config):
    navigator = WebNavigator(timeout=config['timeout'])
    try:
        # Create exports directory
        exports_dir = Path(__file__).parent.parent / 'exports'
        exports_dir.mkdir(exist_ok=True)
        
        # Login to UCD website
        logger.info(f"Attempting login for user: {config['username'][:2]}***")
        navigator.login(config['username'], config['password'])
        logger.info("Successfully logged in")

        # Create single Excel file for all reports
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        exports_dir = Path(__file__).parent.parent / 'exports'
        exports_dir.mkdir(exist_ok=True)  # Ensure exports directory exists
        excel_path = exports_dir / f'sales_data_{timestamp}.xlsx'

        # Export inventory data
        navigator.navigate_to_inventory()
        inventory_df = navigator.extract_inventory_table()
        navigator.export_to_excel(inventory_df, "inventory", excel_path=str(excel_path))
        
        # Return to index and export monthly supply
        navigator.return_to_index()
        navigator.navigate_to_monthly_supply()
        navigator.set_monthly_supply_filter()
        supply_df, supply_title = navigator.extract_monthly_supply_table()
        navigator.export_to_excel(supply_df, "monthly_supply", title=supply_title, excel_path=str(excel_path))
        
        # Return to index and export analysis reports
        navigator.return_to_index()
        navigator.navigate_to_analysis_report()
        
        # Customer analysis
        navigator.set_analysis_report_filter(filter_type='customer')
        customer_df = navigator.extract_analysis_table()
        navigator.export_to_excel(customer_df, "customer_analysis", excel_path=str(excel_path))
        
        # Product analysis
        navigator.set_analysis_report_filter(filter_type='product')
        product_df = navigator.extract_analysis_table()
        navigator.export_to_excel(product_df, "product_analysis", excel_path=str(excel_path))

        # Process weekly and monthly summary reports
        navigator.return_to_index()
        navigator.process_summary_reports(str(excel_path), 'weekly')
        
        navigator.return_to_index()
        navigator.process_summary_reports(str(excel_path), 'monthly')

        # Process order reports (purchase and return)
        navigator.return_to_index()
        navigator.process_order_reports(str(excel_path))

        # Payment menu and discount reports
        navigator.return_to_index()
        logger.info("Processing discount reports...")
        navigator.navigate_to_payment_menu()
        navigator.navigate_to_discount_detail()
        navigator.set_discount_filter()
        
        # Process main discount table and detail reports
        navigator.process_discount_report(str(excel_path))
        logger.info("Completed processing discount reports")

        # Process UCD payment detail
        navigator.return_to_index()
        navigator.navigate_to_payment_menu()
        navigator.navigate_to_payment_detail()
        navigator.set_payment_filter()
        result = navigator.process_payment_detail(str(excel_path))
        if result is None:
            logger.info("No payment details found for the period")
        else:
            logger.info(f"Successfully processed {len(result)} payment details")

        logger.info(f"All reports exported to {excel_path}")
        return navigator

    except Exception as e:
        logger.error(f"Error in automation: {str(e)}")
        raise

def main():
    navigator = None
    try:
        # Load configuration
        config = load_config()
        
        # Perform automation
        navigator = perform_ucd_automation(config)
        
        # Automatically logout and close browser
        logger.info("Initiating logout sequence...")
        navigator.logout_and_quit()  # This already includes closing the browser
        logger.info("Successfully logged out and closed browser")
        
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        if navigator:
            try:
                navigator.close()  # Fallback closing if something goes wrong
                logger.info("Browser closed after error")
            except Exception as close_error:
                logger.error(f"Error while closing browser: {str(close_error)}")

if __name__ == "__main__":
    main()