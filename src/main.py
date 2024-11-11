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
import os
import time

def load_config():
    try:
        config = configparser.ConfigParser()
        # Modified path to look for config.ini in the parent directory's config folder
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        config_path = os.path.join(project_root, 'config', 'config.ini')
        
        if not os.path.exists(config_path):
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
        # Add security options
        chrome_options.add_argument('--incognito')
        chrome_options.add_argument('--disable-cache')
        chrome_options.add_argument('--disable-application-cache')
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except WebDriverException as e:
        logger.error(f"Failed to setup WebDriver: {str(e)}")
        raise

def perform_ucd_automation(config):
    navigator = WebNavigator(timeout=config['timeout'])
    try:
        # Login to UCD website
        logger.info(f"Attempting login for user: {config['username'][:2]}***")
        navigator.login(config['username'], config['password'])
        logger.info("Successfully logged in")

        # Create single Excel file for all reports
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'exports',
            f'sales_data_{timestamp}.xlsx'
        )

        # Export inventory data
        navigator.navigate_to_inventory()
        inventory_df = navigator.extract_inventory_table()
        navigator.export_to_excel(inventory_df, "inventory", excel_path=excel_path)
        
        # Return to index and export monthly supply
        navigator.return_to_index()
        navigator.navigate_to_monthly_supply()
        navigator.set_monthly_supply_filter()
        supply_df, supply_title = navigator.extract_monthly_supply_table()
        navigator.export_to_excel(supply_df, "monthly_supply", title=supply_title, excel_path=excel_path)
        
        # Return to index and export analysis reports
        navigator.return_to_index()
        navigator.navigate_to_analysis_report()
        
        # Customer analysis
        navigator.set_analysis_report_filter(filter_type='customer')
        customer_df = navigator.extract_analysis_table()
        navigator.export_to_excel(customer_df, "customer_analysis", excel_path=excel_path)
        
        # Product analysis
        navigator.set_analysis_report_filter(filter_type='product')
        product_df = navigator.extract_analysis_table()
        navigator.export_to_excel(product_df, "product_analysis", excel_path=excel_path)

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
        
        # Perform automation and keep browser open
        navigator = perform_ucd_automation(config)
        
        # Keep the script running until user input
        while True:
            user_input = input("\nIn this terminal, enter 'q' to logout and quit, or press Enter to keep browsing: ").lower()
            if user_input == 'q':
                logger.info("Initiating logout sequence...")
                navigator.logout_and_quit()  # This already includes closing the browser
                logger.info("Successfully logged out and closed browser")
                break
        
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