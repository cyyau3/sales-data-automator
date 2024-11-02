from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def setup_driver():
    # Set up Chrome options
    chrome_options = Options()
    # Add options if needed
    
    # Set up the Chrome driver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def test_setup():
    driver = setup_driver()
    try:
        # Test by opening a website
        driver.get("https://www.google.com")
        print("Setup successful! Browser opened Google.")
    finally:
        driver.quit()

if __name__ == "__main__":
    test_setup()