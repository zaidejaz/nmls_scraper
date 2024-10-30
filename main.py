import time
import pandas as pd
from bs4 import BeautifulSoup, Comment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from twocaptcha import TwoCaptcha
import base64
import os
import logging
from dotenv import load_dotenv

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load environment variables from .env file
load_dotenv()
API_KEY = os.getenv('API_KEY')
if not API_KEY:
    raise ValueError("API_KEY not found in environment variables.")

solver = TwoCaptcha(API_KEY)

# Read zip codes from CSV
zip_codes_df = pd.read_csv('zip_codes.csv')
zip_codes = zip_codes_df['zip_code'].tolist()

def solve_captcha(driver):
    try:
        logging.info("Attempting to solve CAPTCHA.")
        
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        if soup.find('form', id='aspnetForm'):
            checkbox = None
            # Agree to terms checkbox
            try:
                checkbox = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_MainContent_cbxAgreeToTerms"))
                )
            except:
                print("No checkbox found")
            if checkbox:
                checkbox.click()
                logging.info("Checked agree to terms checkbox.")

            # Captcha image element
            captcha_element = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "c_turingtestpage_ctl00_maincontent_captcha1_CaptchaImage"))
            )
            captcha_image_url = captcha_element.get_attribute('src')

            # Download CAPTCHA image using Selenium
            download_image(driver)
            # Solve CAPTCHA using 2Captcha
            result = solver.normal("captcha.jpeg")
            captcha_solution = result['code']
            logging.info(f"CAPTCHA solution: {captcha_solution}")

            # Input the CAPTCHA solution
            captcha_input = WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtTuringText"))
            )
            captcha_input.clear()
            captcha_input.send_keys(captcha_solution)

            # Click the continue button
            continue_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.ID, "ctl00_MainContent_btnContinue"))
            )
            continue_button.click()

            # Wait for the next page to load
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "td a.individual"))
            )
            logging.info("CAPTCHA solved successfully.")
        else:
            logging.info("No CAPTCHA found on this page.")

    except Exception as e:
        logging.error("Error solving captcha:", e)
        return False
    return True

def download_image(driver):
    try:
        # Wait for the CAPTCHA image element to load
        captcha_element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "c_turingtestpage_ctl00_maincontent_captcha1_CaptchaImage"))
        )

        # Take a screenshot of the CAPTCHA element
        screenshot_base64 = captcha_element.screenshot_as_base64

        # Save the screenshot as an image file locally
        img_file_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "captcha.jpeg")
        with open(img_file_path, "wb") as img_file:
            img_file.write(base64.b64decode(screenshot_base64))

        # Convert the image to base64
        with open(img_file_path, "rb") as image_file:
            img_base64 = base64.b64encode(image_file.read()).decode()

        return img_base64

    except Exception as e:
        logging.error("Error downloading CAPTCHA image:", e)
        return None

def get_individual_links(driver, zip_code):
    links = []
    search_url = f'https://www.nmlsconsumeraccess.org/Home.aspx/SubSearch?searchText={zip_code}'
    driver.get(search_url)
    logging.info(f"Opened search URL for zip code: {zip_code}")

    try:
        solve_captcha(driver)  # Attempt to solve CAPTCHA after initial page load

        time.sleep(5)  # Give time for the page to load and CAPTCHA solving
        content = driver.page_source
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.select('td a.individual')
        logging.info(f"Found {len(results)} results for zip code {zip_code}")
        for link in results:
            try:
                onclick_value = link.get('onclick')
                if onclick_value:
                    url_start_index = onclick_value.find("'") + 1
                    url_end_index = onclick_value.rfind("'")
                    if url_start_index != -1 and url_end_index != -1:
                        detail_url = onclick_value[url_start_index:url_end_index]
                        links.append(detail_url)
            except Exception as e:
                logging.error(f"Error extracting link: {e}")


    except Exception as e:
        logging.error(f"Error processing zip code {zip_code}: {e}")

    return links

def extract_details(driver, url):
    driver.get("https://www.nmlsconsumeraccess.org" + url)
    logging.info(f"Opened detail page URL: {url}")
    time.sleep(5)  # Consider increasing this wait time based on page load performance
    content = driver.page_source
    soup = BeautifulSoup(content, 'html.parser')
    details = {}
    try:
        name_element = soup.find('p', class_='individual')
        details['Name'] = name_element.text.strip() if name_element else ''
        # Extract NMLS ID, Phone, and Fax from the table
        nmls_table = soup.find('tr')
        if nmls_table:
            nmls_data = nmls_table.find_all('td', class_='divider')
            details['NMLS ID'] = nmls_data[0].text.strip() if nmls_data[0] else ''
            details['Phone'] = nmls_data[1].text.strip() if len(nmls_data) > 1 and nmls_data[1] else ''
        else:
            details['NMLS ID'] = ''
            details['Phone'] = ''
    except AttributeError:
        logging.warning("Essential details not found, skipping this entry.")
        return None
    
    # Extract office locations by looking for the "REGISTERED LOCATIONS" comment
    comments = soup.find_all(string=lambda text: isinstance(text, Comment) and "REGISTERED LOCATIONS" in text)
    if comments:
        registered_locations_comment = comments[0]
        office_section = registered_locations_comment.find_next('table')
        if office_section:
            rows = office_section.find_all('tr')[1:]  # Skip the header row
            for idx, row in enumerate(rows):
                cols = row.find_all('td')
                print(cols[0].text.strip())
                if cols[0].text.strip() == "None":
                    details['Company'] = ''
                    details['Company NMLS ID'] = ''
                    details['Type'] = ''
                    details['Street Address'] = ''
                    details['City'] = ''
                    details['State'] = ''
                    details['Zip Code'] = ''
                    details['Start Date'] = ''
                else:
                    details['Company'] = cols[0].text.strip() if cols[0] else ''
                    details['Company NMLS ID'] = cols[1].text.strip() if cols[1] else ''
                    details['Type'] = cols[2].text.strip() if cols[2] else ''
                    details['Street Address'] = cols[3].text.strip() if cols[3] else ''
                    details['City'] = cols[4].text.strip() if cols[4] else ''
                    details['State'] = cols[5].text.strip() if cols[5] else ''
                    details['Zip Code'] = cols[6].text.strip() if cols[6] else ''
                    details['Start Date'] = cols[7].text.strip() if cols[7] else ''
    logging.info(details)
    return details

def save_to_csv(data, filename):
    try:
        df = pd.DataFrame(data)
        if not os.path.isfile(filename):
            # If file does not exist, create it with a header
            df.to_csv(filename, index=False, mode='w')
        else:
            # If file exists, append data without writing the header
            df.to_csv(filename, index=False, mode='a', header=False)
        logging.info(f"Data saved to '{filename}'")
    except Exception as e:
        logging.error(f"Error saving data to CSV: {e}")

def main():
    try:
        # Setup Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")

        # Setup WebDriver
        logging.info("Setting up the WebDriver.")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        logging.info("WebDriver setup complete.")

        all_data = []
        for index, zip_code in enumerate(zip_codes):
            try:
                if zip_codes_df.at[index, 'status'] != 'Done':
                    logging.info(f"Processing zip code: {zip_code}")
                    individual_links = get_individual_links(driver, zip_code)
                    for url in individual_links:
                        details = extract_details(driver, url)
                        if details:
                            all_data.append(details)
                        time.sleep(5)  # To avoid being blocked for sending too many requests too quickly

                    # Mark zip code as done
                    zip_codes_df.at[index, 'status'] = 'Done'
                    zip_codes_df.to_csv('zip_codes.csv', index=False)  # Save the updated status

                    # Save data to CSV after each zip code
                    save_to_csv(all_data, 'loan_officers.csv')
                    all_data.clear()  # Clear the list to avoid appending the same data multiple times

            except Exception as e:
                logging.error(f"An error occurred for zip code {zip_code}: {e}")
                

    except Exception as e:
        logging.error(f"An error occurred in the main process: {e}")
        exit()

    finally:
        driver.quit()
        logging.info("WebDriver quit successfully.")

    logging.info("Data extraction complete.")

if __name__ == "__main__":
    main()
