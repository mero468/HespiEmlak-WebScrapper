import pandas as pd
from openpyxl.reader.excel import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
import time
import random

# WebDriver is being started - Chrome browser
print("Starting browser...")

# Chrome options for setting User-Agent
chrome_options = Options()
user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'

chrome_options.add_argument(f"--user-agent={user_agent}")

# Start Chrome WebDriver with the specified options
driver = webdriver.Chrome(options=chrome_options)

# List to store listings
listing_list = []

# Create an Excel file
excel_file = 'listings.xlsx'
(pd.DataFrame(columns=[
    'Title',
    'Location',
    'Price',
    'Number of Rooms',
    'Area',
    'Building Age',
    'Floor',
    'Date',
    'Link',
    'Furniture Status',
    'Heating Type',
    'Fuel Type',
    'Deposit',
    'Maintenance Fee',
    'House Size'
    ])
 .to_excel(excel_file, index=False))

# Total number of pages. Each page has 24 listings
total_pages = 226

for page in range(85, total_pages + 1):
    # Create page URL
    url = f'https://www.hepsiemlak.com/izmir-kiralik?page={page}'

    # Clear all cookies before visiting a new page
    driver.delete_all_cookies()

    # Load page
    print(f"Loading page {page}: {url}")
    driver.get(url)

    # Wait for the page to load
    WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, 'listing-item')))
    print("Page loaded.")

    # After the page loads, delete the specific Cloudflare cookie if needed
    driver.delete_cookie('__cf_bm')

    # Find listings
    listings = driver.find_elements(By.CLASS_NAME, 'listing-item')
    print(f"{len(listings)} listings found.")

    # List to store links for the listings on the page
    page_listing_links = []

    # Process listings
    for listing in listings:
        try:
            # Extract title
            title = listing.find_element(By.CLASS_NAME, 'card-link').get_attribute('title')

            # Extract date
            date = listing.find_element(By.CLASS_NAME, 'list-view-date').text

            # Extract price
            price = listing.find_element(By.CLASS_NAME, 'list-view-price').text

            # Extract properties
            properties = listing.find_element(By.CLASS_NAME, 'short-property').text.split("\n")

            # Check properties list and extract required fields
            num_rooms = properties[1] if len(properties) > 1 else ''
            area = properties[2] if len(properties) > 2 else ''
            building_age = properties[3] if len(properties) > 3 else ''
            floor = properties[4] if len(properties) > 4 else ''

            # Extract location
            location = listing.find_element(By.CLASS_NAME, 'list-view-location').text

            # Get the listing link and add it to the list of listing links
            listing_link = listing.find_element(By.CLASS_NAME, 'card-link').get_attribute('href')
            page_listing_links.append(listing_link)

            # Add listing data to the list
            listing_list.append({
                'Title': title,
                'Location': location,
                'Price': price,
                'Number of Rooms': num_rooms,
                'Area': area,
                'Building Age': building_age,
                'Floor': floor,
                'Date': date,
                'Link': listing_link,
            })

            print(f"Listing added: {title}, {location}, {price}, {num_rooms}, {area}, {building_age}, {floor}, {date}, {listing_link}")
        except Exception as e:
            print(f"Error processing listing: {e}")

    # Iterate through the listing links
    for listing_link in page_listing_links:
        try:
            # Navigate to the listing link
            print("Loading listing link: ", listing_link)
            driver.get(listing_link)

            # Wait for the page to load
            WebDriverWait(driver, 10).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, 'txt')))
            print("Listing link loaded.")

            # Fetch all spans
            spans = driver.find_elements(By.CLASS_NAME, 'txt')

            furniture_status = ''
            heating_type = ''
            fuel_type = ''
            deposit = ''
            maintenance_fee = ''
            gross_net = ''

            for span in spans:
                if span.text == 'Eşya Durumu':  # If span contains 'Furniture Status'
                    furniture_status = span.find_element(By.XPATH, "following-sibling::span").text

                elif span.text == 'Isınma Tipi':  # If span contains 'Heating Type'
                    heating_type = span.find_element(By.XPATH, "following-sibling::span").text

                elif span.text == 'Yakıt Tipi':  # If span contains 'Fuel Type'
                    fuel_type = span.find_element(By.XPATH, "following-sibling::span").text
                elif span.text == 'Depozito':  # If span contains 'Deposit'
                    deposit = span.find_element(By.XPATH, "following-sibling::span").text
                elif span.text == 'Aidat':  # If span contains 'Maintenance Fee'
                    maintenance_fee = span.find_element(By.XPATH, "following-sibling::span").text
                elif span.text == 'Brüt / Net M2':  # If span contains 'Gross / Net Size'
                    gross_net = span.find_element(By.XPATH, "following-sibling::span").text
                if furniture_status and fuel_type and heating_type != '':
                    break

            listing_index = [listing['Link'] for listing in listing_list].index(listing_link)
            listing_list[listing_index]['Furniture Status'] = furniture_status
            listing_list[listing_index]['Heating Type'] = heating_type
            listing_list[listing_index]['Fuel Type'] = fuel_type
            listing_list[listing_index]['Deposit'] = deposit
            listing_list[listing_index]['Maintenance Fee'] = maintenance_fee
            listing_list[listing_index]['House Size'] = gross_net
            print(f"Listing data added: {furniture_status}, {heating_type}, {fuel_type}")

            # Dynamically write the listing to the Excel file
            df = pd.DataFrame([listing_list[listing_index]])
            rows = df.values.tolist()

            workbook = load_workbook('listings.xlsx')
            sheet = workbook.active

            for row in rows:
                sheet.append(row)
            workbook.save('listings.xlsx')

            print(f"Listing added to Excel: {listing_link}")
        except Exception as e:
            print(f"Error processing listing link: {e}")

        # Random delay to mimic human interaction
        time.sleep(random.uniform(3, 5))

    print(f"Page {page} processed.")

# Close the browser
print("Closing browser...")
driver.quit()
