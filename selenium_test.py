from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
import time

# URL of the webpage
url = "https://www.wtm.com/atm/en-gb/exhibitor-directory.html?refinementList%5B0%5D%5B0%5D=exhibitorFilters.Regions%20Operating%20In.lvl0%3Aid-677864&refinementList%5B1%5D%5B0%5D=exhibitorFilters.Regions%20Operating%20In.lvl1%3Aid-677888"

# Create a new instance of the Firefox driver
driver = webdriver.Firefox()

# Go to the URL
driver.get(url)

# Wait for the JavaScript to load the dynamic content
time.sleep(10)  # Increase the delay if necessary

# Get the scroll height
last_height = driver.execute_script("return document.body.scrollHeight")

# Initialize lists to store extracted data
company_names = []
emails = []
websites = []
phones = []

while True:
    # Scroll down to the bottom
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)  # Wait for the page to load
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)  # Wait for the page to load

    # Calculate new scroll height and compare with last scroll height
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

    # Parse the HTML of the page with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all div tags with the specified class
    div_tags = soup.find_all('div', class_='directory-item-feature-toggled exhibitor-category row')

    for div in div_tags:
        # Find the company name within this div
        company_name_tag = div.find('div', class_='company-info').find('a').find('h3')
        if company_name_tag:
            company_name = company_name_tag.get_text().strip()
        else:
            company_name = ''

        # Extract company information
        contact_options = div.find('ul', class_='contact-options-container-package-redesign')
        if contact_options:
            links = contact_options.find_all('a')
            for link in links:
                if 'Website' in link['aria-label']:
                    website = link['href']
                elif 'Email' in link['aria-label']:
                    email = link['href'].replace('mailto:', '')
                elif 'Phone' in link['aria-label']:
                    phone = link['href'].replace('tel:', '')
                else:
                    website = ''
                    email = ''
                    phone = ''
        else:
            website = ''
            email = ''
            phone = ''

        # Append extracted data to lists
        company_names.append(company_name)
        emails.append(email)
        websites.append(website)
        phones.append(phone)

# Create a DataFrame and save to an Excel file
df = pd.DataFrame({'Company Name': company_names, 'Company Email': emails, 'Company Website': websites, 'Company Phone': phones})
df.to_excel('output.xlsx', index=False)

# Close the browser
driver.quit()

print("Excel file has been created with the extracted company information.")
