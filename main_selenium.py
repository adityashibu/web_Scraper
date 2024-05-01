from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# URL of the webpage
url = "https://www.wtm.com/atm/en-gb/exhibitor-directory.html#/"

# Create a new instance of the Firefox driver
driver = webdriver.Firefox()

# Go to the URL
driver.get(url)

# Wait for the JavaScript to load the dynamic content
time.sleep(10)  # Increase the delay if necessary

# Get the scroll height
last_height = driver.execute_script("return document.body.scrollHeight")

h3_texts = []
emails = []

while True:
    # Scroll down to the bottom
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Wait for the page to load
    time.sleep(5)

    # Parse the HTML of the page with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all div tags with the specified class
    div_tags = soup.find_all('div', class_='directory-item-feature-toggled exhibitor-category row')

    for div in div_tags:
        # Find the h3 tag and the email within this div
        h3_tag = div.find('h3', class_='text-center-mobile wrap-word')
        email_tag = div.find('a', {'aria-label': 'Company Email'})

        if h3_tag:
            h3_texts.append(h3_tag.get_text())
            if email_tag:
                emails.append(email_tag['href'].replace('mailto:', ''))
            else:
                emails.append('')  # Add an empty string if no email is found

    # Wait for the page to load new content
    time.sleep(5)

    # Calculate new scroll height and compare with last scroll height
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Create a DataFrame and save to an Excel file
df = pd.DataFrame({'Title': h3_texts, 'Company Email': emails})
df.to_excel('output.xlsx', index=False)

# Close the browser
driver.quit()

print("Excel file has been created with the extracted titles.")
