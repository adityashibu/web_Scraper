from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# URL of the webpage
url = "https://www.wtm.com/atm/en-gb/exhibitor-directory.html?refinementList%5B0%5D%5B0%5D=exhibitorFilters.Regions%20Operating%20In.lvl0%3Aid-677864&refinementList%5B1%5D%5B0%5D=exhibitorFilters.Regions%20Operating%20In.lvl1%3Aid-677865"

# Create a new instance of the Firefox driver
driver = webdriver.Firefox()

# Go to the URL
driver.get(url)

# Wait for the JavaScript to load the dynamic content
time.sleep(5)

# Get the scroll height
last_height = driver.execute_script("return document.body.scrollHeight")

h3_texts = []

while True:
    # Scroll down to the bottom
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Wait for the page to load
    time.sleep(5)

    # Parse the HTML of the page with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find all h3 tags with the class "text-center-mobile wrap-word"
    h3_tags = soup.find_all('h3', class_='text-center-mobile wrap-word')

    # Extract the text from these tags
    new_h3_texts = [tag.get_text() for tag in h3_tags]

    # If no new h3 tags were found, we've reached the end of the dynamic content
    if set(new_h3_texts).issubset(set(h3_texts)):
        break

    h3_texts = new_h3_texts

    # Calculate new scroll height and compare with last scroll height
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Create a DataFrame and save to an Excel file
df = pd.DataFrame(h3_texts, columns=['Title'])
df.to_excel('output.xlsx', index=False)

# Close the browser
driver.quit()
