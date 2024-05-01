from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
import time

# URL of the webpage
url = "https://www.wtm.com/atm/en-gb/exhibitor-directory.html#/"

# Create a new instance of the Firefox driver
driver = webdriver.Firefox()

# Go to the URL
driver.get(url)

# Wait for the JavaScript to load the dynamic content
time.sleep(5)

# Parse the HTML of the page with BeautifulSoup
soup = BeautifulSoup(driver.page_source, 'html.parser')

# Find all h3 tags with the class "text-center-mobile wrap-word"
h3_tags = soup.find_all('h3', class_='text-center-mobile wrap-word')

# Extract the text from these tags
h3_texts = [tag.get_text() for tag in h3_tags]

# Create a DataFrame and save to an Excel file
df = pd.DataFrame(h3_texts, columns=['Title'])
df.to_excel('output.xlsx', index=False)

# Close the browser
driver.quit()
