import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL of the webpage to scrape
url = 'https://www.wtm.com/atm/en-gb/exhibitor-directory.html#/'

# Send a GET request to the URL
response = requests.get(url)

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Find all h3 tags with the specified class
h3_tags = soup.find_all('h3', class_='text-center-mobile wrap-word')

# Extract the text content of the h3 tags
h3_text = [tag.get_text() for tag in h3_tags]

# Create a DataFrame to store the data
df = pd.DataFrame({'Title': h3_text})

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)

print("Excel file has been created with the extracted titles.")
