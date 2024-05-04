import tkinter as tk
from tkinter import filedialog, ttk
from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
import time
import threading

# Function to perform web scraping
def scrape_exhibitor_info(url, class_name):
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
        div_tags = soup.find_all('div', class_=class_name)

        for div in div_tags:
            # Find the company name within this div
            if class_name == 'directory-item-feature-toggled exhibitor-category row':
                company_name_tag = div.find('div', class_='company-info').find('a').find('h3')
            else:
                company_name_tag = div.find('h3', class_='text-center-mobile wrap-word')
            if company_name_tag:
                company_name = company_name_tag.get_text().strip()
            else:
                company_name = ''

            # Extract company information
            if class_name == 'directory-item-feature-toggled exhibitor-category row':
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
            else:
                # Open the link in a new tab
                link_tag = div.find('a')
                if link_tag:
                    link_url = link_tag['href']
                    driver.execute_script(f"window.open('{link_url}', '_blank');")
                    # Switch to the new tab (assuming it's the last one)
                    driver.switch_to.window(driver.window_handles[-1])

                    # Wait for the page to load
                    time.sleep(5)

                    # Parse the HTML of the page with BeautifulSoup
                    soup_new_tab = BeautifulSoup(driver.page_source, 'html.parser')

                    # Find the correct div containing company information
                    div_new_tab = soup_new_tab.find('div', class_='col-md-12 col-sm-6')

                    # Check if the div contains company information
                    if div_new_tab:
                        # Extract company information from the side row div

                        # Try to extract email
                        try:
                            email = div_new_tab.find('div', id='exhibitor_details_email').find('a').text.strip()
                        except AttributeError:
                            email = ''

                        # Try to extract website
                        try:
                            website = div_new_tab.find('div', id='exhibitor_details_website').find('a').text.strip()
                        except AttributeError:
                            website = ''

                        # Try to extract phone
                        try:
                            phone = div_new_tab.find('div', id='exhibitor_details_phone').find('a').text.strip()
                        except AttributeError:
                            phone = ''

                        # Close the new tab and switch back to the main tab
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])

            # Append extracted data to lists
            company_names.append(company_name)
            emails.append(email)
            websites.append(website)
            phones.append(phone)

        # Wait for the page to load new content
        time.sleep(5)

    # Close the browser
    driver.quit()

    # Create a DataFrame
    df = pd.DataFrame({'Company Name': company_names, 'Company Email': emails, 'Company Website': websites, 'Company Phone': phones})
    return df

# Function to start scraping process
def start_scraping():
    global progress_bar, submit_button, url_entry

    # Disable submit button
    submit_button.config(state='disabled')

    # Get URL from entry
    url = url_entry.get()

    # Perform scraping in a separate thread
    threading.Thread(target=perform_scraping, args=(url,)).start()

# Function to perform scraping
def perform_scraping(url):
    global progress_bar, submit_button

    # URLs and class names for scraping
    urls_and_classes = [
        (url, 'directory-item-feature-toggled exhibitor-category row'),
        (url, 'directory-item directory-item-feature-toggled exhibitor-category')
    ]

    # Initialize an empty list to store DataFrames
    dfs = []

    # Loop through each URL and class name
    for url, class_name in urls_and_classes:
        df = scrape_exhibitor_info(url, class_name)
        dfs.append(df)

    # Concatenate DataFrames
    final_df = pd.concat(dfs, ignore_index=True)

    # Save the final DataFrame to an Excel file
    save_file(final_df)

# Function to save DataFrame to Excel file
def save_file(df):
    global submit_button

    # Enable submit button
    submit_button.config(state='normal')

    # Ask user for file save location
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    # Check if user canceled save operation
    if file_path == "":
        return

    # Save DataFrame to Excel file
    df.to_excel(file_path, index=False)

    # Inform user that file has been saved
    tk.messagebox.showinfo("File Saved", "Excel file has been created with the extracted company information.")

# Create GUI
root = tk.Tk()
root.title("Web Scraping Tool")

# URL Entry
url_label = tk.Label(root, text="Enter URL:")
url_label.grid(row=0, column=0, padx=5, pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5)

# Submit Button
submit_button = tk.Button(root, text="Submit", command=start_scraping)
submit_button.grid(row=0, column=3, padx=5, pady=5)

# Progress Bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.grid(row=1, column=0, columnspan=4, padx=5, pady=5)

root.mainloop()
