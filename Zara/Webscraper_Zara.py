from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
import time
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.517 Safari/537.36'
chrome_options.add_argument('user-agent={0}'.format(user_agent))
# Path to chromedriver executable
chrome_driver_path = r'C:\Users\adrie\OneDrive\Bureau\Seconde Main\chromedriver-win64\chromedriver.exe' #A adapter

# Initialize Chrome webdriver with the specified options and path to chromedriver
driver = webdriver.Chrome(service=Service(chrome_driver_path), options=chrome_options)

# Target website URL
target_website = 'https://www.zara.com/ej/fr/preowned-resell/products/femme/'

# Open the webpage using Selenium WebDriver
driver.get(target_website)

# Time to scroll in seconds
SCROLL_DURATION = 10

# Get the start time for scrolling
start_time = time.time()

# Scroll for SCROLL_DURATION seconds
while (time.time() - start_time) < SCROLL_DURATION:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)  # Adjust the wait time if needed
    
#Define variables 
data = [] #To store into the file 
number_of_elements = 50 #Elements in the file
index = 0 #Current element
for i in range (number_of_elements) :
    product_cards = driver.find_elements(By.CLASS_NAME,'product-card') #Find all the cards at each iteration to avoid list being stale
    product_cards[index].click() # Click on the product card at the current index
    
    #Scrolling to ensure all html is uncovered
    driver.implicitly_wait(100)  
    SCROLL_DURATION = 3  
    start_time = time.time()
    while (time.time() - start_time) < SCROLL_DURATION:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        
    page_source = driver.page_source #HTML code
    soup = BeautifulSoup(page_source, 'html.parser')
    
    #Extract product name
    product_name_element = soup.find('div', class_='product-detail-ws__details__product__name')
    product_name = product_name_element.text.strip() if product_name_element else None

    # Extract product size and color
    product_size_element = soup.find('div', class_='product-detail-ws__details__product__size')
    product_size_color = product_size_element.text.strip().split('|') if product_size_element else [None, None]
    product_size = product_size_color[0].strip() if product_size_color else None
    product_color = product_size_color[1].strip() if len(product_size_color) > 1 else None

    # Extract product condition
    product_condition_element = soup.find('div', class_='product-detail-ws__details__product__condition')
    product_condition = product_condition_element.text.strip() if product_condition_element else None

    # Extract seller comment
    seller_comment_element = soup.find('div', class_='product-detail-ws__details__seller-comment')
    seller_comment = seller_comment_element.text.strip() if seller_comment_element else None

    # Extract year of purchase
    year_of_purchase_element = soup.find('div', class_='product-detail-ws__details__product__year-purchase')
    year_of_purchase = year_of_purchase_element.text.strip().replace('Année d\'achat : ', '') if year_of_purchase_element else None

    # Print extracted details
    data.append([index, product_name,product_size, product_color, product_condition,seller_comment, year_of_purchase])
    print(f"card {index}")
    print("Product Name:", product_name)
    print("Product Size:", product_size)
    print("Product Color:", product_color)
    print("Product Condition:", product_condition)
    print("Seller Comment:", seller_comment)
    print("Year of Purchase:", year_of_purchase)
    
    # Navigate back to the previous page
    driver.back()
    
    # Move to the next card
    index += 1

    # Adjust the wait time if needed
    time.sleep(2)

# Close the webdriver & save
driver.quit()
df = pd.DataFrame(data, columns=['Index', 'Product Name', 'Product Size', 'Product Color', 'Product Condition', 'Seller Comment', 'Year of Purchase'])
directory_path = r"C:\Users\adrie\OneDrive\Bureau\Seconde Main\Zara\Zara_data"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename = f"products_{timestamp}.xlsx"
file_path = directory_path + "/" + filename
df.to_excel(file_path, index=False)

print(f"File saved successfully: {file_path}") 
