from helium import *


# List of URLs to process
urls = [
    "url"
    # Add more URLs here
]

# XPath of the element to extract text and href from
xpath = "//h1[@class='result-title']//a"

# Dictionary to store extracted text and href
extracted_data = {}

# Start the browser
start_chrome(headless=False)

for url in urls:
    # Navigate to the URL
    go_to(url)
    
    # Find the element using XPath
    element = S(xpath)
    
    # Extract text and href if the element exists
    if element.exists():
        text = element.web_element.text
        href = element.web_element.get_attribute('href')
        extracted_data[text] = href
    else:
        extracted_data[url] = "No results found"

# Close the browser
kill_browser()

# Print the extracted data
print(extracted_data)
