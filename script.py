"""
LXT118
8/17/2024
Python eBay scraper to find highest and lowest price for figurines - DEFINITIONS FILE
Got help with the parsing and scraping part using https://scrapfly.io/blog/how-to-scrape-ebay/
"""

# import all necessary modules and specific function, with alias as required
import tkinter
import tkinter.filedialog
import docx
import pandas as pd
from parsel import Selector
import httpx
import re
import tkinter, time

# file path of the query input file
tkinter.Tk().withdraw()
print("Opening directory explorer. You will be asked to locate the \nword file containing the queries found in the same directory \nas this script...")
time.sleep(3.5)
file_path = tkinter.filedialog.askopenfilename()
if re.search(".docx$", file_path):
    print("Reading and opening word file...\n")
    pass
else:
    raise TypeError("Only the .docx file in the working directory of this script is allowed")

# the dictionary used to collect the raw scrape data to make the pandas DataFrame object
item_all = {}

# Read search queries from a Word file, and convert them to scraping format URLs
def queries_to_url(file_path: str):
    doc = docx.Document(file_path)
    url_base = "https://www.ebay.ca/sch/i.html?_from=R40&_nkw=^^^&_sacat=0&_ipg=240"
    queries = [para.text for para in doc.paragraphs if para.text.strip() != ""]
    new_queries = [query.replace(" ", "+") for query in queries]
    global urls
    urls = [url_base.replace("^^^", query) for query in new_queries]
    return urls

# Take the HTML files for each query, parse out the query used, its associated eBay URL, the price of each listing, and the name of each listing
def parse_product(response: httpx.Response, url) -> dict:
    """Parse Ebay's product listing page for core product data"""
    sel = Selector(response.text) # Create a CSS selector object of class Selector
    # define helper functions that chain the extraction process
    css_join = lambda css: "\n".join(sel.css(css).getall()).strip()  # join all CSS selected elements
    item = {}
    prices = css_join('.s-item__price>span::text').replace("C $", "").split("\n")
    item["url"] = [url] * len(prices)
    item["query"] = [re.search(r"(?<=nkw=).*?(?=&_sacat)", url).group()] * len(prices) # regular expressions to extract the query from the URL
    item["price"] = prices
    item["name"] = list(filter(("Shop on eBay").__ne__, css_join('.s-item__title>span::text').split("\n")))
    return item

# create an HTML request client to pull data from each URL
with httpx.Client(headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.35",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",}, 
        http2=True,
        follow_redirects = True) as client:
    
    # Take each query from the word doc, turn it into a legitimate eBay URL, and parse out information, and append to a master dictionary 
    for url in queries_to_url(file_path):
        response = client.get(url)
        item = parse_product(response, url)
        for key, value in item.items():
            if key in item_all:
                item_all[key].extend(value)
            else:
                item_all[key] = value

# using Pandas library, turn master dictionary into a Pandas dataframe
df = pd.DataFrame.from_dict(item_all, orient = 'index').transpose() # fixes the issue with all Series column object arrays not being of same length

# manipulate basic dataframe for each Series column to have correct properties, excluding any observations that are incomplete or incorrect
df = df[df["price"] != " to "]
df = df[df["price"].notna()] # remove no price rows
df["url"] = df["url"].astype(str)
df["query"] = df["query"].astype(str); df["query"] = df["query"].apply(lambda x: re.sub("\+", " ", x))
df["price"] = [re.sub(",", "", price) for price in df["price"]]
df["price"] = pd.to_numeric(df["price"])

# Remove any rows for which the query terms are not found in the name
#df = df[df["query"].split(), df["name"]] re.split(" ", ...)

# Ask where the excel file containing the data would like to be saved
print("Where would you like to save the excel file?")
excel_filepath = tkinter.filedialog.askdirectory()

## STATISTICS
# Define statistics (numeric_only flag is necessary for this version of Pandas)
min_price = df[["query", "price"]].groupby("query").min(numeric_only = True) # minimum price
max_price = df[["query", "price"]].groupby("query").max(numeric_only = True) # maximum price
avg_price = df[["query", "price"]].groupby("query").mean(numeric_only = True) # average price

# Print the lowest, highest, and mean prices for each query
print(f"The minimum prices per queries are: \n{min_price}\n")
print(f"The maximum prices per queries are: \n{max_price}\n")
print(f"The average prices per queries are: \n{avg_price}\n")

# Export all the data as an excel file
df.to_excel(f"{excel_filepath}/all_listings.xlsx", sheet_name = "Sheet1")

# For script running in console, prompt user to press a key before the script exits successfully to allow the user to see python output
input("Press enter to exit...")