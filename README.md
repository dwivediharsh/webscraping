Currency Data Fetcher
This Python script helps you fetch real-time and historical currency data from a specified website and manage it within an Excel file. The script supports two main functionalities:

Real-Time Price Retrieval:

Fetches and prints the real-time price of a chosen currency pair from a designated website.
Uses Selenium WebDriver to load the webpage and BeautifulSoup to parse the HTML for the price.
Historical Data Management:

Checks if the latest row of historical data from the website is already present in an Excel file.
If the row does not exist, it adds the new data to the Excel file and adjusts column widths for better readability.
Usage
Choose a Currency Pair:

Select from available currency pairs (e.g., EUR-INR or USD-INR).
Choose an Option:

Check Historical Data: Verify if the latest data is already in the Excel file and add it if not.
Get Real-Time Data: Retrieve and display the current price of the chosen currency pair.
Setup
Ensure you have Selenium, BeautifulSoup, and pandas installed.
Download the Chrome WebDriver and ensure it is available in your system's PATH.
Instructions
Run the script and follow the prompts to:

Select a currency pair.
Choose whether to check historical data or get real-time prices.
Provide the path to the Excel file for historical data management.
This script is designed to be easily extendable for additional currencies or other data sources.
