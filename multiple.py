from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
import time

def get_real_time_price(url, real_time_selector):
    """Fetches and prints the real-time price from the website."""
    print("Fetching real-time data from the website...")
    
    # Initialize the Selenium WebDriver
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(15)  # Wait for the page to load fully

    # Parse the page source with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()

    # Find the real-time price element using the provided selector
    try:
        real_time_data = soup.select_one(real_time_selector)
        if real_time_data:
            price = real_time_data.get_text(strip=True)
            print(f"Real-time price: {price}")
        else:
            print("Real-time price data not found. The HTML structure may have changed.")
    except Exception as e:
        print(f"An error occurred while fetching real-time price: {e}")

def check_latest_row_in_excel(url_historical, excel_path, table_selector):
    """Checks if the latest row from the website is present in the Excel file."""
    print("Checking if the latest row is present in the Excel file...")
    
    try:
        existing_df = pd.read_excel(excel_path)
    except FileNotFoundError:
        print(f"Excel file not found at {excel_path}.")
        return
    
    # Fetch the latest row from the website
    driver = webdriver.Chrome()
    driver.get(url_historical)
    time.sleep(15)  # Wait for the page to load fully
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()
    
    try:
        # Find the table by the given selector
        table = soup.find('table', class_=table_selector)
        if table:
            rows = table.find_all('tr')
            if len(rows) > 1:  # Ensure there is at least one row (excluding the header)
                first_row = rows[1]  # The first row after the header
                cells = first_row.find_all('td')
                row_data = [cell.text.strip() for cell in cells]
                
                # Convert the date to the Excel format
                row_data[0] = pd.to_datetime(row_data[0], format='%b %d, %Y').strftime('%d-%m-%y')
                
                # Check if the date already exists in the 'Date' column
                if row_data[0] in existing_df['Date'].astype(str).values:
                    print(f"The row with date {row_data[0]} is already present.")
                else:
                    # Create a DataFrame with the new data
                    new_df = pd.DataFrame([row_data], columns=existing_df.columns[:len(row_data)])
                    
                    # Concatenate the new row at the top of the existing data
                    final_df = pd.concat([new_df, existing_df], ignore_index=True)
                    
                    # Save the updated DataFrame to Excel
                    final_df.to_excel(excel_path, index=False)
                    
                    # Adjust the column width using openpyxl
                    wb = load_workbook(excel_path)
                    ws = wb.active
                    for col in ws.columns:
                        max_length = 0
                        column = col[0].column_letter  # Get the column name
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        ws.column_dimensions[column].width = adjusted_width
                    wb.save(excel_path)

                    print(f"New row added with date: {row_data[0]}")
            else:
                print("No rows found in the table.")
        else:
            print("Table not found. The HTML structure may have changed.")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    print("Choose a currency pair:")
    print("1. EUR-INR")
    print("2. USD-INR")
    # Add more currencies as needed
    
    currency_choice = input("Enter the number of the currency pair you want to check: ").strip()
    
    if currency_choice == '1':
        # Set URLs and selectors for EUR-INR
        url_real_time = 'https://in.investing.com/currencies/eur-inr'
        url_historical = 'https://in.investing.com/currencies/eur-inr-historical-data'
        real_time_selector = '.mb-3.flex.flex-wrap.items-center.gap-x-4.gap-y-2.md\\:mb-0\\.5.md\\:gap-6'
        table_selector = 'freeze-column-w-1 w-full overflow-x-auto text-xs leading-4'  # Update selector based on the HTML structure
    elif currency_choice == '2':
        # Set URLs and selectors for USD-INR
        url_real_time = 'https://in.investing.com/currencies/usd-inr'
        url_historical = 'https://in.investing.com/currencies/usd-inr-historical-data'
        real_time_selector = '.mb-3.flex.flex-wrap.items-center.gap-x-4.gap-y-2.md\\:mb-0\\.5.md\\:gap-6'
        table_selector = 'freeze-column-w-1 w-full overflow-x-auto text-xs leading-4'  # Update selector based on the HTML structure
    else:
        print("Invalid choice. Please enter 1 or 2.")
        return
    
    print("Choose an option:")
    print("1. Check if the latest row is present in the Excel file")
    print("2. Get real-time data from the website")
    
    sub_choice = input("Enter your choice (1 or 2): ").strip()
    
    if sub_choice == '1':
        excel_path = input("Enter the path to the Excel file: ").strip()
        check_latest_row_in_excel(url_historical, excel_path, table_selector)
    
    elif sub_choice == '2':
        get_real_time_price(url_real_time, real_time_selector)
    
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()
