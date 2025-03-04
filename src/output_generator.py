import os
import csv
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

class OutputGenerator:
    def __init__(self, output_dir='data/output'):
        """
        Initialize output generator.
        
        Args:
            output_dir: Directory to save output files
        """
        self.output_dir = output_dir
        self.google_sheets_client = None
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
    
    def save_to_csv(self, data_list, filename='delivery_notes.csv'):
        """
        Save extracted data to CSV file.
        
        Args:
            data_list: List of dictionaries containing extracted data
            filename: Output CSV filename
            
        Returns:
            Path to the CSV file
        """
        if not data_list:
            print("No data to save")
            return None
        
        output_path = os.path.join(self.output_dir, filename)
        
        # Flatten the data structure for CSV
        flattened_data = []
        for item in data_list:
            # Get shipping address, handling different field names
            shipping_address = item.get('shipping_address', item.get('buyer', ''))
            
            # Clean up shipping address - remove excess whitespace and newlines
            if shipping_address:
                shipping_address = re.sub(r'\s+', ' ', shipping_address.replace('\n', ' '))
            
            # Format total price to ensure it's a valid number
            total_price = item.get('total_price', '0.00')
            if not total_price or total_price == "Unknown":
                total_price = '0.00'
            
            # Try to clean up the total price
            try:
                # Remove any non-numeric characters except decimal point
                total_price = re.sub(r'[^\d.]', '', str(total_price))
                # Format as 2 decimal places
                total_price = f"{float(total_price):.2f}"
            except:
                total_price = '0.00'
            
            # If no products were found, create a single row with just the metadata
            if not item.get('products'):
                base_record = {
                    'sender': item.get('sender', ''),
                    'shipping_address': shipping_address,
                    'date': item.get('date', 'Unknown'),
                    'product_description': '',
                    'quantity': '',
                    'price': '',
                    'total_price': total_price
                }
                flattened_data.append(base_record)
            else:
                # Create a row for each product
                for product in item.get('products', []):
                    # Clean up price and quantity
                    price = product.get('price', '')
                    if price:
                        try:
                            # Remove any non-numeric characters except decimal point
                            price = re.sub(r'[^\d.]', '', str(price))
                            # Format as 2 decimal places if it's a valid number
                            price = f"{float(price):.2f}"
                        except:
                            price = ''
                    
                    quantity = product.get('quantity', '')
                    if quantity:
                        try:
                            # Remove any non-numeric characters
                            quantity = re.sub(r'[^\d]', '', str(quantity))
                        except:
                            quantity = ''
                    
                    # Calculate individual product total if not provided
                    product_total = total_price
                    if price and quantity and (total_price == '0.00' or len(item.get('products', [])) > 1):
                        try:
                            product_total = f"{float(price) * float(quantity):.2f}"
                        except:
                            product_total = total_price
                    
                    record = {
                        'sender': item.get('sender', ''),
                        'shipping_address': shipping_address,
                        'date': item.get('date', 'Unknown'),
                        'product_description': product.get('description', ''),
                        'quantity': quantity,
                        'price': price,
                        'total_price': product_total
                    }
                    flattened_data.append(record)
        
        # Write to CSV
        try:
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                if flattened_data:
                    fieldnames = flattened_data[0].keys()
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(flattened_data)
            return output_path
        except Exception as e:
            print(f"Error saving to CSV: {str(e)}")
            return None
    
    def setup_google_sheets(self, credentials_path):
        """
        Set up Google Sheets API client.
        
        Args:
            credentials_path: Path to Google Sheets API credentials
            
        Returns:
            Success status (boolean)
        """
        try:
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
            self.google_sheets_client = gspread.authorize(creds)
            return True
        except Exception as e:
            print(f"Error setting up Google Sheets: {str(e)}")
            return False
    
    def save_to_google_sheet(self, data_list, spreadsheet_key, worksheet_name='DeliveryNotes'):
        """
        Save data to Google Sheet.
        
        Args:
            data_list: List of dictionaries containing extracted data
            spreadsheet_key: Google Sheet ID
            worksheet_name: Name of the worksheet to update
            
        Returns:
            Success status (boolean)
        """
        if not self.google_sheets_client:
            print("Google Sheets client not initialized. Call setup_google_sheets() first.")
            return False
        
        if not data_list:
            print("No data to save")
            return False
        
        try:
            # Open the spreadsheet
            spreadsheet = self.google_sheets_client.open_by_key(spreadsheet_key)
            
            # Check if worksheet exists, create if not
            try:
                worksheet = spreadsheet.worksheet(worksheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=20)
            
            # Flatten the data for the sheet, similar to CSV
            flattened_data = []
            headers = ['sender', 'shipping_address', 'date', 'product_description', 'quantity', 'price', 'total_price']
            
            for item in data_list:
                # Get shipping address, handling different field names
                shipping_address = item.get('shipping_address', item.get('buyer', ''))
                
                # Clean up shipping address - remove excess whitespace and newlines
                if shipping_address:
                    shipping_address = re.sub(r'\s+', ' ', shipping_address.replace('\n', ' '))
                
                # Format total price
                total_price = item.get('total_price', '0.00')
                if not total_price or total_price == "Unknown":
                    total_price = '0.00'
                
                try:
                    total_price = re.sub(r'[^\d.]', '', str(total_price))
                    total_price = f"{float(total_price):.2f}"
                except:
                    total_price = '0.00'
                
                if not item.get('products'):
                    # Single row with just metadata
                    row = [
                        item.get('sender', ''),
                        shipping_address,
                        item.get('date', ''),
                        '',  # product_description
                        '',  # quantity
                        '',  # price
                        total_price
                    ]
                    flattened_data.append(row)
                else:
                    # Row for each product
                    for product in item.get('products', []):
                        # Clean up price and quantity
                        price = product.get('price', '')
                        if price:
                            try:
                                price = re.sub(r'[^\d.]', '', str(price))
                                price = f"{float(price):.2f}"
                            except:
                                price = ''
                        
                        quantity = product.get('quantity', '')
                        if quantity:
                            try:
                                quantity = re.sub(r'[^\d]', '', str(quantity))
                            except:
                                quantity = ''
                        
                        # Calculate individual product total if not provided
                        product_total = total_price
                        if price and quantity and (total_price == '0.00' or len(item.get('products', [])) > 1):
                            try:
                                product_total = f"{float(price) * float(quantity):.2f}"
                            except:
                                product_total = total_price
                        
                        row = [
                            item.get('sender', ''),
                            shipping_address,
                            item.get('date', ''),
                            product.get('description', ''),
                            quantity,
                            price,
                            product_total
                        ]
                        flattened_data.append(row)
            
            # Clear existing data and update with new data
            worksheet.clear()
            
            # Add headers and data
            all_rows = [headers] + flattened_data
            worksheet.update('A1', all_rows)
            
            # Format header row
            worksheet.format('A1:G1', {
                'textFormat': {'bold': True},
                'backgroundColor': {'red': 0.8, 'green': 0.8, 'blue': 0.8}
            })
            
            return True
            
        except Exception as e:
            print(f"Error saving to Google Sheet: {str(e)}")
            return False

# Test code
if __name__ == "__main__":
    # Sample data for testing
    test_data = [
        {
            'sender': 'test@example.com',
            'shipping_address': 'ABC Company, 123 Main St, Anytown, CA 12345',
            'date': '2023-01-01',
            'products': [
                {'description': 'Product 1', 'quantity': '5', 'price': '10.00'},
                {'description': 'Product 2', 'quantity': '3', 'price': '15.00'}
            ],
            'total_price': '95.00'
        }
    ]
    
    generator = OutputGenerator()
    csv_path = generator.save_to_csv(test_data, 'test_output.csv')
    print(f"CSV saved to: {csv_path}")
    
    # Test Google Sheets if credentials are available
    if os.path.exists('credentials/google_sheets_credentials.json'):
        if generator.setup_google_sheets('credentials/google_sheets_credentials.json'):
            # Replace with a valid spreadsheet ID
            success = generator.save_to_google_sheet(test_data, 'your_spreadsheet_id')
            print(f"Save to Google Sheets: {'Success' if success else 'Failed'}") 