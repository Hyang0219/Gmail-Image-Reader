#!/usr/bin/env python3
# main.py
import os
import argparse
import hashlib
from src.gmail_connector import GmailConnector
from src.email_processor import EmailProcessor
from src.attachment_handler import AttachmentHandler
from src.data_extractor import DataExtractor
from src.output_generator import OutputGenerator

def get_file_hash(file_path):
    """Calculate MD5 hash of a file to identify duplicates"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def main():
    parser = argparse.ArgumentParser(description='Process delivery notes from Gmail')
    parser.add_argument('--local', action='store_true', help='Process local files only (skip Gmail)')
    parser.add_argument('--use-ocr', action='store_true', help='Use OCR instead of OpenAI API')
    parser.add_argument('--gmail-credentials', default='credentials/credentials.json', 
                        help='Path to Gmail API credentials')
    parser.add_argument('--sheets-credentials', default='credentials/google_sheets_credentials.json',
                        help='Path to Google Sheets API credentials')
    parser.add_argument('--sheet-id', help='Google Sheet ID to save data')
    parser.add_argument('--output-csv', default='delivery_notes.csv', help='Output CSV filename')
    parser.add_argument('--search-query', default='subject:delivery note', help='Gmail search query')
    parser.add_argument('--local-dir', default='data/notes', help='Directory with local files')
    args = parser.parse_args()
    
    # Create all necessary directories
    os.makedirs('data/attachments', exist_ok=True)
    os.makedirs('data/output', exist_ok=True)
    
    # Initialize components
    extractor = DataExtractor(use_openai=not args.use_ocr)
    output_gen = OutputGenerator()
    
    extracted_data = []
    processed_files = set()  # Track processed files by hash to avoid duplicates
    
    # Process local files if requested
    if args.local:
        if os.path.exists(args.local_dir):
            print(f"Processing local files from {args.local_dir}")
            for filename in os.listdir(args.local_dir):
                file_path = os.path.join(args.local_dir, filename)
                
                if filename.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                    print(f"Processing: {filename}")
                    
                    mime_type = "application/pdf" if filename.lower().endswith('.pdf') else "image/png"
                    file_info = {
                        'path': file_path,
                        'filename': filename,
                        'mime_type': mime_type,
                        'email_metadata': {
                            'sender': 'Local File',
                            'date': 'Unknown',
                            'subject': filename
                        }
                    }
                    
                    try:
                        data = extractor.extract_from_file(file_info)
                        if data:
                            extracted_data.append(data)
                            print(f"Successfully extracted data from {filename}")
                        else:
                            print(f"Failed to extract data from {filename}")
                            # Try again with OCR if OpenAI failed
                            print("Retrying with OCR...")
                            ocr_extractor = DataExtractor(use_openai=False)
                            data = ocr_extractor.extract_from_file(file_info)
                            if data:
                                extracted_data.append(data)
                                print(f"Successfully extracted data with OCR from {filename}")
                            else:
                                print(f"Failed to extract data with OCR from {filename}")
                    except Exception as e:
                        print(f"Error processing {filename}: {e}")
                        # Try again with OCR if an exception occurred
                        try:
                            print("Retrying with OCR...")
                            ocr_extractor = DataExtractor(use_openai=False)
                            data = ocr_extractor.extract_from_file(file_info)
                            if data:
                                extracted_data.append(data)
                                print(f"Successfully extracted data with OCR from {filename}")
                            else:
                                print(f"Failed to extract data with OCR from {filename}")
                        except Exception as e2:
                            print(f"Error processing {filename} with OCR: {e2}")
        else:
            print(f"Local directory {args.local_dir} not found")
    else:
        # Connect to Gmail and process emails
        if not os.path.exists(args.gmail_credentials):
            print(f"Gmail credentials file not found: {args.gmail_credentials}")
            return
        
        print("Connecting to Gmail...")
        connector = GmailConnector(args.gmail_credentials)
        success, message = connector.test_connection()
        
        if not success:
            print(f"Failed to connect to Gmail: {message}")
            return
        
        print(message)
        service = connector.service
        
        # Search for emails
        print(f"Searching for emails matching: {args.search_query}")
        processor = EmailProcessor(service)
        messages = processor.search_emails(args.search_query)
        
        if not messages:
            print("No matching emails found")
            return
        
        print(f"Found {len(messages)} matching emails")
        
        # Process emails and download attachments
        handler = AttachmentHandler(service)
        
        for i, message in enumerate(messages):
            print(f"Processing email {i+1}/{len(messages)}")
            details = processor.get_email_details(message['id'])
            
            if not details:
                print("  Failed to get email details")
                continue
            
            print(f"  Subject: {details['subject']}")
            print(f"  From: {details['sender']}")
            print(f"  Date: {details['date']}")
            
            if not details.get('attachments'):
                print("  No attachments found")
                continue
            
            print(f"  Found {len(details['attachments'])} attachments")
            
            # Download attachments
            downloaded = handler.download_all_attachments(details)
            
            for attachment in downloaded:
                file_path = attachment['path']
                
                # Check if this file has already been processed (avoid duplicates)
                file_hash = get_file_hash(file_path)
                if file_hash in processed_files:
                    print(f"  Skipping duplicate attachment: {attachment['filename']}")
                    continue
                
                processed_files.add(file_hash)
                print(f"  Processing attachment: {attachment['filename']}")
                
                # Extract data from attachment
                data = extractor.extract_from_file(attachment)
                
                if data:
                    extracted_data.append(data)
                    print(f"  Successfully extracted data")
                else:
                    print(f"  Failed to extract data")
                    print("  Falling back to OCR...")
                    ocr_extractor = DataExtractor(use_openai=False)
                    data = ocr_extractor.extract_from_file(attachment)
                    if data:
                        extracted_data.append(data)
                        print(f"  Successfully extracted data with OCR")
                    else:
                        print(f"  Failed to extract data with OCR")
    
    # Generate output
    if extracted_data:
        print(f"Extracted data from {len(extracted_data)} files")
        
        # Save to CSV
        csv_path = output_gen.save_to_csv(extracted_data, args.output_csv)
        if csv_path:
            print(f"Data saved to CSV: {csv_path}")
        
        # Save to Google Sheets if requested
        if args.sheet_id and os.path.exists(args.sheets_credentials):
            print("Saving to Google Sheets...")
            if output_gen.setup_google_sheets(args.sheets_credentials):
                if output_gen.save_to_google_sheet(extracted_data, args.sheet_id):
                    print(f"Data saved to Google Sheet with ID: {args.sheet_id}")
                else:
                    print("Failed to save to Google Sheet")
        
        print("Processing complete")
    else:
        print("No data was extracted")

if __name__ == "__main__":
    main() 