# Delivery Note Reader

An automated system to process delivery notes from Gmail, extract key information, and generate structured output in CSV or Google Sheets format.

## Features

- Connect to Gmail and search for emails with "delivery note" in the subject
- Download attachments from matching emails (supports PDF and image formats)
- Handle both regular attachments and inline images
- Extract data from delivery notes using OpenAI's Vision API for high accuracy
- Automatic fallback to OCR (Optical Character Recognition) if OpenAI API fails
- Identify key information such as sender, shipping address, date, products, quantities, and prices
- Generate structured output in CSV format
- Optional Google Sheets integration
- Deduplication of attachments to avoid processing the same file multiple times

## Requirements

- Python 3.6+
- OpenAI API key (with sufficient quota)
- Tesseract OCR (for fallback processing)
- Gmail API credentials (for Gmail integration)
- Google Sheets API credentials (optional, for Google Sheets integration)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/delivery-note-reader.git
   cd delivery-note-reader
   ```

2. Create a virtual environment and install dependencies:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Set up OpenAI API:
   - Get an API key from [OpenAI](https://platform.openai.com/)
   - Create a `.env` file in the project root based on the `.env.sample` file
   - Add your OpenAI API key to the `.env` file

4. Install Tesseract OCR for fallback processing:
   - **macOS**: `brew install tesseract`
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **Windows**: Download and install from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)

5. Set up Gmail API credentials:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project
   - Enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials JSON file and save it as `credentials/credentials.json`
   - Copy `credentials/credentials.json.sample` to `credentials/credentials.json` and update with your actual credentials

6. (Optional) Set up Google Sheets API credentials:
   - Enable the Google Sheets API in your Google Cloud project
   - Create a service account
   - Download the service account JSON key and save it as `credentials/google_sheets_credentials.json`
   - Copy `credentials/google_sheets_credentials.json.sample` to `credentials/google_sheets_credentials.json` and update with your actual credentials
   - Share your target Google Sheet with the service account email

7. Create necessary directories:
   ```
   mkdir -p data/attachments data/output data/notes
   ```

## Usage

### Process Local Files with OpenAI

To process local delivery note files using OpenAI's Vision API:

```
python main.py --local --local-dir=data/notes
```

### Process Local Files with OCR (without OpenAI)

If you don't have an OpenAI API key or prefer to use OCR:

```
python main.py --local --use-ocr
```

### Process Emails from Gmail

To download and process delivery notes from Gmail:

```
python main.py
```

By default, this will search for emails with "delivery note" in the subject. You can customize the search query:

```
python main.py --search-query="subject:invoice from:supplier@example.com"
```

### Save to Google Sheets

To save the extracted data to a Google Sheet:

```
python main.py --sheet-id=YOUR_GOOGLE_SHEET_ID
```

### Additional Options

```
python main.py --help
```

## Troubleshooting

### OpenAI API Issues

If you're experiencing issues with the OpenAI API:
- Ensure your API key is correctly set in the `.env` file
- Check your API usage limits on the OpenAI dashboard
- The application will automatically fall back to OCR-based extraction if OpenAI API fails
- You can also use the `--use-ocr` flag to force OCR-based extraction

### OCR Issues

If you're experiencing poor OCR results:
- Ensure Tesseract is properly installed
- Try improving the quality of the input images
- Adjust the data extraction patterns in `data_extractor.py`

### Gmail Authentication

If you encounter Gmail authentication issues:
- Ensure your credentials file is correctly placed
- Delete the `token.pickle` file in the credentials directory to force re-authentication
- Check that your Google Cloud project has the Gmail API enabled
- Make sure your email is added as a test user in the Google Cloud Console if your app is in testing mode

### Google Sheets Integration

If Google Sheets integration isn't working:
- Verify your service account credentials
- Ensure you've shared your Google Sheet with the service account email
- Check that the Google Sheets API is enabled in your Google Cloud project

## Development

### Project Structure

```
delivery-note-reader/
├── credentials/                # API credentials
│   ├── credentials.json.sample        # Sample Gmail API credentials
│   └── google_sheets_credentials.json.sample  # Sample Google Sheets credentials
├── data/
│   ├── attachments/            # Downloaded email attachments
│   ├── notes/                  # Local files for processing
│   └── output/                 # Generated output files
├── src/
│   ├── attachment_handler.py   # Handles email attachments
│   ├── data_extractor.py       # Extracts structured data from files
│   ├── email_processor.py      # Processes emails from Gmail
│   ├── gmail_connector.py      # Connects to Gmail API
│   ├── openai_extractor.py     # Extracts data using OpenAI
│   └── output_generator.py     # Generates output files
├── .env.sample                 # Sample environment variables
├── .gitignore                  # Git ignore file
├── main.py                     # Main application entry point
├── README.md                   # This file
└── requirements.txt            # Python dependencies
```

### Testing

The repository includes several test scripts in the `test` directory:

- `test_gmail.py`: Test Gmail API connection
- `test_ocr.py`: Test OCR-based extraction
- `test_openai.py`: Test OpenAI-based extraction

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 