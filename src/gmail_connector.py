import os
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class GmailConnector:
    def __init__(self, credentials_path):
        """
        Initialize Gmail connector.
        
        Args:
            credentials_path: Path to credentials.json file
        """
        self.credentials_path = credentials_path
        self.scopes = ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = None
        
    def authenticate(self):
        """Authenticate with Gmail API and return service object."""
        creds = None
        token_path = os.path.join(os.path.dirname(self.credentials_path), 'token.pickle')
        
        # Load existing credentials if available
        if os.path.exists(token_path):
            with open(token_path, 'rb') as token:
                creds = pickle.load(token)
                
        # If credentials don't exist or are invalid, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path, self.scopes)
                creds = flow.run_local_server(port=0)
            
            # Save credentials for future use
            with open(token_path, 'wb') as token:
                pickle.dump(creds, token)
        
        # Build the Gmail service
        self.service = build('gmail', 'v1', credentials=creds)
        return self.service
    
    def test_connection(self):
        """Test Gmail API connection."""
        if not self.service:
            self.authenticate()
        
        try:
            # Try to get profile info to verify connection
            profile = self.service.users().getProfile(userId='me').execute()
            return True, f"Connected to Gmail API. Email: {profile['emailAddress']}"
        except Exception as e:
            return False, f"Failed to connect: {str(e)}"

# Test code
if __name__ == "__main__":
    connector = GmailConnector('credentials/credentials.json')
    success, message = connector.test_connection()
    print(message) 