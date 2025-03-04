import base64
import re
from datetime import datetime

class EmailProcessor:
    def __init__(self, gmail_service):
        """
        Initialize email processor.
        
        Args:
            gmail_service: Authenticated Gmail API service
        """
        self.service = gmail_service
        
    def search_emails(self, query="subject:delivery note", max_results=50):
        """
        Search for emails matching the query.
        
        Args:
            query: Search query string
            max_results: Maximum number of results to return
            
        Returns:
            List of message IDs matching the query
        """
        try:
            response = self.service.users().messages().list(
                userId='me', q=query, maxResults=max_results).execute()
            
            messages = []
            if 'messages' in response:
                messages.extend(response['messages'])
                
            while 'nextPageToken' in response:
                page_token = response['nextPageToken']
                response = self.service.users().messages().list(
                    userId='me', q=query, maxResults=max_results, 
                    pageToken=page_token).execute()
                if 'messages' in response:
                    messages.extend(response['messages'])
            
            return messages
        except Exception as e:
            print(f"Error searching emails: {str(e)}")
            return []
    
    def get_email_details(self, message_id):
        """
        Get details of a specific email.
        
        Args:
            message_id: ID of the email message
            
        Returns:
            Dictionary containing email metadata and attachment info
        """
        try:
            message = self.service.users().messages().get(
                userId='me', id=message_id, format='full').execute()
            
            headers = message['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
            sender = next((h['value'] for h in headers if h['name'].lower() == 'from'), 'Unknown')
            date_str = next((h['value'] for h in headers if h['name'].lower() == 'date'), '')
            
            # Parse date from string
            try:
                # Try to parse the date in a common format
                date_obj = datetime.strptime(date_str.split(' (')[0].strip(), '%a, %d %b %Y %H:%M:%S %z')
                date = date_obj.strftime('%Y-%m-%d')
            except:
                date = date_str
            
            # Check for attachments and inline images
            attachments = []
            
            # Function to recursively process message parts
            def process_parts(parts):
                for part in parts:
                    # Check if this part has a filename (regular attachment)
                    if part.get('filename') and part.get('filename') != '':
                        attachments.append({
                            'id': part['body'].get('attachmentId'),
                            'filename': part['filename'],
                            'mimeType': part['mimeType'],
                            'type': 'attachment'
                        })
                    
                    # Check if this part is an inline image
                    if part.get('mimeType', '').startswith('image/') and part.get('body', {}).get('attachmentId'):
                        # Generate a filename for the inline image
                        ext = part['mimeType'].split('/')[1]
                        filename = f"inline_image_{len(attachments)}.{ext}"
                        attachments.append({
                            'id': part['body'].get('attachmentId'),
                            'filename': filename,
                            'mimeType': part['mimeType'],
                            'type': 'inline'
                        })
                    
                    # Recursively process nested parts
                    if 'parts' in part:
                        process_parts(part['parts'])
            
            # Start processing from the top level
            if 'parts' in message['payload']:
                process_parts(message['payload']['parts'])
            # Handle single-part messages
            elif message['payload'].get('body', {}).get('attachmentId'):
                mime_type = message['payload'].get('mimeType', '')
                if mime_type.startswith('image/'):
                    ext = mime_type.split('/')[1]
                    filename = f"inline_image_0.{ext}"
                    attachments.append({
                        'id': message['payload']['body'].get('attachmentId'),
                        'filename': filename,
                        'mimeType': mime_type,
                        'type': 'inline'
                    })
            
            return {
                'id': message_id,
                'subject': subject,
                'sender': sender,
                'date': date,
                'attachments': attachments
            }
        except Exception as e:
            print(f"Error getting email details: {str(e)}")
            return None

# Test code
if __name__ == "__main__":
    from gmail_connector import GmailConnector
    
    connector = GmailConnector('credentials/credentials.json')
    service = connector.authenticate()
    
    processor = EmailProcessor(service)
    messages = processor.search_emails()
    
    if messages:
        for message in messages[:3]:  # First 3 messages for testing
            details = processor.get_email_details(message['id'])
            print(f"Subject: {details['subject']}")
            print(f"From: {details['sender']}")
            print(f"Date: {details['date']}")
            print(f"Attachments: {len(details['attachments'])}")
            print("---") 