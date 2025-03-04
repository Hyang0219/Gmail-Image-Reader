import os
import base64
from googleapiclient.errors import HttpError

class AttachmentHandler:
    def __init__(self, gmail_service, output_dir='data/attachments'):
        """
        Initialize attachment handler.
        
        Args:
            gmail_service: Authenticated Gmail API service
            output_dir: Directory to save attachments
        """
        self.service = gmail_service
        self.output_dir = output_dir
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
    
    def download_attachment(self, message_id, attachment_info):
        """
        Download an email attachment.
        
        Args:
            message_id: ID of the email message
            attachment_info: Dictionary with attachment metadata
            
        Returns:
            Path to downloaded file
        """
        try:
            att_id = attachment_info['id']
            filename = attachment_info['filename']
            mime_type = attachment_info['mimeType']
            is_inline = attachment_info.get('type') == 'inline'
            
            # Get the attachment content
            attachment = self.service.users().messages().attachments().get(
                userId='me', messageId=message_id, id=att_id).execute()
            
            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data)
            
            # Generate a unique filename to avoid overwriting
            base_name, ext = os.path.splitext(filename)
            if is_inline:
                # For inline images, use a more descriptive name
                output_path = os.path.join(self.output_dir, f"inline_{message_id[:8]}_{base_name}{ext}")
            else:
                output_path = os.path.join(self.output_dir, f"{base_name}_{message_id[:8]}{ext}")
            
            # Save the attachment
            with open(output_path, 'wb') as f:
                f.write(file_data)
                
            return {
                'path': output_path,
                'filename': filename,
                'mime_type': mime_type,
                'is_inline': is_inline
            }
        except Exception as e:
            print(f"Error downloading attachment: {str(e)}")
            return None
    
    def download_all_attachments(self, email_details):
        """
        Download all attachments from an email.
        
        Args:
            email_details: Dictionary containing email metadata
            
        Returns:
            List of paths to downloaded attachments
        """
        downloaded = []
        
        if 'attachments' in email_details and email_details['attachments']:
            for attachment in email_details['attachments']:
                result = self.download_attachment(email_details['id'], attachment)
                if result:
                    result['email_metadata'] = {
                        'sender': email_details['sender'],
                        'date': email_details['date'],
                        'subject': email_details['subject']
                    }
                    downloaded.append(result)
        
        return downloaded

# Test code
if __name__ == "__main__":
    from gmail_connector import GmailConnector
    from email_processor import EmailProcessor
    
    connector = GmailConnector('credentials/credentials.json')
    service = connector.authenticate()
    
    processor = EmailProcessor(service)
    handler = AttachmentHandler(service)
    
    messages = processor.search_emails()
    
    if messages:
        for message in messages[:1]:  # Test with first message only
            details = processor.get_email_details(message['id'])
            if details and 'attachments' in details and details['attachments']:
                downloaded = handler.download_all_attachments(details)
                print(f"Downloaded {len(downloaded)} attachments:")
                for attachment in downloaded:
                    print(f"  - {attachment['filename']} -> {attachment['path']}") 