"""
Extended Google Docs API Script
- Creates a new document
- Adds formatted text to the document
- Reads the document title
"""
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
# Added documents scope (not just readonly) to allow document creation and editing
SCOPES = ["https://www.googleapis.com/auth/documents"]

def get_credentials():
    """Get and refresh OAuth credentials."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    
    return creds

def create_document(service, title):
    """Create a new Google Doc with the given title."""
    body = {
        'title': title
    }
    
    try:
        document = service.documents().create(body=body).execute()
        print(f"Created document with title: {document.get('title')}")
        print(f"Document ID: {document.get('documentId')}")
        return document
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def insert_text(service, document_id, text):
    """Insert text at the beginning of the document."""
    requests = [
        {
            'insertText': {
                'location': {
                    'index': 1,  # 1 is the beginning of the document
                },
                'text': text
            }
        }
    ]
    
    try:
        result = service.documents().batchUpdate(
            documentId=document_id, 
            body={'requests': requests}
        ).execute()
        
        print(f"Text inserted successfully!")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def format_text(service, document_id, start_index, end_index, formatting):
    """Apply formatting to a range of text in the document."""
    # Formatting examples:
    # formatting = {'bold': True}
    # formatting = {'italic': True}
    # formatting = {'fontSize': {'magnitude': 12, 'unit': 'PT'}}
    # formatting = {'foregroundColor': {'color': {'rgbColor': {'red': 1.0, 'green': 0.0, 'blue': 0.0}}}}
    
    requests = [
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'textStyle': formatting,
                'fields': ','.join(formatting.keys())
            }
        }
    ]
    
    try:
        result = service.documents().batchUpdate(
            documentId=document_id, 
            body={'requests': requests}
        ).execute()
        
        print(f"Text formatting applied successfully!")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def add_heading(service, document_id, text, heading_level=1):
    """Add a heading to the document."""
    # First insert the text
    insert_text(service, document_id, text + "\n")
    
    # Get the length of the text to determine the end_index
    text_length = len(text)
    
    # Then apply the heading style
    requests = [
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': 1 + text_length
                },
                'paragraphStyle': {
                    'namedStyleType': f'HEADING_{heading_level}'
                },
                'fields': 'namedStyleType'
            }
        }
    ]
    
    try:
        result = service.documents().batchUpdate(
            documentId=document_id, 
            body={'requests': requests}
        ).execute()
        
        print(f"Heading added successfully!")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def create_table(service, document_id, rows, cols, position=1):
    """Create a table with the specified number of rows and columns."""
    requests = [
        {
            'insertTable': {
                'location': {
                    'index': position
                },
                'rows': rows,
                'columns': cols
            }
        }
    ]
    
    try:
        result = service.documents().batchUpdate(
            documentId=document_id, 
            body={'requests': requests}
        ).execute()
        
        print(f"Table created successfully!")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def insert_bullet_list(service, document_id, items):
    """Insert a bullet list with the given items."""
    requests = []
    current_index = 1  # Start at the beginning of the document
    
    for item in items:
        # Insert the text for this bullet point
        requests.append({
            'insertText': {
                'location': {'index': current_index},
                'text': item + '\n'
            }
        })
        
        # Calculate the range of the inserted text
        item_length = len(item) + 1  # +1 for newline
        
        # Add bullet formatting
        requests.append({
            'createParagraphBullets': {
                'range': {
                    'startIndex': current_index,
                    'endIndex': current_index + item_length - 1  # -1 to exclude the newline
                },
                'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
            }
        })
        
        # Update current_index for the next item
        current_index += item_length
    
    try:
        result = service.documents().batchUpdate(
            documentId=document_id, 
            body={'requests': requests}
        ).execute()
        
        print(f"Bullet list created successfully!")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def main():
    """Main function to demonstrate Google Docs API capabilities."""
    try:
        # Get credentials and build service
        creds = get_credentials()
        service = build("docs", "v1", credentials=creds)
        
        # Create a new document
        doc = create_document(service, "My Python Generated Document")
        document_id = doc.get('documentId')
        
        # Add a heading
        add_heading(service, document_id, "Welcome to My Document", 1)
        
        # Add some text
        text = "\nThis document was created using the Google Docs API and Python. " \
               "Here's an example of what you can do with programmatic document creation!\n\n"
        insert_text(service, document_id, text)
        
        # Format some text (make the first 10 characters bold)
        format_text(service, document_id, 
                   1, 25,  # The heading text range
                   {'bold': True, 'foregroundColor': {'color': {'rgbColor': {'red': 0.2, 'green': 0.4, 'blue': 0.8}}}})
        
        # Add a bullet list
        bullet_items = [
            "This is the first bullet point",
            "This is the second bullet point with some details",
            "This is the third bullet point showing the power of the API"
        ]
        insert_bullet_list(service, document_id, bullet_items)
        
        # Create a table
        create_table(service, document_id, 3, 3)
        
        print(f"\nDocument created successfully! You can view it in your Google Drive.")
        print(f"Document ID: {document_id}")
        print(f"Direct link: https://docs.google.com/document/d/{document_id}/edit")
        
    except HttpError as err:
        print(f"An error occurred: {err}")

if __name__ == "__main__":
    main()