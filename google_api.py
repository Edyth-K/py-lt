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
    

def append_text(service, document_id, text):
    """Append text to the end of the document."""
    # First get the current document to find its end
    document = service.documents().get(documentId=document_id).execute()
    
    # Find the end of the document
    doc_content = document.get('body').get('content')
    end_index = doc_content[-1].get('endIndex', 1)

    # Now append text at that location
    requests = [
        {
            'insertText': {
                'location': {
                    'index': end_index - 1  # -1 because the last index is typically after the final paragraph mark
                },
                'text': text
            }
        }
    ]
    
    result = service.documents().batchUpdate(
        documentId=document_id, 
        body={'requests': requests}
    ).execute()
    
    return result

def get_document_content(service, document_id):
    """
    Retrieves the text content from a Google Document and returns it in a format
    that can be used to insert into another document.
    
    Args:
        service: Authenticated Google Docs service instance
        document_id: ID of the document to retrieve content from
        
    Returns:
        A list of structural elements (paragraphs, tables, etc.) in a format 
        compatible with document.batchUpdate() requests
    """
    try:
        # Get the complete document content
        document = service.documents().get(documentId=document_id).execute()
        
        # Extract the document content
        doc_content = document.get('body', {}).get('content', [])
        
        # Format the content for insertion into another document
        insert_requests = []
        
        # Process each structural element (paragraphs, tables, etc.)
        for element in doc_content:
            # Skip the endOfSegmentMarker which appears at the end of documents
            if 'endIndex' in element and 'startIndex' in element:
                # Handle paragraph elements
                if 'paragraph' in element:
                    paragraph = element['paragraph']
                    
                    # Extract the text and text styles from paragraph elements
                    paragraph_style = paragraph.get('paragraphStyle', {})
                    elements = paragraph.get('elements', [])
                    
                    # For each text run in the paragraph
                    for text_element in elements:
                        if 'textRun' in text_element:
                            text_run = text_element['textRun']
                            content = text_run.get('content', '')
                            text_style = text_run.get('textStyle', {})
                            
                            # Skip empty content
                            if not content or content == '\n':
                                continue
                                
                            # Create an insertText request
                            insert_text = {
                                'insertText': {
                                    'text': content,
                                    'endOfSegmentLocation': {
                                        'segmentId': ''
                                    }
                                }
                            }
                            
                            insert_requests.append(insert_text)
                            
                            # If there's styling to apply, create an updateTextStyle request
                            if text_style:
                                # Need to reference the range we just inserted
                                # This will need adjustment in actual implementation
                                pass
                
                # Handle tables (simplified - you may need more complex handling)
                elif 'table' in element:
                    # Process table content
                    table = element['table']
                    for row in table.get('tableRows', []):
                        for cell in row.get('tableCells', []):
                            for cell_content in cell.get('content', []):
                                if 'paragraph' in cell_content:
                                    # Process paragraphs within table cells
                                    # Similar to above paragraph processing
                                    pass
                
                # Handle other element types as needed
                # (lists, images, etc.)
        
        return insert_requests
    
    except Exception as e:
        print(f"Error retrieving document content: {e}")
        return []