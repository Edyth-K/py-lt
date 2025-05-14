from google_api import *
from constants import GLOBAL_DOC_ID
import time



def test_google_api():
    """Main function to demonstrate Google Docs API capabilities."""
    try:
        # Get credentials and build service
        creds = get_credentials()
        service = build("docs", "v1", credentials=creds)
        
        # Create a new document
        doc = create_document(service, "My Python Generated Document 2")
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

def main():
    """Main function to demonstrate Google Docs API capabilities."""
    try:
        # Get credentials and build service
        creds = get_credentials()
        service = build("docs", "v1", credentials=creds)
        
        # Set document ID
        document_id = GLOBAL_DOC_ID
               
        # Add some text
        # for i in range(0, 5):
        #     text = (f"\nTesting Loop Iteration. i = {i}\n\n") 
        #     append_text(service, document_id, text)
        #     time.sleep(5)
        
        print(f"\nProgram Complete!")
        print(f"Document ID: {document_id}")
        print(f"Direct link: https://docs.google.com/document/d/{document_id}/edit")
        
    except HttpError as err:
        print(f"An error occurred: {err}")


if __name__ == "__main__":
    main()