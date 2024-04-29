from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os

credentials_path = r"C:\Users\aryam\Downloads\client_secret_41662120290-sukvumcl65cn2e26esphffp44il4elck.apps.googleusercontent.com.json"
SCOPES = ['https://www.googleapis.com/auth/documents'] # Specifying this is allowing code to edit google docs. Might need to change cus I think I did this wrong
auth_token_path = r"C:\Users\aryam\Downloads\token.json"  # json specifying user has allowed me to access drive


def auth_google_docs_api():
    creds = None
    if os.path.exists(auth_token_path):
        print("reached")
        creds = Credentials.from_authorized_user_file(auth_token_path, SCOPES)
    if not creds or not creds.valid:
        print("reached2")
        if creds and creds.expired and creds.refresh_token:
            print("reached3")
            creds.refresh(Request())
            print("broke here")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open(auth_token_path, 'w') as token:
            token.write(creds.to_json())
    return creds


def replace_text_in_doc(document_id, text_to_replace, replacement_text):
    creds = auth_google_docs_api()
    service = build('docs', 'v1', credentials=creds)
    requests = [
        {
            'replaceAllText': {
                'containsText': {
                    'text': text_to_replace,
                    'matchCase': 'true',
                },
                'replaceText': replacement_text,
            },
        },
    ]
    result = service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
    print(f"{result.get('replies')[0].get('replaceAllText').get('occurrencesChanged')} replacements made in document {document_id}")

def delete_lines_starting_with(document_id, prefix):
    creds = auth_google_docs_api()
    service = build('docs', 'v1', credentials=creds)

    # First, retrieve the document's content
    document = service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content')

    requests = []
    runningDel = 0 # Keep track of this int so you can subtract and keep track of index
    for element in content:
        # Assuming paragraphs as primary containers
        if 'paragraph' in element:
            paragraph_elements = element.get('paragraph').get('elements')
            for elem in paragraph_elements:
                text_run = elem.get('textRun')
                if text_run and text_run.get('content').startswith(prefix):
                    # Calculate range to delete. Adjust this based on actual document structure and text segmentation.
                    start_index = elem.get('startIndex') - runningDel
                    end_index = elem.get('endIndex') - runningDel
                    runningDel += (end_index - start_index)
                    # Create a request to delete this range of text
                    requests.append({
                        'deleteContentRange': {
                            'range': {
                                'startIndex': start_index,
                                'endIndex': end_index,
                            }
                        }
                    })

    # Execute batch update to delete all identified ranges starting with the prefix
    if requests:
        result = service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print(f"{len(requests)} deletion operations performed in document {document_id}")
    else:
        print("No lines found starting with the specified prefix")

#document_id = '1P1dTdR4rMMN4M2kbKy4VcH4ld29xh0ZObDjVScZ-H7k' # id of actual text so you can identify speakers
#text_to_replace = 'SPEAKER_06' #
#replacement_text = "Attorney" # Witness (Sergeant Fox) - Speaker_05
#replace_text_in_doc(document_id, text_to_replace, replacement_text)


#document_id2 = '1xjbvgwJo6rHIFW6lkIDvnoMzb8yG8Z0JWZeKtcFQYpY' # id of dataset to remove time lines
#delete_lines_starting_with(document_id2, "time: ")

creds = auth_google_docs_api()