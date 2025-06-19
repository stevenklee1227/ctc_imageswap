import os
import json
import logging

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# -------------- CONFIG ----------------
SCOPES = ['https://www.googleapis.com/auth/presentations',
          'https://www.googleapis.com/auth/drive']

logging.basicConfig(level=logging.INFO)

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def create_presentation(service, title):
    body = {'title': title}
    presentation = service.presentations().create(body=body).execute()
    logging.info(f'Created presentation: {presentation["presentationId"]}')
    return presentation['presentationId']

def add_slide(service, presentation_id, slide_data):
    requests = [{
        'createSlide': {
            'slideLayoutReference': {
                'predefinedLayout': 'TITLE_AND_BODY'
            }
        }
    }]
    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()

    slide_id = response['replies'][0]['createSlide']['objectId']
    logging.info(f'Added slide: {slide_id}')

    requests = [
        {
            'insertText': {
                'objectId': slide_id,
                'insertionIndex': 0,
                'text': slide_data['title']
            }
        },
        {
            'insertText': {
                'objectId': slide_id,
                'insertionIndex': 0,
                'text': slide_data['body']
            }
        },
        {
            'createImage': {
                'url': slide_data['image_url'],
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'height': {'magnitude': 300, 'unit': 'PT'},
                        'width': {'magnitude': 400, 'unit': 'PT'}
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 100,
                        'translateY': 150,
                        'unit': 'PT'
                    }
                }
            }
        }
    ]

    try:
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        logging.info(f'Slide updated with text and image: {slide_id}')
    except HttpError as error:
        logging.error(f'An error occurred: {error}')
        raise

def main():
    creds = authenticate()
    service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    with open('slide_plan.json') as f:
        slide_plan = json.load(f)

    presentation_id = create_presentation(service, "Radiology Review Automated")

    for i, slide_data in enumerate(slide_plan):
        try:
            add_slide(service, presentation_id, slide_data)
        except Exception as e:
            logging.error(f"Failed at slide {i+1}: {e}")

    drive_service.permissions().create(
        fileId=presentation_id,
        body={'role': 'reader', 'type': 'anyone'},
    ).execute()

    logging.info(f"All done! Your slides: https://docs.google.com/presentation/d/{presentation_id}/edit")

if __name__ == '__main__':
    main()
