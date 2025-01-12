#!/usr/bin/env python3

import requests
import os
import sys

UNINITIALIZED_VALUE = 'UNINITIALIZED'

# Configuration
CONFIG = {
    'BASE_URL': 'https://pdf-services-ue1.adobe.io',
    'CLIENT_ID': UNINITIALIZED_VALUE, # replace with your client ID
    'CLIENT_SECRET': UNINITIALIZED_VALUE, # replace with your client secret
}

def get_access_token(client_id, client_secret, base_url):
    print(f"Requesting access token with client_id: {client_id}")
    response = requests.post(
        url=base_url + '/token',
        headers={'Content-Type': 'application/x-www-form-urlencoded'},
        data={
            'client_id': client_id,
            'client_secret': client_secret
        }
    )
    response.raise_for_status()
    print("Access token received successfully.")
    return response.json()['access_token']

def get_upload_uri(access_token, client_id, base_url):
    print(f"Requesting upload URI with client_id: {client_id}")
    response = requests.post(
        base_url + '/assets',
        headers={
            'Authorization': f'Bearer {access_token}',
            'x-api-key': client_id,
            'Content-Type': 'application/json',
        },
        json={
            'mediaType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }
    )
    response.raise_for_status()
    data = response.json()
    print(f"Upload URI received: {data['uploadUri']}")
    return data['uploadUri'], data['assetID']

def upload_docx(upload_url, docx_filename):
    print(f"Uploading DOCX file: {docx_filename} to {upload_url}")
    with open(docx_filename, 'rb') as f:
        file_size = os.path.getsize(docx_filename)
        response = requests.put(
            upload_url,
            headers={
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Length': str(file_size)
            },
            data=f
        )
    response.raise_for_status()
    print("DOCX file uploaded successfully.")

def create_pdf(access_token, client_id, asset_id, base_url):
    print(f"Creating PDF for asset_id: {asset_id}")
    response = requests.post(
        base_url + '/operation/createpdf',
        headers={
            'Authorization': f'Bearer {access_token}',
            'x-api-key': client_id,
            'Content-Type': 'application/json'
        },
        json={
            'assetID': asset_id
        }
    )
    response.raise_for_status()
    print("PDF creation initiated.")
    return response.headers['Location']

def retrieve_pdf(access_token, client_id, location):
    print(f"Retrieving PDF from location: {location}")
    while True:
        response = requests.get(
            location,
            headers={
                'Authorization': f'Bearer {access_token}',
                'x-api-key': client_id,
            }
        )
        response.raise_for_status()
        data = response.json()
        if data['status'] == 'done':
            print("PDF is ready for download.")
            return data['asset']['downloadUri']
        elif data['status'] != 'in progress':
            raise Exception(f'Unknown status: {data["status"]}')
        print("PDF creation in progress...")

def download_pdf(download_uri, pdf_filename):
    print(f"Downloading PDF to {pdf_filename} from {download_uri}")
    response = requests.get(download_uri)
    response.raise_for_status()
    with open(pdf_filename, 'wb') as f:
        f.write(response.content)
    print("PDF downloaded successfully.")

def delete_asset(access_token, client_id, asset_id, base_url):
    print(f"Deleting asset with asset_id: {asset_id}")
    response = requests.delete(
        base_url + f'/assets/{asset_id}',
        headers={
            'Authorization': f'Bearer {access_token}',
            'x-api-key': client_id,
        }
    )
    response.raise_for_status()
    print("Asset deleted successfully.")

def main(docx_filename):
    base_url = CONFIG['BASE_URL']
    client_id = CONFIG['CLIENT_ID']
    client_secret = CONFIG['CLIENT_SECRET']

    if client_id == UNINITIALIZED_VALUE or client_secret == UNINITIALIZED_VALUE:
        raise Exception("Client ID or Secret not set")

    access_token = get_access_token(client_id, client_secret, base_url)
    upload_url, asset_id = get_upload_uri(access_token, client_id, base_url)
    upload_docx(upload_url, docx_filename)
    location = create_pdf(access_token, client_id, asset_id, base_url)
    download_uri = retrieve_pdf(access_token, client_id, location)
    pdf_filename = os.path.splitext(docx_filename)[0] + '.pdf'
    download_pdf(download_uri, pdf_filename)
    delete_asset(access_token, client_id, asset_id, base_url)
    print(f"PDF generated successfully: {pdf_filename}")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python docx2pdf.py <docx_filename>")
        sys.exit(1)
    main(sys.argv[1])
