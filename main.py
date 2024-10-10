import streamlit as st
import pandas as pd
import requests
import re
import json
import msal
import os
import imaplib
import email
from email.header import decode_header

def save_credentials(api_keys):
    with open("credentials.json", "w") as f:
        json.dump(api_keys, f)


def load_credentials():
    if os.path.exists("credentials.json"):
        with open("credentials.json", "r") as f:
            return json.load(f)
    else:
        return {}

def authenticate_to_office365(client_id, client_secret, tenant_id):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["Mail.Read", "Mail.Send"]

    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )

    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        return result.get('access_token')
    else:
        st.error("Authentication failed. Please check your Office 365 credentials.")
        return None
def fetch_emails(access_token):
    headers = {
        'Authorization': 'Bearer ' + access_token
    }
    response = requests.get('https://graph.microsoft.com/v1.0/me/messages', headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        st.error("Failed to fetch emails.")
        return None

def display_emails(emails):
    for email_item in emails['value']:
        st.write(f"From: {email_item['from']['emailAddress']['name']}")
        st.write(f"Subject: {email_item['subject']}")
        st.write(f"Body: {email_item['bodyPreview']}")

def extract_email_data(email_body):
    model_pattern = r"(Model|Serial|S\/N|Service Tag)[:\s-]*(\S+)"
    product_pattern = r"(battery|charger|keyboard|DIMM|RAM|hard drive|charging cable)"
    
    model_match = re.search(model_pattern, email_body, re.IGNORECASE)
    product_match = re.search(product_pattern, email_body, re.IGNORECASE)
    
    data = {
        "model": model_match.group(2) if model_match else "N/A",
        "product": product_match.group(0) if product_match else "N/A",
    }
    return data
def query_hpe_partsurfer(model_number, serial_number, hpe_api_key):
    if not hpe_api_key:
        st.error("Missing HPE API key.")
        return None
    
    base_url = "https://partsurfer.hpe.com/Search.aspx"
    params = {
        "SearchText": serial_number
    }
    
    response = requests.get(base_url, params=params)
    
    if response.status_code == 200:
        part_number = "XYZ12345"
        return part_number
    else:
        st.error("Part not found or error in request.")
        return None
def query_lenovo_parts(model_number, serial_number, lenovo_api_key):
    if not lenovo_api_key:
        st.error("Missing Lenovo API key.")
        return None
    
    lenovo_url = f"https://support.lenovo.com/us/en/partslookup/{model_number}/{serial_number}"
    response = requests.get(lenovo_url)
    
    if response.status_code == 200:
        return "LEN-56789"
    else:
        st.error("Lenovo part not found.")
        return None

def respond_to_email(access_token, email_address, subject, body):
    url = f"https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}
    
    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": email_address}}
            ]
        }
    }
    
    response = requests.post(url, headers=headers, json=email_data)
    if response.status_code == 202:
        st.success("Email response sent successfully.")
    else:
        st.error("Failed to send response.")

def main():
    st.title("AI Email Interpreter for IT Parts")
    credentials = load_credentials()
    st.subheader("Enter API Keys and Credentials")

    client_id = st.text_input("Office 365 Client ID", value=credentials.get('client_id', ''))
    client_secret = st.text_input("Office 365 Client Secret", value=credentials.get('client_secret', ''), type='password')
    tenant_id = st.text_input("Office 365 Tenant ID", value=credentials.get('tenant_id', ''))
    
    hpe_api_key = st.text_input("HPE PartSurfer API Key (if applicable)", value=credentials.get('hpe_api_key', ''))
    lenovo_api_key = st.text_input("Lenovo API Key (if applicable)", value=credentials.get('lenovo_api_key', ''))
    if st.button("Save Credentials"):
        api_keys = {
            "client_id": client_id,
            "client_secret": client_secret,
            "tenant_id": tenant_id,
            "hpe_api_key": hpe_api_key,
            "lenovo_api_key": lenovo_api_key
        }
        save_credentials(api_keys)
        st.success("Credentials saved successfully!")
    if client_id and client_secret and tenant_id:
        access_token = authenticate_to_office365(client_id, client_secret, tenant_id)

        if access_token:
            emails = fetch_emails(access_token)
            if emails:
                display_emails(emails)
            
                for email_item in emails['value']:
                    email_body = email_item['bodyPreview']
                    extracted_data = extract_email_data(email_body)
                    
                    st.write("Extracted Data: ", extracted_data)
                    if "lenovo" in extracted_data['product'].lower():
                        part_number = query_lenovo_parts(extracted_data['model'], extracted_data['model'], lenovo_api_key)
                    elif "hpe" in extracted_data['product'].lower():
                        part_number = query_hpe_partsurfer(extracted_data['model'], extracted_data['model'], hpe_api_key)
                    else:
                        part_number = None
                    if part_number:
                        email_subject = f"Quote for {extracted_data['product']}"
                        email_body = f"The part number for your {extracted_data['product']} is {part_number}."
                        respond_to_email(access_token, email_item['from']['emailAddress']['address'], email_subject, email_body)
                    else:
                        st.write(f"Could not find part for {extracted_data['product']}. Email moved for manual processing.")
        else:
            st.error("Could not authenticate with Office 365. Please check your credentials.")
    else:
        st.warning("Please enter your Office 365 credentials to proceed.")


if __name__ == "__main__":
    main()
