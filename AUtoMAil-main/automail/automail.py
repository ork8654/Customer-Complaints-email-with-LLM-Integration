import requests
import smtplib
import ssl
import imaplib
import time
import email
from email.message import EmailMessage
from email.parser import BytesParser
from mistralai.client import MistralClient
from mistralai.models.chat_completion import ChatMessage
import os
import random
import nltk
from nltk.tokenize import sent_tokenize
import pandas as pd
import re
from datetime import datetime

nltk.download('punkt')

# Environment variables for sensitive information
EMAIL_ADDRESS = os.environ.get('TATA_MOTORS_EMAIL', 'tatamotorscomplaintsassist@gmail.com')
EMAIL_PASSWORD = os.environ.get('TATA_MOTORS_EMAIL_PASSWORD', 'yjhk bkll kqnc mxod')
MISTRAL_API_KEY = os.environ.get('MISTRAL_API_KEY', 'YyWXDwBEHDXUURBriRwPPkFw1xlLEchi')

customer_data_file = "customer_data.csv"

# Load CSV into a DataFrame
try:
    customer_data = pd.read_csv(customer_data_file)
except FileNotFoundError:
    customer_data = pd.DataFrame(columns=[
        'Name', 'Email', 'Car Name', 'Reg No', 'Dealer', 'Area', 'Phone No',
        'Problem Area', 'Complaints', 'Complaint Status', 'Complaint raised date',
        'Action taken', 'Expected Time of Completion'
    ])

client = MistralClient(api_key=MISTRAL_API_KEY)

def send_standard_email(recipient, subject, body, html_body=None):
    msg = EmailMessage()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg['Subject'] = subject
    if html_body:
        msg.set_content(body)
        msg.add_alternative(html_body, subtype='html')
    else:
        msg.set_content(body)
    
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

def fetch_emails():
    imap_server = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    imap_server.select('INBOX')
    
    status, email_ids = imap_server.search(None, 'UNSEEN')
    email_ids = email_ids[0].split()
    
    emails = []
    
    for email_id in email_ids:
        status, email_data = imap_server.fetch(email_id, '(RFC822)')
        raw_email = email_data[0][1]
        msg = email.message_from_bytes(raw_email)
        
        email_obj = {
            'id': email_id,
            'email': msg
        }
        emails.append(email_obj)
        
    imap_server.logout()
    return emails

def mark_email_as_read(email_id):
    imap_server = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    imap_server.select('INBOX')
    imap_server.store(email_id, '+FLAGS', '\\Seen')
    imap_server.logout()

def process_emails():
    emails = fetch_emails()
    
    for email_obj in emails:
        try:
            email_message = email_obj['email']
            reply_to_email(email_message)
            mark_email_as_read(email_obj['id'])
        except Exception as e:
            print(f"Error processing email: {str(e)}")

def extract_body(email_message):
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == 'text/plain':
                return part.get_payload(decode=True).decode('utf-8')
    else:
        return email_message.get_payload(decode=True).decode('utf-8')

def generate_ticket_number():
    return f"TM{random.randint(100000000, 999999999)}"

def extract_detail(email_body, detail_type):
    patterns = {
        'reg_no': r'\b[A-Z]{2}[ -]?\d{1,2}[A-Z]{0,2}[ -]?\d{4}\b',
        'car_name': r'\b(Tata Harrier|Tata Safari|Tata Altroz|Tata Nexon|Tata Tiago|Tata Tigor|Tata Punch|Tata Tiago NRG|Tata Nexon EV|Tata Punch EV)\b',
        'dealer': r"dealer(?:ship)?[:\-]?\s*([\w\s]+)",
        'phone_no': r'\b\d{10}\b'
    }
    
    pattern = patterns.get(detail_type)
    if not pattern:
        return None
    
    match = re.search(pattern, email_body, re.IGNORECASE)
    if match:
        return match.group(0) if detail_type != 'dealer' else match.group(1).strip()
    return None

def get_mistral_response(messages):
    try:
        chat_response = client.chat(
            model="open-mixtral-8x22b",
            messages=messages,
            max_tokens=2000  # Increase max tokens to avoid truncation
        )
        return chat_response.choices[0].message.content
    except Exception as e:
        print(f"Error generating reply from Mistral AI: {str(e)}")
        return None

def save_customer_data_with_retry(data, file_path, max_retries=5):
    for attempt in range(max_retries):
        try:
            data.to_csv(file_path, index=False)
            break
        except PermissionError:
            if attempt < max_retries - 1:
                print(f"Unable to save. Retrying in a moment... (Attempt {attempt + 1}/{max_retries})")
                time.sleep(random.uniform(1, 3))
            else:
                alternative_file = f'customer_data_backup_{int(time.time())}.csv'
                print(f"Unable to save to {file_path} after {max_retries} attempts. Saving to {alternative_file} instead.")
                data.to_csv(alternative_file, index=False)

def extract_name_from_email(email_address):
    return email_address.split('@')[0]

def determine_area_from_reg_no(reg_no):
    if not reg_no or len(reg_no) < 2:
        return "Unknown"
    
    reg_code = reg_no[:2].upper()
    state_mapping = {
        "AP": "Andhra Pradesh", "AR": "Arunachal Pradesh", "AS": "Assam", "BR": "Bihar",
        "CG": "Chhattisgarh", "GA": "Goa", "GJ": "Gujarat", "HR": "Haryana",
        "HP": "Himachal Pradesh", "JH": "Jharkhand", "KA": "Karnataka", "KL": "Kerala",
        "MP": "Madhya Pradesh", "MH": "Maharashtra", "MN": "Manipur", "ML": "Meghalaya",
        "MZ": "Mizoram", "NL": "Nagaland", "OR": "Odisha", "PB": "Punjab",
        "RJ": "Rajasthan", "SK": "Sikkim", "TN": "Tamil Nadu", "TG": "Telangana",
        "TR": "Tripura", "UP": "Uttar Pradesh", "UK": "Uttarakhand", "WB": "West Bengal",
        "AN": "Andaman and Nicobar Islands", "CH": "Chandigarh", "DD": "Dadra and Nagar Haveli and Daman and Diu",
        "LD": "Lakshadweep", "DL": "Delhi", "PY": "Puducherry", "LA": "Ladakh", "JK": "Jammu and Kashmir"
    }
    return state_mapping.get(reg_code, "Unknown")

def identify_problem_area(email_body):
    problem_keywords = [
        "engine", "transmission", "brakes", "battery", "AC", "suspension",
        "breakdown", "display", "servicing", "product malfunctioning",
        "service adviser", "Part availability", "dealer service"
    ]
    for keyword in problem_keywords:
        if keyword in email_body.lower():
            return keyword.capitalize()
    return "General"

def reply_to_email(email):
    global customer_data
    msg = BytesParser().parsebytes(email.as_bytes())
    sender = msg['From']
    subject = msg['Subject']
    body = extract_body(msg)

    reg_no = extract_detail(body, 'reg_no')

    if not reg_no:
        request_missing_details(sender, ['Reg No'])
        return

    customer_record = customer_data[customer_data['Reg No'] == reg_no]

    if customer_record.empty:
        handle_new_customer(sender, body, reg_no)
    else:
        handle_existing_customer(sender, subject, body, reg_no, customer_record)

    save_customer_data_with_retry(customer_data, customer_data_file)

def handle_new_customer(sender, body, reg_no):
    car_name = extract_detail(body, 'car_name')
    dealer = extract_detail(body, 'dealer')
    phone_no = extract_detail(body, 'phone_no')
    
    missing_details = []
    if not car_name: missing_details.append('Car Name')
    if not dealer: missing_details.append('Dealer')
    if not phone_no: missing_details.append('Phone No')
    
    if missing_details:
        request_missing_details(sender, missing_details)
        return

    area = determine_area_from_reg_no(reg_no)
    problem_area = identify_problem_area(body)
    ticket_number = generate_ticket_number()

    complaint_with_ticket = f"Ticket {ticket_number}: {body}"

    name = extract_name_from_email(sender)
    new_customer_data = pd.DataFrame({
        'Name': [name],
        'Email': [sender],
        'Car Name': [car_name],
        'Reg No': [reg_no],
        'Dealer': [dealer],
        'Area': [area],
        'Phone No': [phone_no],
        'Problem Area': [problem_area],
        'Complaints': [complaint_with_ticket],
        'Complaint Status': ["Open"],
        'Complaint raised date': [datetime.now().strftime('%d-%m-%Y')],
        'Action taken': [""],
        'Expected Time of Completion': [""]
    })
    global customer_data
    customer_data = pd.concat([customer_data, new_customer_data], ignore_index=True)

    reply = compose_reply("", "New Complaint", sender, body, ticket_number, pd.DataFrame())
    send_standard_email(sender, "Re: New Complaint", "Please view this email in HTML format for the best experience.", reply)

def handle_existing_customer(sender, subject, body, reg_no, customer_record):
    customer_name = customer_record['Name'].values[0]
    complaint_status = customer_record['Complaint Status'].values[0]
    action_taken = customer_record['Action taken'].values[0]
    expected_completion = customer_record['Expected Time of Completion'].values[0]
    ticket_number = generate_ticket_number()

    close_request = "close the complaint" in body.lower() or "close the status" in body.lower()
    if close_request:
        customer_data.loc[customer_data['Reg No'] == reg_no, 'Complaint Status'] = 'Closed'
        customer_data.loc[customer_data['Reg No'] == reg_no, 'Action taken'] = 'Complaint closed as per customer request'
        customer_data.loc[customer_data['Reg No'] == reg_no, 'Expected Time of Completion'] = datetime.now().strftime('%d-%m-%Y')

    reply = compose_reply(customer_name, subject, sender, body, ticket_number, customer_record)
    send_standard_email(sender, f"Re: {subject}", "Please view this email in HTML format for the best experience.", reply)

def request_missing_details(sender, missing_details):
    details_list = ", ".join(missing_details)
    follow_up_email_body = (
        f"Dear Customer,\n\n"
        f"We noticed that some details are missing from your complaint:\n"
        f"{details_list}\n\n"
        f"Please provide these details so we can assist you better.\n\n"
        f"Thank you,\n"
        f"Tata Motors Customer Service"
    )

    send_standard_email(sender, "Additional Information Required", follow_up_email_body)

def compose_reply(greeting, subject, sender, body, ticket_number, customer_record):
    reg_no = extract_detail(body, 'reg_no')
    car_model = extract_detail(body, 'car_name') or customer_record['Car Name'].values[0] if not customer_record.empty else "your vehicle"

    # Extract name from the email body if available, otherwise use the name from the customer record
    name_match = re.search(r'Name:\s*([\w\s]+)', body)
    if name_match:
        customer_name = name_match.group(1).strip()
    elif not customer_record.empty:
        customer_name = customer_record['Name'].values[0]
    else:
        customer_name = extract_name_from_email(sender)

    if not customer_record.empty:
        complaint_status = customer_record['Complaint Status'].values[0]
        action_taken = customer_record['Action taken'].values[0]
        expected_completion = customer_record['Expected Time of Completion'].values[0]
    else:
        complaint_status = "New"
        action_taken = "Your complaint has been registered"
        expected_completion = "To be determined"

    close_request = "close the complaint" in body.lower() or "close the status" in body.lower()

    if close_request:
        predefined_prompt = (
            f"The customer with registration number {reg_no} has requested to close their complaint. "
            f"Provide a polite response acknowledging the request, confirming that the complaint has been closed, "
            f"and asking if there's anything else we can assist with. Use proper spacing between paragraphs."
        )
    else:
        predefined_prompt = (
            f"You are a customer service assistant for Tata Motors. A customer with the following details has contacted us:\n\n"
            f"- Name: {customer_name}\n"
            f"- Registration Number: {reg_no}\n"
            f"- Car Model: {car_model}\n"
            f"- Current Complaint Status: {complaint_status}\n"
            f"- Current Action Taken: {action_taken}\n"
            f"- Expected Completion Time: {expected_completion}\n\n"
            f"The customer's current message is: \"{body}\"\n\n"
            f"Please provide a professional, helpful, and thorough response. Address all reported issues individually. "
            f"If the customer is asking about the status, provide the current status, action taken, and expected completion time. "
            f"Keep the response comprehensive and maintain a polite and empathetic tone. "
            f"Use proper spacing between paragraphs and ensure a clear structure in your response. "
            f"Do not use any placeholder text like [Your Name] or [Your Position]. "
            f"End the email with 'Best regards, Tata Motors Customer Service Team'."
        )
    messages = [
        ChatMessage(role="system", content=predefined_prompt),
        ChatMessage(role="user", content=body)
    ]

    reply_content = get_mistral_response(messages)

    if not reply_content:
        reply_content = f"""
        Dear {customer_name},<br><br>
        
        Thank you for reaching out to Tata Motors Customer Service regarding your {car_model} (Registration Number: {reg_no}). We sincerely apologize for any issues you're experiencing with your vehicle and the inconvenience this has caused.<br><br>
        
        We have carefully noted your concerns and are committed to resolving them promptly. Our team will thoroughly investigate the issues you've reported and take appropriate action.<br><br>
        
        Here are the immediate steps we will take:<br><br>
        
        1. We will contact your dealership to expedite any pending service requests or part orders.<br>
        2. A senior technician specializing in your vehicle model will be assigned to investigate the reported issues.<br>
        3. We will schedule a comprehensive diagnostic test of your vehicle to ensure there are no other underlying problems.<br><br>
        
        Our customer service team will contact you within the next 24 hours to schedule these services at your convenience. We aim to have your vehicle inspected and the necessary repairs initiated within the next 3-5 business days.<br><br>
        
        Again, we deeply regret any inconvenience caused and appreciate your patience. Your satisfaction is our top priority, and we are committed to restoring your confidence in your {car_model} and our brand.<br><br>
        
        If you have any further questions or concerns in the meantime, please don't hesitate to contact our dedicated support line at 1800-209-8282.<br><br>
        
        Thank you for choosing Tata Motors. We value your trust and will do our utmost to resolve these issues swiftly.<br><br>
        
        Best regards,<br>
        Tata Motors Customer Service Team
        """

    formatted_reply = f"""
        <html>
        <head>
        <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
        }}
        .header {{
            background-color: #003366;
            color: white;
            padding: 10px;
            text-align: center;
        }}
        .content {{
            padding: 20px;
        }}
        .content p {{
            margin-bottom: 15px;
        }}
        .footer {{
            background-color: #f2f2f2;
            padding: 10px;
            text-align: center;
            font-size: 0.9em;
        }}
        </style>
        </head>
        <body>
        <div class="header">
        <h2>Tata Motors Customer Service</h2>
        <p>Ticket ID: {ticket_number}</p>
        </div>
        <div class="content">
        {reply_content.replace('\n', '<br>')}
        </div>
        <div class="footer">
        <p>For further assistance, please contact us at 1800 209 8282</p>
        <p>Thank you for choosing Tata Motors. We value your trust and support.</p>
        </div>
        </body>
        </html>
        """
    return formatted_reply


def main():
    while True:
        try:
            process_emails()
            time.sleep(20)
        except Exception as e:
            print(f"An error occurred in the main loop: {str(e)}")
            time.sleep(60)  # Wait for a minute before retrying

if __name__ == "__main__":
    main()




