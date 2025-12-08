import gspread
from google.oauth2.service_account import Credentials
import qrcode
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import io
import pandas as pd
import traceback 
GOOGLE_SHEET_NAME = "Event Registration (Responses)" 
# SENDER_EMAIL = "ieee@ejust.edu.eg"
SENDER_PASSWORD = "iagb xgux cqgh xdgs" 

EMAIL_COLUMN = "Email address"
NAME_COLUMN = "Name"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def connect_to_google_sheet():
    """Connect to Google Sheet with detailed debugging"""
    print("... Attempting to connect to Google API ...")
    try:
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        client = gspread.authorize(creds)
        
        try:
            sheet = client.open(GOOGLE_SHEET_NAME).sheet1
            print(f"✓ Successfully connected to: {GOOGLE_SHEET_NAME}")
            return sheet
        except gspread.SpreadsheetNotFound:
            print(f"\nX ERROR: Could not find a spreadsheet named '{GOOGLE_SHEET_NAME}'")
            print("  Listing all spreadsheets this bot can see:")
            try:
                available_sheets = client.list_spreadsheet_files()
                if not available_sheets:
                    print("  [NONE] The bot cannot see ANY sheets. Did you share the sheet with the bot's email?")
                else:
                    for s in available_sheets:
                        print(f"  - Found: '{s['name']}' (ID: {s['id']})")
                print("\n  Double check your GOOGLE_SHEET_NAME and the Share settings.")
            except Exception as list_err:
                print(f"  (Could not list files: {list_err} - Check if Google Drive API is enabled in Cloud Console)")
            return None

    except Exception as e:
        print("\nX CRITICAL CONNECTION ERROR:")
        traceback.print_exc()
        return None

def generate_qr_code(data):
    """Generate QR code"""
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    return img_byte_arr

def send_email_with_qr(recipient_email, recipient_name, qr_image):
    """Send email with QR code"""
    msg = MIMEMultipart('related')
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg['Subject'] = "Welcome to TechXChange '25 at E-JUST – Get Ready!"
    
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; padding: 20px;">
        <h2>Dear {recipient_name},</h2>
        <p>We are excited to welcome you to <strong>TechXChange 2025</strong>, taking place at <strong>Egypt-Japan University of Science and Technology (E-JUST)</strong>. Get ready for a one-of-a-kind experience filled with inspiring talks, hands-on workshops, competitions, and valuable networking opportunities!
<p> </p>
<p><strong>To confirm your attendance</strong>, please scan the <strong>QR code below</strong> to register your entry.</p>

<div><strong>📌 Important Notes:</strong></div>
<div>-<strong> Kindly save the QR code after scanning</strong>, as it will be required for entry at the event gate.</div>
<div>- Do <strong>not share your QR code</strong> with others; it is unique to you.</div>
<div>- Ensure you <strong>arrive early</strong> to complete the check-in smoothly.</div>
- In case of any issue, please contact us.</p>
        <div style="margin: 20px 0; padding: 20px; background-color: #f5f5f5; text-align: center;">
            <img src="cid:qrcode" alt="Your QR Code" style="width: 300px; height: 300px;">
        </div>
        <p>We can't wait to see you at TechXChange ’25 and make it an unforgettable experience together!

    <div>Best regards,<div>  
    <div>Esraa Elhossieny</div>  
    <div>IEEE CS EJUST SBC Chair</div>  
    <lable>ieeecsejust@gmail.com</p>
    </body>
    </html>
    """
    
    msg.attach(MIMEText(body, 'html'))
    img = MIMEImage(qr_image.read())
    img.add_header('Content-ID', '<qrcode>')
    msg.attach(img)
    
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"  X Failed to send to {recipient_email}: {str(e)}")
        return False

def main():
    """Send QR codes to all attendees"""
    print("\n" + "=" * 60)
    print(" SENDING QR CODES TO ALL ATTENDEES")
    print("=" * 60 + "\n")
    
    sheet = connect_to_google_sheet()
    if not sheet:
        print("\n⚠️ STOPPING: Please fix the connection error above.")
        return
    
    try:
        records = sheet.get_all_records()
        df = pd.DataFrame(records)
        print(f"✓ Found {len(df)} attendees\n")
    except Exception as e:
        print(f"X Error loading data from sheet: {str(e)}")
        traceback.print_exc()
        return
    

    df['unique_id'] = df[EMAIL_COLUMN].astype(str)
    
    df['row_number'] = df.index + 2
    
    success = 0
    for index, row in df.iterrows():
        name = row[NAME_COLUMN]
        email = row[EMAIL_COLUMN]
        
        qr_content = email 
        
        print(f"[{index + 1}/{len(df)}] Sending to {name} ({email})...", end=" ")
        
        qr_image = generate_qr_code(qr_content)
        
        if send_email_with_qr(email, name, qr_image):
            print("✓")
            success += 1
        else:
            print("X")
    
    df[['unique_id', 'row_number', NAME_COLUMN, EMAIL_COLUMN]].to_csv('qr_mapping.csv', index=False)
    
    print(f"\n{'=' * 60}")
    print(f" SUMMARY: {success}/{len(df)} emails sent successfully")
    print(f" MAPPING FILE: 'qr_mapping.csv' created.")
    print(f"{'=' * 60}\n")

if __name__ == "__main__":
    main()