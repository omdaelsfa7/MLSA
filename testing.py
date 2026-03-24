import pandas as pd
import smtplib
from email.message import EmailMessage

# --- CONFIGURATION ---
CSV_FILE = "Assignment 2 submission(Sheet1).csv"
SENDER_EMAIL = "@nu.edu.eg"  # Your NU email
APP_PASSWORD = "##########" # Your generated App Password
# --------------------

def live_test_to_self():
    try:
        # Load data to get the first student's name/grade for realism
        df = pd.read_csv(CSV_FILE)
        student = df.iloc[0] 
        
        print(f"🔄 Connecting to Microsoft SMTP server...")
        
        msg = EmailMessage()
        msg['Subject'] = f"TEST: Assignment 2 Grade - {student['Name']}"
        msg['From'] = f"MLSA SWE Team <{SENDER_EMAIL}>"
        msg['To'] = SENDER_EMAIL  # <--- CRITICAL: Sending to YOURSELF for the test
        
        body = (
            f"Dear {student['Name']},\n\n"
            f"This email is from the Microsoft Learn Student Ambassadors (MLSA) SWE team "
            f"regarding your recent submission for Assignment 2.\n\n"
            f"Your grade is: {student['assignment 2']}/10.5\n\n"
            f"If you have any inquiries about the grade, please contact:\n"
            f"- Omar Sholkamy: 01004366511\n"
            f"- Mohamed Ashraf: 01068989720\n\n"
            f"Best regards,\n"
            f"Microsoft Learn Student Ambassadors - SWE Team\n"
            f"Nile University"
        )
        
        msg.set_content(body)

        # The actual sending part
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()  # Secure the connection
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
            
        print(f"✅ Success! Check your inbox ({SENDER_EMAIL}) for the test email.")

    except Exception as e:
        print(f"❌ Error: {e}")
        if "AuthenticationUnsuccessful" in str(e):
            print("\n💡 Tip: Check your App Password. Your regular password won't work.")

if __name__ == "__main__":
    live_test_to_self()