import pandas as pd
import smtplib
import time
from email.message import EmailMessage

# --- CONFIGURATION ---
CSV_FILE = "Assignment 2 submission(Sheet1).csv"
SENDER_EMAIL = "@nu.edu.eg" 
APP_PASSWORD = "########!" 
# ---------------------

def send_all_student_grades():
    try:
        # Load the full list of 18 students
        df = pd.read_csv(CSV_FILE)
        print(f"🚀 Found {len(df)} students. Starting the mailing process...")

        # Connect to the Microsoft Server once
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, APP_PASSWORD)

            for index, row in df.iterrows():
                student_name = row['Name']
                student_email = row['Email']
                grade = row['assignment 2']

                # Create the customized email for this student
                msg = EmailMessage()
                msg['Subject'] = f"Assignment 2 Grade - {student_name}"
                msg['From'] = f"MLSA SWE Team <{SENDER_EMAIL}>"
                msg['To'] = student_email

                body = (
                    f"Dear {student_name},\n\n"
                    f"This email is from the Microsoft Learn Student Ambassadors (MLSA) SWE team "
                    f"regarding your recent submission for Assignment 2.\n\n"
                    f"Your grade is: {grade}/10\n\n"
                    f"If you have any inquiries about the grade, please contact:\n"
                    f"- Omar Sholkamy: 01004366511\n"
                    f"- Mohamed Ashraf: 01068989720\n\n"
                    f"Best regards,\n"
                    f"Microsoft Learn Student Ambassadors - SWE Team\n"
                    f"Nile University"
                )
                
                msg.set_content(body)
                
                # Send the email
                server.send_message(msg)
                print(f"✅ [{index + 1}/{len(df)}] Sent to: {student_name} ({student_email})")
                
                # Wait 2 seconds before the next one to stay under the radar
                time.sleep(2) 

        print("\n✨ Mission Accomplished! All grades have been sent.")

    except Exception as e:
        print(f"❌ An error occurred: {e}")

if __name__ == "__main__":
    send_all_student_grades()