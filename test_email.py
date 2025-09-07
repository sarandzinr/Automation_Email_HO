import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, time
import schedule
import time as time_module
import os

class ShiftHandoverSystem:
    def __init__(self):
        # Email configuration
        self.sender_email = "sender@gmail.com" # Your gmail id
        self.app_password = "ztxc dsds sdsd dssd" # your gmail app password
        self.recipients = ["recevier1@gmail.com", "recevier2@gmail.com"] # List of recipient emails
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587
        
        # Excel file path - update this to your local file path
        self.excel_file = r"C:\Users\saran vikram\shift handover.xlsx" # your excel file path
        
        # Shift timings
        self.shifts = {
            "Morning Shift": {"start": time(6, 0), "end": time(15, 0)},
            "Afternoon Shift": {"start": time(14, 0), "end": time(23, 0)},
            "Night Shift": {"start": time(21, 0), "end": time(6, 0)}
        }
    
    def read_excel_data(self):
        """Read data from Excel file"""
        try:
            # Read the Excel file
            df = pd.read_excel(self.excel_file)
            
            # Get today's date
            today = datetime.now().strftime('%d-%m-%Y')
            
            # Filter data for today (you can modify this logic as needed)
            # For now, getting the latest 10 entries or today's entries
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
                today_data = df[df['Date'].dt.strftime('%d-%m-%Y') == today]
                
                if today_data.empty:
                    # If no today's data, get latest 5 entries
                    latest_data = df.tail(5)
                else:
                    latest_data = today_data
            else:
                # If no Date column, get latest entries
                latest_data = df.tail(10)
            
            return latest_data
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def format_email_content(self, data):
        """Format the data into email content"""
        if data is None or data.empty:
            return "No handover data available."
        
        # Create the table header
        email_body = """Hi Team,

Please find today's shift handover details below:

Date         | Shift           | Description            | Ticket No   | Assignee | Follow-up | Comments
-------------|------------------|-------------------------|-------------|----------|------------|----------------------------
"""
        
        # Add data rows
        for _, row in data.iterrows():
            date_str = row.get('Date', '').strftime('%d-%m-%Y') if pd.notna(row.get('Date', '')) else 'N/A'
            shift = str(row.get('Shift', 'N/A'))[:15].ljust(15)
            description = str(row.get('Description', 'N/A'))[:22].ljust(22)
            ticket = str(row.get('Ticket No', 'N/A'))[:11].ljust(11)
            assignee = str(row.get('Assignee', 'N/A'))[:8].ljust(8)
            followup = str(row.get('Follow-up', 'N/A'))[:9].ljust(9)
            comments = str(row.get('Comments', 'N/A'))[:28]
            
            email_body += f"{date_str.ljust(12)} | {shift} | {description} | {ticket} | {assignee} | {followup} | {comments}\n"
        
        email_body += "\nThanks,\nSaran"
        
        return email_body
    
    def send_email(self, subject, body):
        """Send email with the handover details"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(self.recipients)
            msg['Subject'] = subject
            
            # Add body to email
            msg.attach(MIMEText(body, 'plain'))
            
            # Gmail SMTP configuration
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()  # Enable security
            server.login(self.sender_email, self.app_password)
            
            # Send email
            text = msg.as_string()
            server.sendmail(self.sender_email, self.recipients, text)
            server.quit()
            
            print(f"Email sent successfully at {datetime.now()}")
            return True
            
        except Exception as e:
            print(f"Error sending email: {e}")
            return False
    
    def get_current_shift(self):
        """Determine current shift based on time"""
        current_time = datetime.now().time()
        
        # Check morning shift (6 AM to 3 PM)
        if time(6, 0) <= current_time <= time(15, 0):
            return "Morning Shift"
        # Check afternoon shift (2 PM to 11 PM)
        elif time(14, 0) <= current_time <= time(23, 0):
            return "Afternoon Shift"
        # Night shift (9 PM to 6 AM) - spans midnight
        else:
            return "Night Shift"
    
    def send_shift_handover(self):
        """Main function to send shift handover email"""
        try:
            # Read data from Excel
            data = self.read_excel_data()
            
            # Format email content
            email_body = self.format_email_content(data)
            
            # Create subject with timestamp
            current_shift = self.get_current_shift()
            timestamp = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            subject = f"Shift Handover - {current_shift} - {timestamp}"
            
            # Send email
            success = self.send_email(subject, email_body)
            
            if success:
                print("Shift handover email sent successfully!")
            else:
                print("Failed to send shift handover email.")
                
        except Exception as e:
            print(f"Error in send_shift_handover: {e}")
    
    def schedule_emails(self):
        """Schedule emails at shift end times"""
        # Schedule emails at the end of each shift
        schedule.every().day.at("15:00").do(self.send_shift_handover)  # End of morning shift
        schedule.every().day.at("23:00").do(self.send_shift_handover)  # End of afternoon shift
        schedule.every().day.at("06:00").do(self.send_shift_handover)  # End of night shift
        
        print("Email scheduler started...")
        print("Scheduled times:")
        print("- Morning Shift end: 15:00 (3:00 PM)")
        print("- Afternoon Shift end: 23:00 (11:00 PM)")
        print("- Night Shift end: 06:00 (6:00 AM)")
        
        # Keep the scheduler running
        while True:
            schedule.run_pending()
            time_module.sleep(60)  # Check every minute

# Test function
def test_system():
    """Test the system with sample data"""
    print("Testing the shift handover system...")
    
    # Create sample Excel file for testing
    sample_data = {
        'Date': ['15-06-2025', '15-06-2025', '16-06-2025'],
        'Shift': ['Morning Shift', 'Morning Shift', 'Afternoon Shift'],
        'Description': ['Disk space usage', 'Restart the instance', 'Property file update'],
        'Ticket No': ['INC259855', 'INC5484156', 'CHG245465'],
        'Assignee': ['Saran', 'Saran', 'Saran'],
        'Follow-up': ['No', 'No', 'No'],
        'Comments': ['Reduced the disk space on /log', 'Restart the instance', 'Updated the property file']
    }
    
    df = pd.DataFrame(sample_data)
    df.to_excel('shift_handover.xlsx', index=False)
    print("Sample Excel file created: shift_handover.xlsx")
    
    # Initialize and test the system
    handover_system = ShiftHandoverSystem()
    handover_system.send_shift_handover()

if __name__ == "__main__":
    # Uncomment the line below to run the test
    test_system()
    
    # Uncomment the lines below to start the scheduler
    # handover_system = ShiftHandoverSystem()
    # handover_system.schedule_emails()
