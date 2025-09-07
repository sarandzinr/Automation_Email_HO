# Automation_Email_HO 
# The following steps outline the process to achieve Email Shift Handover Automation.

## 📋 Prerequisites Checklist
Before starting, ensure you have:
- Windows/Mac/Linux computer
- Internet connection
- Gmail account with 2-factor authentication enabled
- Your Excel file downloaded locally

## 🔧 Step 1: Install Python

### For Windows:
1. Go to https://www.python.org/downloads/
2. Download Python 3.8 or higher
3. IMPORTANT: Check "Add Python to PATH" during installation
4. Verify installation:
```cmd
python --version
pip --version
```

### For Mac:
```bash
# Install using Homebrew (recommended)
brew install python3
# Or download from python.org
```

### For Linux (Ubuntu/Debian):
```bash
sudo apt update
sudo apt install python3 python3-pip
```

## 📦 Step 2: Install Required Packages

Open Terminal/Command Prompt and run:
```bash
pip install pandas openpyxl schedule smtplib-ssl
```

If you get permission errors:
```bash
pip install --user pandas openpyxl schedule smtplib-ssl
```

## 📧 Step 3: Gmail App Password Setup

### 3.1 Enable 2-Factor Authentication:
1. Go to your Google Account settings
2. Security → 2-Step Verification
3. Turn it ON if not already enabled

### 3.2 Generate App Password:
1. Go to Google Account → Security
2. 2-Step Verification → App passwords
3. Select app: "Mail"
4. Select device: "Other (custom name)" → Type "Shift Handover"
5. Copy the 16-character password (e.g., "abcd efgh ijkl mnop")
6. Keep this password safe!

## 📁 Step 4: Download and Prepare Files

### 4.1 Download Your Excel File:
1. Go to your OneDrive link:
   https://1drv.ms/x/c/F21F44059DDF5076/EVBiaqF7Y6RGj1pT8wQwJEgBk2zE2rl8v-sjlVHF0YzTmA?e=lsk3Gd
2. Click "Download"
3. Save as shift_handover.xlsx in a folder (e.g., C:\ShiftHandover\)

### 4.2 Create Project Folder:
```bash
# Windows
mkdir C:\ShiftHandover
cd C:\ShiftHandover

# Mac/Linux
mkdir ~/ShiftHandover
cd ~/ShiftHandover
```

### 4.3 Create Python Files:
Create these files in your project folder:
- **File 1:** test_email.py (for immediate testing)
- **File 2:** shift_handover_main.py (for production)

## 🧪 Step 5: Test the System Immediately

### 5.1 Create Test File:
Save this as **test_email.py**:

```python
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
```

### 5.2 Run Test:
```bash
python test_email.py
```

**Expected output:**
```
🔄 Testing email system...
✅ SUCCESS! Test email sent successfully!
📧 Check these inboxes: recevier1@gmail.com, recevier2@gmail.com
```

## 🏭 Step 6: Production Setup

### 6.1 Update Configuration:
Create **config.py**:

```python
# Email Configuration
SENDER_EMAIL = "sender@gmail.com"
APP_PASSWORD = "your_actual_app_password_here"  # Replace with your app password
RECIPIENTS = ["recevier1@gmail.com", "recevier2@gmail.com"]

# File Configuration
EXCEL_FILE_PATH = "shift_handover.xlsx"  # Update with your file path

# Shift Schedule (24-hour format)
SHIFT_END_TIMES = {
    "morning": "15:00",    # 3:00 PM
    "afternoon": "23:00",  # 11:00 PM
    "night": "06:00"       # 6:00 AM
}
```

### 6.2 Create Main Script:
Save as **shift_handover_main.py** (copy from the main system artifact above)

### 6.3 Update File Paths:
In **shift_handover_main.py**, update:
```python
# Update this line with your actual Excel file path
self.excel_file = r"C:\ShiftHandover\shift_handover.xlsx"  # Windows example
# self.excel_file = "/Users/yourname/ShiftHandover/shift_handover.xlsx"  # Mac example
```

## ▶ Step 7: Run the System

### 7.1 Test with Your Excel File:
```bash
python shift_handover_main.py
```

### 7.2 Start Automatic Scheduling:
```bash
# This will run continuously and send emails at scheduled times
python shift_handover_main.py
```

## 🔄 Step 8: Run as Background Service

### For Windows:
1. **Create run_handover.bat:**
```batch
@echo off
cd /d C:\ShiftHandover
python shift_handover_main.py
pause
```

2. **Create Windows Task:**
   - Open Task Scheduler
   - Create Basic Task
   - Name: "Shift Handover Emails"
   - Trigger: At startup
   - Action: Start program → run_handover.bat

### For Mac/Linux:
1. **Create service script:**
```bash
#!/bin/bash
cd ~/ShiftHandover
python3 shift_handover_main.py
```

2. **Add to crontab:**
```bash
crontab -e
# Add this line:
@reboot /path/to/your/script.sh
```

## 📊 Step 9: Verify Excel File Format

Your Excel file should have these columns:
- **Date** (format: DD-MM-YYYY)
- **Shift** (Morning Shift/Afternoon Shift/Night Shift)
- **Description**
- **Ticket No**
- **Assignee**
- **Follow-up** (Yes/No)
- **Comments**

## 🚨 Common Issues & Solutions

### Issue 1: "Authentication failed"
**Solution:**
- Verify app password is correct
- Ensure 2FA is enabled
- Try generating new app password

### Issue 2: "No module named 'pandas'"
**Solution:**
```bash
pip install pandas openpyxl
```

### Issue 3: "Excel file not found"
**Solution:**
- Check file path is correct
- Use absolute path (e.g., `C:\ShiftHandover\shift_handover.xlsx`)
- Ensure file has correct permissions

### Issue 4: "Permission denied"
**Solution:**
```bash
# Try user installation
pip install --user pandas openpyxl schedule
```

## 🎯 Testing Checklist

- ✅ Python installed and working
- ✅ Required packages installed
- ✅ Gmail app password generated
- ✅ Test email sent successfully
- ✅ Excel file downloaded and accessible
- ✅ Main script runs without errors
- ✅ Emails received in correct format
- ✅ Scheduler running continuously

## 📞 Quick Test Commands

```bash
# Test 1: Check Python
python --version

# Test 2: Check packages
python -c "import pandas, schedule; print('All packages installed!')"

# Test 3: Send test email
python test_email.py

# Test 4: Run main system
python shift_handover_main.py
```

## 🎉 Success Indicators

✅ **System is working when you see:**
- Test email received successfully
- Console shows "Email sent successfully"
- No error messages in terminal
- Scheduler shows "Email scheduler started..."

**Need help? Run each step and let me know where you encounter issues!**
