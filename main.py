import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO  # Import BytesIO for in-memory file handling

# Function to create a sample Excel file
def create_sample_excel():
    sample_data = {
        "Name": ["John Doe", "Jane Smith", "Alice Johnson"],
        "Email": ["john.doe@example.com", "jane.smith@example.com", "alice.johnson@example.com"],
        "Company": ["Tech Corp", "Innovate Inc", "Data Solutions"]
    }
    df = pd.DataFrame(sample_data)
    return df

# Function to send bulk emails
def send_bulk_emails(df, start_row, end_row, sender_email, sender_password, subject, body):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    
    subset_df = df.iloc[start_row:end_row+1]
    
    for index, row in subset_df.iterrows():
        email_body = body.format(hr_name=row["Name"], company_name=row["Company"])
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = row["Email"]
        msg['Subject'] = subject
        msg.attach(MIMEText(email_body, 'plain'))
        
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, row["Email"], msg.as_string())
            server.quit()
            st.success(f"Email sent to {row['Email']}")
        except smtplib.SMTPException as e:
            st.error(f"Failed to send email to {row['Email']}: {str(e)}")
        except Exception as e:
            st.error(f"An unexpected error occurred while sending email to {row['Email']}: {str(e)}")

# Streamlit App
st.title("📧 Bulk Gmail Sender")

# Note about Excel file format
st.write("**Note:** Your Excel file must contain the following columns: `Name`, `Email`, `Company`.")

# Create and Download Sample Excel File
sample_df = create_sample_excel()

# Use BytesIO to create an in-memory Excel file
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    sample_df.to_excel(writer, index=False)
output.seek(0)  # Reset the pointer to the beginning of the file

# Download Button for Sample Excel File
st.download_button(
    label="Download Sample Excel File",
    data=output,
    file_name="sample_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Gmail Credentials Section
st.write("### Gmail Credentials")
st.write("""
To send emails, you need to provide your **Gmail address** and an **App Password**.
- **Example Gmail Address:** `your.email@gmail.com`
- **Example App Password:** `abcd efgh ijkl mnop` (16 characters, no spaces)
""")

# Instructions for Generating App Password
with st.expander("How to Generate a Gmail App Password"):
    st.write("""
1. Go to your [Google Account](https://myaccount.google.com/).
2. Enable **2-Step Verification** (if not already enabled).
3. Under **Security**, find **App Passwords**.
4. Select **App** as `Mail` and **Device** as `Other`.
5. Generate the password and use it in this app.
    """)

# Upload Excel File
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Preview of Uploaded Data")
    st.dataframe(df.head())
    
    # Gmail Credentials Input
    sender_email = st.text_input("Your Gmail Address", placeholder="your.email@gmail.com")
    sender_password = st.text_input("App Password", type="password", placeholder="abcd efgh ijkl mnop")
    
    # Email Content
    subject = st.text_input("Email Subject", "Let's Connect: Empowering Fresh Talent For Your Hiring Needs!")
    body = st.text_area("Email Body", '''
Hi {hr_name},

I hope you’re having a great day!

I’m Mahak Chugh from Veridical Technologies and we’re passionate about transforming fresh graduates into skilled professionals. With over 15 years of experience in training students across key IT domains like Cybersecurity, Data Analysis, Machine Learning, Web Development, Digital Marketing, and more, we pride ourselves on providing the knowledge and hands-on training that companies like yours need.

It would be great to work with {company_name} to help bridge the gap between talented students and meaningful job opportunities. Our goal is to deliver motivated, well-trained candidates who are ready to hit the ground running and add value to your team right away.

We are keen to have a discussion about any current or upcoming hiring needs and how we can support your recruitment goals and help streamline your hiring process.

Looking forward to hearing from you!

Best regards,
Mahak Chugh
Placement Coordinator
Veridical Technologies
''')
    
    # Row Selection
    start_row = st.number_input("Start Row", min_value=0, max_value=len(df)-1, value=0)
    end_row = st.number_input("End Row", min_value=start_row, max_value=len(df)-1, value=len(df)-1)
    
    # Send Emails Button
    if st.button("Send Emails"):
        if sender_email and sender_password:
            send_bulk_emails(df, start_row, end_row, sender_email, sender_password, subject, body)
        else:
            st.error("Please enter your Gmail credentials.")