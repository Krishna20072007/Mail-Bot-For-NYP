import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import win32com.client as win32

# Function to read .msg file and extract subject, HTML content, images, and links
def read_msg_file(file_path):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(file_path)
    
    subject = msg.Subject
    
    html_content = None
    images = []
    links = []
    
    # Extract HTML body content
    html_body = msg.HTMLBody
    if html_body:
        html_content = html_body
    
    # Extract images and links
    for attachment in msg.Attachments:
        if attachment.Type == 1:  # Type 1 corresponds to embedded images
            img_name = attachment.FileName
            img_data = attachment.Content
            images.append((img_name, img_data))
        elif attachment.Type == 5:  # Type 5 corresponds to attached files (including links)
            link_name = attachment.FileName
            links.append(link_name)
    
    return subject, html_content, images, links

# Function to send email with embedded images and links
def send_email(subject, html_content, images, links, to_email):
    smtp_server = 'smtp.gmail.com'  # Update with your SMTP server
    smtp_port = 587  # Update with your SMTP port (if necessary)
    sender_email = 'kamalmahek1610@gmail.com'  # Update with your email address
    sender_password = 'hhsd hawx coib qtqy'  # Update with your email password
    
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = to_email
    message['Subject'] = subject
    
    # Add HTML content to email
    message.attach(MIMEText(html_content, 'html'))
    
    # Attach images as MIMEImage and set Content-ID for embedding
    for img_name, img_data in images:
        img = MIMEImage(img_data)
        img.add_header('Content-ID', f'<{img_name}>')
        img.add_header('Content-Disposition', 'inline', filename=img_name)
        message.attach(img)
    
    # Add links as plain text in the email body
    if links:
        html_content += "<p>Links:</p>"
        for link in links:
            html_content += f"<p>{link}</p>"
    
    try:
        # Create a secure SSL context
        context = smtplib.SMTP(smtp_server, smtp_port)
        context.starttls()
        
        # Login to server
        context.login(sender_email, sender_password)
        
        # Send email
        context.sendmail(sender_email, to_email, message.as_string())
        print(f"Email sent successfully to {to_email}")
        
    except Exception as e:
        print(f"Error: Unable to send email. {e}")
        
    finally:
        context.quit()

if __name__ == "__main__":
    msg_file_path = 'c:\Users\kamal\Desktop\Mail-Bot\1.msg'  # Path to your .msg file
    recipient_email = 'krishnakashab@gmail.com'  # Recipient's email address
    
    # Read .msg file and extract subject, HTML content, images, and links
    subject, html_content, images, links = read_msg_file(msg_file_path)
    
    # Send email with embedded images and links
    send_email(subject, html_content, images, links, recipient_email)
