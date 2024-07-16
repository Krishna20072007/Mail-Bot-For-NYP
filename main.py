import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import requests
from bs4 import BeautifulSoup
import openpyxl

# Function to fetch and clean HTML content from URL
def fetch_and_clean_html_content(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Remove unwanted div with id="awesomewrap"
            unwanted_div = soup.find('div', id='awesomewrap')
            if unwanted_div:
                unwanted_div.decompose()
            
            # Get cleaned HTML content
            cleaned_html = str(soup)
            return cleaned_html
        else:
            print(f"Error fetching content from URL. Status code: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error fetching content from URL: {e}")
        return None

# Function to send email with embedded images
def send_email(subject, html_content, images, to_email):
    sender_name = 'Nurture Your Pet'  # Update with your sender's name
    smtp_server = 'smtp.gmail.com'  # Update with your SMTP server
    smtp_port = 587  # Update with your SMTP port (if necessary)
    sender_email = 'nurtureyourpet@gmail.com'  # Update with your email address
    sender_password = 'kooi qvni mxmb lbub'  # Update with your email password
    reply_to = 'contact@nurtureyourpet.com'  # Update with your reply-to address
    
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message['From'] = f'{sender_name} <{sender_email}>'
    message['To'] = to_email
    message['Subject'] = subject
    message['Reply-To'] = reply_to
    
    # Add HTML content to email
    message.attach(MIMEText(html_content, 'html'))
    
    # Attach images as MIMEImage and set Content-ID for embedding
    for img_name, img_data in images:
        img = MIMEImage(img_data)
        img.add_header('Content-ID', f'<{img_name}>')
        img.add_header('Content-Disposition', 'inline', filename=img_name)
        message.attach(img)
    
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

def main():
    url = 'https://us4.campaign-archive.com/?e=__test_email__&u=ab8c81ebfd5310096b6de2a2a&id=bc0bb065a1'
    excel_file = 'emails.xlsx'  # Path to your Excel file with email addresses
    
    try:
        # Load Excel workbook
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active
        
        # Iterate over rows and send emails
        for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
            email = row[0].strip()  # Assuming email addresses are in column A
            
            # Fetch and clean HTML content from URL
            html_content = fetch_and_clean_html_content(url)
            
            if html_content:
                # Placeholder values for images (since links are removed)
                images = []
                subject = "Testing from bot - 16/07/2024"  # Replace with appropriate subject
                
                # Send email with fetched and cleaned HTML content and images
                send_email(subject, html_content, images, email)
    
    except Exception as e:
        print(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    main()
