import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import requests
from bs4 import BeautifulSoup
import openpyxl


def fetch_and_clean_html_content(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, "html.parser")

            unwanted_div = soup.find("div", id="awesomewrap")
            if unwanted_div:
                unwanted_div.decompose()

            cleaned_html = str(soup)
            return cleaned_html
        else:
            print(
                f"Error fetching content from URL. Status code: {
                    response.status_code}"
            )
            return None
    except Exception as e:
        print(f"Error fetching content from URL: {e}")
        return None


def send_email(subject, html_content, images, to_email, sender_email, sender_password):
    sender_name = "Nurture Your Pet"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    reply_to = "contact@nurtureyourpet.com"

    message = MIMEMultipart()
    message["From"] = f"{sender_name} <{sender_email}>"
    message["To"] = to_email
    message["Subject"] = subject
    message["Reply-To"] = reply_to

    message.attach(MIMEText(html_content, "html"))

    for img_name, img_data in images:
        img = MIMEImage(img_data)
        img.add_header("Content-ID", f"<{img_name}>")
        img.add_header("Content-Disposition", "inline", filename=img_name)
        message.attach(img)

    try:
        context = smtplib.SMTP(smtp_server, smtp_port)
        context.starttls()
        context.login(sender_email, sender_password)

        context.sendmail(sender_email, to_email, message.as_string())
        print(f"Email sent successfully to {to_email}")

    except Exception as e:
        print(f"Error: Unable to send email. {e}")

    finally:
        context.quit()


def main():
    # url and path to excel file containing email ids needs to be changed everytime
    url = "https://us22.campaign-archive.com/?e=__test_email__&u=3ea6479ffa9fd8f9d056f5bd1&id=80159c0d47"
    excel_file = "NYP IDs - 17 July 2024 -3.xlsx"

    sender_email = input("Enter sender's email: ")
    sender_password = input("Enter app password: ")

    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
            email = row[0].strip()

            html_content = fetch_and_clean_html_content(url)

            if html_content:
                images = []
                subject = "Honoring Your Pet's Final Moments: A Compassionate Guide for Loving Pet Parents"

                send_email(
                    subject, html_content, images, email, sender_email, sender_password
                )

    except Exception as e:
        print(f"Error processing Excel file: {e}")


if __name__ == "__main__":
    main()
