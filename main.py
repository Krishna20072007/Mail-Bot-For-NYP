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


def send_email(subject, html_content, images, to_email):
    sender_name = "Nurture Your Pet"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    reply_to = "contact@nurtureyourpet.com"
    sender_email = "nurtureyourpet@gmail.com"  # change sender email here
    sender_password = "kooi qvni mxmb lbub"  # change app password here

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
    # add body url here
    url = "https://us4.campaign-archive.com/?e=__test_email__&u=ab8c81ebfd5310096b6de2a2a&id=bc0bb065a1"
    excel_file = "NYP test emails - 17 JULY 2024.xlsx"  # add path to emails excel

    try:

        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):
            email = row[0].strip()

            html_content = fetch_and_clean_html_content(url)

            if html_content:

                images = []
                subject = "Testing from bot - 16/07/2024"

                send_email(subject, html_content, images, email)

    except Exception as e:
        print(f"Error processing Excel file: {e}")


if __name__ == "__main__":
    main()
