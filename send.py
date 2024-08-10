import csv
import os
import time  # Import the time module
from settings import SENDER_EMAIL, PASSWORD, DISPLAY_NAME
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def get_msg(csv_file_path, template):
    with open(csv_file_path, 'r') as file:
        headers = file.readline().split(',')
        headers[len(headers) - 1] = headers[len(headers) - 1][:-1]
    # Open the CSV file again
    with open(csv_file_path, 'r') as file:
        data = csv.DictReader(file)
        for row in data:
            required_string = template
            for header in headers:
                value = row[header]
                required_string = required_string.replace(f'${header}', value)
            yield row['EMAIL'], required_string

def send_emails(server: SMTP, template):
    sent_count = 0

    for receiver, message in get_msg('data.csv', template):
        multipart_msg = MIMEMultipart("related")
        multipart_msg["Subject"] = message.splitlines()[0]
        multipart_msg["From"] = DISPLAY_NAME + f' <{SENDER_EMAIL}>'
        multipart_msg["To"] = receiver

        # HTML content with embedded image using cid method
        if 0:
            html_content = """
                            <html>
                                <body>
                                    <p>Hello,</p>
                                    <p>I hope this email finds you well. I wanted to share some exciting news with you.</p>
                                    <p>Here's the image:</p>
                                    <img src="cid:simple.png" alt="My Image">
                                    <p>Isn't it amazing? Feel free to reach out if you have any questions or comments.</p>
                                    <p>Best regards,<br>[Your Name]</p>
                                </body>
                            </html>
                            """
        with open("index.html", "r") as file:
            html_content = file.read()
        part2 = MIMEText(html_content, "html")
        multipart_msg.attach(part2)

        
        # Iterate over each file in the ATTACH folder
        attach_folder = 'ATTACH'
        for filename in os.listdir(attach_folder):
            if filename.endswith('.png') or filename.endswith('.jpg') or filename.endswith('.jpeg') or filename.endswith('.webp'):
                with open(os.path.join(attach_folder, filename), 'rb') as img_file:
                    img_data = img_file.read()
                    # Create a MIMEImage instance for each image
                    attach_part = MIMEImage(img_data)
                    # Set Content-ID header using the filename
                    attach_part.add_header('Content-ID', f'<{filename}>')
                    # Attach the image to the multipart message
                    multipart_msg.attach(attach_part)

        try:
            server.sendmail(SENDER_EMAIL, receiver, multipart_msg.as_string())
            sent_count += 1
            print("Email sent to : " + receiver)
            time.sleep(1)  # Introduce a 1-second delay
        except Exception as err:
            print(f'Problem occurred while sending to {receiver}')
            print(err)
            input("PRESS ENTER TO CONTINUE")

    print(f"Sent {sent_count} emails")

if __name__ == "__main__":
    host = "smtp.gmail.com"
    port = 587  # TLS replaced SSL in 1999

    with open('compose.md') as f:
        template = f.read()

    server = SMTP(host=host, port=port)
    server.connect(host=host, port=port)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(user=SENDER_EMAIL, password=PASSWORD)

    send_emails(server, template)

    server.quit()
