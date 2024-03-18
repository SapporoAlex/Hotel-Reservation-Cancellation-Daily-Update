from playwright.sync_api import sync_playwright
import ssl
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import gspread
from google.oauth2.service_account import Credentials
from bs4 import BeautifulSoup
import yaml


local_reservation_table_list = []
local_cancellation_table_list = []

japanese_to_english = {
    "1部屋": "1 room",
    "2部屋": "2 rooms",
    "3部屋": "3 rooms",
    "4部屋": "4 rooms",
    "5部屋": "5 rooms",
    "6部屋": "6 rooms",
    "7部屋": "7 rooms",
    "8部屋": "8 rooms",
    "9部屋": "9 rooms",
    "10部屋": "10 rooms",
    "11部屋": "11 room",
    "12部屋": "12 rooms",
    "13部屋": "13 rooms",
    "14部屋": "14 rooms",
    "15部屋": "15 rooms",
    "16部屋": "16 rooms",
    "17部屋": "17 rooms",
    "18部屋": "18 rooms",
    "19部屋": "19 rooms",
    "20部屋": "20 rooms",
    "21部屋": "21 rooms",
    "22部屋": "22 room",
    "23部屋": "23 rooms",
    "24部屋": "24 rooms",
    "25部屋": "25 rooms",
    "26部屋": "26 rooms",
    "27部屋": "27 rooms",
    "28部屋": "28 rooms",
    "29部屋": "29 rooms",
    "30部屋": "30 rooms",
    "1人": "1 person",
    "2人": "2 people",
    "3人": "3 people",
    "4人": "4 people",
    "5人": "5 people",
    "6人": "6 people",
    "7人": "7 people",
    "8人": "8 people",
    "9人": "9 people",
    "10人": "10 people",
    "11人": "11 people",
    "12人": "12 people",
    "13人": "13 people",
    "14人": "14 people",
    "15人": "15 people",
    "16人": "16 people",
    "17人": "17 people",
    "18人": "18 people",
    "19人": "19 people",
    "20人": "20 people",
    "21人": "21 people",
    "22人": "22 people",
    "23人": "23 people",
    "24人": "24 people",
    "25人": "25 people",
    "26人": "26 people",
    "27人": "27 people",
    "28人": "28 people",
    "29人": "29 people",
    "30人": "30 people",
}

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

zeniyacho_sheet_id = "**************"
zeniyacho_sheet = client.open_by_key(zeniyacho_sheet_id)
zeniyacho_password = zeniyacho_sheet.sheet1.cell(10, 9).value


zeniyacho_contract_code = "**************"
zeniyacho_account_name = "**************"
zeniyacho_neppan_site_url = "**************"
zeniyacho_screenshot_page = "#topForm > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(1)"



fukune_contract_code = "**************"
fukune_account_name = "**************"
fukune_password = "**************"
fukune_neppan_site_url = "**************"
fukune_screenshot_page = ("#topForm > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(2) >"
                          " div > table > tbody > tr:nth-child(2) > td:nth-child(2) > table > tbody > tr:nth-child(4)"
                          " > td")
data_table = ("#topForm > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(1) > div > table > "
              "tbody > tr:nth-child(2) > td:nth-child(2) > table > tbody")
cancellation_data_table = ("#topForm > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > "
                           "div > table > tbody > tr:nth-child(2) > td:nth-child(2) > table > tbody")

login_code_box = "#clientCode"
login_id_box = "#loginId"
login_password_box = "#password"
login_button = "#LoginBtn"


def delete_if_exists(filename):
    if os.path.exists(filename):
        os.remove(filename)
        print(f"Deleted existing file: {filename}")
    else:
        print(f"No existing file to delete: {filename}")


def main():
    fukune_png = "Fukune.png"
    zeniyacho_png = "Zeniyacho.png"

    delete_if_exists(fukune_png)
    delete_if_exists(zeniyacho_png)

    def screenshot_and_data(hotel_name, contract_code, account_name, password, url, screenshot):
        with sync_playwright() as p:
            page_url = url
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(page_url)
            page.is_visible(login_button)
            page.fill(login_code_box, contract_code)
            page.fill(login_id_box, account_name)
            page.fill(login_password_box, password)
            page.click(login_button)
            page.is_visible(screenshot)
            page.wait_for_timeout(10000)
            html = page.inner_html(data_table)
            soup = BeautifulSoup(html, 'html.parser')
            local_reservation_table_list = []
            data_point = 0
            for row in soup.select('tr'):
                value = row.text.strip()
                data_point += 1
                if data_point == 1:
                    translated = value.replace('本日(', 'Today (')
                    local_reservation_table_list.append(translated)
                elif data_point == 4:
                    local_reservation_table_list.append(value)
                elif data_point == 2 or data_point == 3:
                    english_values = japanese_to_english.get(value)
                    local_reservation_table_list.append(english_values)
            html = page.inner_html(cancellation_data_table)
            soup = BeautifulSoup(html, 'html.parser')
            local_cancellation_table_list = []
            data_point = 0
            for row in soup.select('tr'):
                value = row.text.strip()
                data_point += 1
                if data_point == 1:
                    translated = value.replace('本日(', 'Today (')
                    local_cancellation_table_list.append(translated)
                elif data_point == 4:
                    local_cancellation_table_list.append(value)
                elif data_point == 2 or data_point == 3:
                    english_values = japanese_to_english.get(value)
                    local_cancellation_table_list.append(english_values)
            clip = {
                'x': 50,
                'y': 80,
                'width': 580,
                'height': 200
            }
            page.screenshot(path=f'{hotel_name}.png', clip=clip)
            browser.close()
            return local_reservation_table_list, local_cancellation_table_list

    fukune_reservations, fukune_cancellations = screenshot_and_data("Fukune", fukune_contract_code,
                                                                    fukune_account_name, fukune_password,
                                                                    fukune_neppan_site_url, fukune_screenshot_page)
    fukune_png = "Fukune.png"
    print('Fukune screenshot saved')
    print('Fukune text saved')

    zeniyacho_reservations, zeniyacho_cancellations = screenshot_and_data("Zeniyacho",
                                                                          zeniyacho_contract_code,
                                                                          zeniyacho_account_name,
                                                                          zeniyacho_password,
                                                                          zeniyacho_neppan_site_url,
                                                                          zeniyacho_screenshot_page)
    zeniyacho_png = "Zeniyacho.png"
    print('Zeniyacho screenshot saved')
    print('Zeniyacho text saved')

    fukune_reservations_clean = str(fukune_reservations).replace("[", "").replace("]", "").replace("'", "")
    fukune_cancellations_clean = str(fukune_cancellations).replace("[", "").replace("]", "").replace("'", "")
    zeniyacho_reservations_clean = str(zeniyacho_reservations).replace("[", "").replace("]", "").replace("'", "")
    zeniyacho_cancellations_clean = str(zeniyacho_cancellations).replace("[", "").replace("]", "").replace("'", "")

    # Email setup
    with open("email.yaml") as f:
        content = f.read()

    email_credentials = yaml.load(content, Loader=yaml.FullLoader)
    email_sender, email_password, receiver_address_1, receiver_address_2 = (email_credentials["address"],
                                                                            email_credentials["password"],
                                                                            email_credentials["receiver_address_1"],
                                                                            email_credentials["receiver_address_2"])

    email_receivers = [receiver_address_1, receiver_address_2]
    email_bcc = ['alexanderfromaustralia@gmail.com']

    subject = "Daily update for Fukune & Zeniyacho"
    body = f"""

    <html>
        <body>
            <h2>Fukune</h2>
            <p>Reservations: {fukune_reservations_clean}<br>
            Cancellations: {fukune_cancellations_clean}<br>
               <img src="cid:Fukune">
            </p>
            <h2>Zeniyacho</h2>
            <p>Reservations: {zeniyacho_reservations_clean}<br>
            Cancellations: {zeniyacho_cancellations_clean}<br>
               <img src="cid:Zeniyacho">
            </p>
            <div style="text-align: center;">
            <img src="https://kyotomachiyas.com/wp-content/uploads/2022/10/Logo-Transparent-Logo.png" 
            style="width: 30%; height: auto;" 
            alt="Kyoto Machiya Collection">
            </div>
        </body>
    </html>

    """

    # Create a multipart message
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = email_sender
    msg['To'] = ", ".join(email_receivers)
    msg['Bcc'] = ", ".join(email_bcc)

    # Attach the HTML part
    msg.attach(MIMEText(body, 'html'))

    # Attach the images
    with open(fukune_png, 'rb') as img:
        mime = MIMEImage(img.read(), 'png')
        mime.add_header('Content-ID', '<Fukune>')
        msg.attach(mime)

    with open(zeniyacho_png, 'rb') as img:
        mime = MIMEImage(img.read(), 'png')
        mime.add_header('Content-ID', '<Zeniyacho>')
        msg.attach(mime)
    print('Png files attached')

    # Send the email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.send_message(msg, from_addr=email_sender,
                          to_addrs=email_receivers)
    print('Email sent')


if __name__ == '__main__':
    main()
