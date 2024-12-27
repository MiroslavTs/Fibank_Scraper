import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

url = "https://my.fibank.bg/EBank/public/offices"

response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")

offices = []
office_containers = soup.find_all("div", class_="module badge-box ng-scope")

for container in office_containers:
    try:
        name = container.find("p", {"bo-bind": "item.name"}).get_text(strip=True)
        address = container.find("p", {"bo-bind": "item.address"}).get_text(strip=True)
        phone = container.find("p", class_="grey").get_text(strip=True)
        hours = container.find("dl", class_="dl-horizontal").find_all("dd")
        sat_hours = hours[0].get_text(strip=True) if len(hours) > 0 else "N/A"
        sun_hours = hours[1].get_text(strip=True) if len(hours) > 1 else "N/A"

        if "N/A" not in (sat_hours, sun_hours):
            offices.append({
                "Име на офис": name,
                "Адрес": address,
                "Телефон": phone,
                "Раб. време събота": sat_hours,
                "Раб. време неделя": sun_hours
            })
    except Exception as e:
        print(f"Грешка при обработка на офис: {e}")

output_dir = "D:\\Programing\\Fibank_Test"
output_file = os.path.join(output_dir, "fibank_branches.xlsx")
os.makedirs(output_dir, exist_ok=True)

df = pd.DataFrame(offices)
df.to_excel(output_file, index=False)
print(f"Данните са записани в {output_file}")


def send_email():
    sender_email = "tsintsarski.work@gmail.com"
    app_password = "iyxi jsrb evlg zbrv"
    recipient_email = "db.rpa@fibank.bg"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Fibank Branches Report"

    with open(output_file, "rb") as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(output_file)}"')
        msg.attach(part)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)
            print("Имейлът е изпратен успешно.")
    except Exception as e:
        print(f"Грешка при изпращане на имейл: {e}")


send_email()
