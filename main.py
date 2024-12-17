# API_Key: 29fd6b0aad7b45629e97a6d1c025a0f1

import dotenv
import requests
from dotenv import load_dotenv
import smtplib
import os
import pandas as pd

dotenv.load_dotenv(".env")
api_key = os.environ["API_KEY"]
email_sender = os.environ["EMAIL"]
email_password = os.environ["PASSWORD"]


class ExcelFile:
    def __init__(self, filepath):
        self.filepath = filepath

    def getdata(self):
        pandas_dataframe = pd.read_excel(self.filepath)
        list_dict = pandas_dataframe.set_index("name").to_dict(orient="list")
        return list_dict


class Email:
    def __init__(self, sender, password, receiver, subject, body):
        self.sender = sender
        self.password = password
        self.receiver = receiver
        self.subject = subject
        self.body = body

    def send_email(self):
        for i in range(len(self.receiver)):
            server = smtplib.SMTP("sandbox.smtp.mailtrap.io", 2525)
            server.starttls()
            server.login(self.sender, self.password)
            message = f"Subject:{self.subject}\n\n{self.body}"
            server.sendmail(self.sender, self.receiver, message)
            server.quit()


class NewsFeeds:
    def __init__(self, data):
        self.data = data

    def get_news(self):
        news_list = []
        for i in range(len(data["interest"])):
            response = requests.get(
                f"https://newsapi.org/v2/everything?qInTitle={data['interest'][i]}&from=2024-11-17&to=2024-12-17&apiKey={api_key}"
            )
            news_list.append(response.json()["articles"][i])
        return news_list


excel_file = ExcelFile(filepath="people.xlsx")
data = excel_file.getdata()

news_feeds = NewsFeeds(data)
articles = news_feeds.get_news()
for i in range(len(data["interest"])):
    news_pack = f"{articles[i]["title"]}\n{articles[i]['url']}\n\n"
    email = Email(
        email_sender,
        email_password,
        data["email"][i],
        f"Your interest news {data['interest'][i]}",
        f"{news_pack}",
    )
    print(email.body)
    email.send_email()
    print(f"Email sent to {data['surname'][i]} of the subject {data['interest'][i]}")
