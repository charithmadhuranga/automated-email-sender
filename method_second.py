import smtplib
import pandas as pd
import requests
import pprint
import os
import dotenv

dotenv.load_dotenv(".env")




class NewsFeeds:
    def __init__(self, interest,base_url,date_from,date_to,api_key):
        self.interest = interest
        self.base_url = base_url
        self.date_from = date_from
        self.date_to = date_to
        self.api_key = api_key

    def get_news(self):

        for j in range(len(self.interest)):
            response = requests.get(
                f"{self.base_url}"
                f"qInTitle={self.interest[j]}&"
                f"from={self.date_from}&"
                f"to={self.date_to}&"
                f"apiKey={self.api_key}"
            )
            articles_all = response.json()["articles"]
            email_body = ""
            for article in articles_all:
                email_body = email_body + article["title"]+ "\n"+ article["url"]+"\n\n"
            return email_body


class ExcelFile:
    def __init__(self, filepath):
        self.filepath = filepath

    def getdata(self):
        pandas_dataframe = pd.read_excel("people.xlsx")
        pandas_dataframe = pandas_dataframe.reset_index(drop=True)
        pandas_dataframe['ID'] = pandas_dataframe.index
        list_dict = pandas_dataframe.set_index("ID").to_dict(orient="list")
        return list_dict

class Email:
    def __init__(self, sender, password, receiver, subject, email_content,smtp_server,smtp_port):
        self.sender = sender
        self.password = password
        self.receiver = receiver
        self.subject = subject
        self.email_content = email_content
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def send_email(self):
        for i in range(len(self.receiver)):
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.sender, self.password)
            server.sendmail(self.sender,self.receiver,self.email_content)
            server.quit()


api_key = os.environ["API_KEY"]
email_sender = os.environ["EMAIL"]
email_password = os.environ["PASSWORD"]
base_url = "https://newsapi.org/v2/everything?"
date_from = "from=2024-11-17"
date_to = "to=2024-11-27"

excel_file = ExcelFile(filepath="people.xlsx")
list_dict = excel_file.getdata()
interest = list_dict["interest"]
email_receivers = list_dict["email"]
email_subject = list_dict["interest"]
smtp_server ="sandbox.smtp.mailtrap.io"
smtp_port = 2525
news_feeds = NewsFeeds(interest,base_url,date_from,date_to,api_key)



for i in range(len(list_dict)):
    email_content = (f"hi {list_dict["name"][i]},\nThis is the News you are interested in {interest[i]},"
                     f"\n\n{news_feeds.get_news()}")
    print(email_content)
    email = Email(email_sender, email_password,
                      email_receivers, email_subject,
                      email_content, smtp_server, smtp_port)
    email.send_email()









