## Title: Automated email sender
#### Description: This is app that read name,surname,email and user interest from a excel file and send an email to each person with the user interest news links every morning.
Objects:

    - ExcelFile :
        filepath: path to excel file
        getdata(): get data from excel file

    - Email:
        sender: sender email address
        receiver: receiver email address
        subject: email subject
        body: email body
        send_email(): send email

    - NewsFeeds:
        data: data from excel file
        get_news(): get news