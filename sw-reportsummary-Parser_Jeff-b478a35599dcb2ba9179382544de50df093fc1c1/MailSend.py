# import the smtplib module. It should be included in Python by default
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime


def SendMail():
    date_with_year = datetime.today().strftime('%Y%m%d')
    date = datetime.today().strftime('%m%d')
    content = MIMEMultipart()
    content["subject"] = F"Auto test summary CCPA/CCPB/G7800/Cisco {date_with_year}"
    content["from"] = "auto-test@looptelecom.com"  #寄件者
    to_list = ["ccw@looptelecom.com","wei-chen@looptelecom.com","jason-huang@looptelecom.com"]
    # to_list = ["jason-huang@looptelecom.com"]
    cc = ["tony-kuo@looptelecom.com","peter_lee@looptelecom.com", "gary-lin@looptelecom.com","louiswu@looptelecom.com", "zack-chang@looptelecom.com", "yalin-yang@looptelecom.com"]
    content["to"] = ", ".join(to_list)
    content["cc"] = ", ".join(cc)

    html ='<div style="font-family: Calibri, san-serif" >Dear all,<br><br>' \
    'Please see the summary report in the below path:<br>' \
    '<a href="K:\jason-huang_Data\AutoIt\Test summary">K:\jason-huang_Data\AutoIt\Test summary</a><br><br><br>'\
    'For G7800 IP/MEF8/MIB test, please go to the below link:<br>'\
    F'<a href="K:\jason-huang_Data\AutoIt\Test summary\G7800\{date}">K:\jason-huang_Data\AutoIt\Test summary\G7800\{date}</a><br><br><br>'\
    'Note: currently the unframe mode and the CAS mode do not support zero loss in BERT.<br>'\
    'thanks.<br></div>'

    content.attach(MIMEText(html,'html'))
    # set up the SMTP server
    with smtplib.SMTP(host='172.16.1.9', port=25) as smtp:
        try:
            smtp.ehlo()
            # smtp.login("jason-huang@looptelecom.com", "")  # 登入寄件者gmail
            smtp.send_message(content)  # 寄送郵件
            print("Send mail complete!")
        except Exception as e:
            print("Error message: ", e)

if __name__ == "__main__":
    SendMail()