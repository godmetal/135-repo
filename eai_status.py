# coding=<utf-8>

import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

import datetime

now = datetime.datetime.now()
latest_time = now.strftime('%Y-%m-%d') + ' ' + now.strftime('%H:%M:%S')

sender = 'platform@cjautoreport.com'

with open('recipients.txt', 'r') as f:
    recipients = []
    for line in f:
        recipients.append(line)

html = ""
debug = ""

with open('C:\\INSA\\results.txt', "rt", encoding="us-ascii") as infile:
    for line in infile:

        html = html + "<p>" + line + "</p>"
        print(line[0:8])
        if line[0:8] =="* DataEn":
            indate = line[32:42]
        elif line[0:8] =="  USER I":
            userinfo = line[30:]
        elif line[0:8] == "  DEPT I":
            deptinfo = line[30:]
        elif line[0:8] =="debug sq":
            debug = " ** debug 발생 - 점검 필요 ** "

SERVER = 'as2.cj.net:25'
SENDER_EMAIL = 'platform@cjautoreport.com'

# 수신인 등록
with open('recipients.txt', 'r') as f:
    RECEIVER_EMAIL = []
    for line in f:
        RECEIVER_EMAIL.append(line)

# 첨부파일도 포함 가능 옵션
message = MIMEBase('multipart', 'mixed')

# 본문 삽입
message.attach(MIMEText(html, 'html'))

message['Subject'] = '[EAI-STATUS] ' + debug + indate + ' User Insert/Fail : ' + userinfo + ' Dept Insert/Fail : ' + deptinfo
message['From'] = SENDER_EMAIL
message['To'] = ", ".join(RECEIVER_EMAIL)

server = smtplib.SMTP(SERVER)
server.ehlo()
server.starttls()
# server.login(SENDER_EMAIL, SENDER_PASSWORD) 인증 필요 없음
server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, message.as_string())
server.quit()