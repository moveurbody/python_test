#!/usr/bin/python
# coding=UTF-8

import os
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib

# Setup log format,level and path
logging.basicConfig(level=logging.DEBUG,
                    filename='debug.log',
                    format='%(asctime)s %(levelname)s %(message)s',
                    datefmt='%y-%m-%d %H:%M:%S')


def send_mail(mail_from, mail_to, mail_subject, mail_body, mail_attachment=None):
    msg = MIMEMultipart()
    # Mail information
    msg['from'] = mail_from
    # Combine the multiple recipients to a string for MIMEText
    msg['to'] = ";".join(mail_to)
    msg['subject'] = mail_subject
    # Mail body
    text = MIMEText(mail_body, 'html')
    msg.attach(text)

    # Attachment
    if mail_attachment is None:
        logging.debug("No attachment")
    else:
        folder_path, file_name = os.path.split(mail_attachment)
        att1 = MIMEText(open(mail_attachment, 'rb').read(), 'base64', 'gb2312')
        att1["Content-Type"] = 'application/octet-stream'
        att1["Content-Disposition"] = 'attachment; filename='+file_name
        msg.attach(att1)
    # Send mail

    try:
        server = smtplib.SMTP()
        server.connect('10.122.128.28')
        logging.info("connect to server....")
        # server.connect('msrelay.htc.com')
        # account and password
        # server.login('PowerExpertSystem@htc.com', '!264h59P')
        # SMTP receiver must be a list, so use mail_to
        server.sendmail(msg['from'], mail_to, msg.as_string())
        server.quit()
        logging.info("Mail is sent")
    except Exception, e:
        print str(e)
        logging.warning(e)