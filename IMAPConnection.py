#!/usr/bin/python
# coding=UTF-8

import imaplib
import inspect
import eventlog
import email

try:
    # set connection for my mail
    inspect.getmro(imaplib.IMAP4_SSL)
    imap = imaplib.IMAP4_SSL('taomail.htc.com')
    imap.login('yuhsuan_chen@htc.com', "p@ss201601")

    # get all list from imap.list()
    print imap.list()
    '''
    for list in imap.list():
        for all in list:
            eventlog.write(all)
    '''
    #  select inbox
    print imap.select(mailbox='INBOX')

    print imap.search(None, "UNSEEN")

    mail218 = imap.fetch(218, '(RFC822)')
    print mail218
    print imap.store(218, '-FLAGS', '\\Seen')

    message = email.message_from_file(mail218)
    print message

    imap.close()
    imap.logout()
except Exception, e:
    print str(e)
    eventlog.write(str(e))
