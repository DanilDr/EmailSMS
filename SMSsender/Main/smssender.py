'''
Created on 07 нояб. 2014 г.

@author: Данил
'''

import os
import email
import imaplib
import sqlite3
import datetime
from urllib.parse import urlencode
from urllib.request import urlopen
from time import strftime
from email.header import decode_header
from email.utils import parsedate
try:
    import configparser
except ImportError:
    import ConfigParser as configparser

currentPath = os.path.dirname(os.path.realpath(__file__))

class CheckEmailSMS(object):
    
    config = None # кофигурация работы
    M = None # коннектор
    sqliteconn = None # SQL коннектор
    sqlitecurs = None # курсок SQLite БД
    smsurl = None
    
    def __init__(self):
        # получение конфигурации из файла
        cfgfile = os.path.join(currentPath, "config.cfg")
        self.config = configparser.ConfigParser()
        self.config.read(cfgfile, 'utf-8')
        self.__imapConnect()
        # БД SQLite
        self.sqliteconn = sqlite3.connect(os.path.join(currentPath, 'sqllite.db'))
        self.sqlitecurs = self.sqliteconn.cursor()
        # формирование URL отправки сообщений
        self.smsmesstemp = self.config['smsmessage']['message']
        
    def __imapConnect(self):
        # подключение
        config_email = self.config['email']
        if bool(config_email['usessl']):
            self.M = imaplib.IMAP4_SSL(config_email['imaphost'], int(config_email['imapport']))
        else:
            self.M = imaplib.IMAP4(config_email['imaphost'], int(config_email['imapport']))    
        self.M.login(config_email['email'], config_email['password'])
        self.M.list()
        self.M.select('inbox')
    
    def __del__(self):
        self.sqliteconn.close()
#        if self.M:
#            self.M.close()
#            self.M.logout()

    def __saveCheckMail(self, message_id):# Добавление информации и БД
        self.sqlitecurs.execute("INSERT INTO checkmails (mail_id, date) VALUES ('{message_id}', '{curtimestr}')".format(message_id=message_id, curtimestr=self.__getSQLTime()))
    
    def __checkMailDB(self, message_id):
        self.sqlitecurs.execute("SELECT * FROM checkmails WHERE mail_id='{message_id}'".format(message_id=message_id))
        mails = self.sqlitecurs.fetchone()
        if mails:
            return True # уже обработано
        else:
            return False # нет такого письма
    
    def __sendMessage(self, message_id, cfbox, message_subject, message_date):
        smsmessage = self.smsmesstemp.format(emailaddr=cfbox, emaildate=self.__converDate(message_date), 
                                             emailtitle=self.__convertSubject(message_subject))
        recipients = self.config['recipients']['recipients'].split(',')
        for recipient in recipients:
            self.__sendSMS(recipient, smsmessage, message_id)
        self.__saveCheckMail(message_id)
    
    def __sendSMS(self, recipient, messagetext, message_id):
        args = {"login" : self.config['smsc']['login'] ,
                "psw" : self.config['smsc']['pwd'],
                "phones" : recipient,
                "mes" : messagetext,
                "charset" : "utf-8"}
        urlrequest = "http://smsc.ru/sys/send.php?%s" % (urlencode(args))
        cursms = urlopen(urlrequest)
        self.__saveLog("smsc", message_id, recipient, cursms.read())
    
    def __saveLog(self, typemsg, message_id, recipient, message):
        self.sqlitecurs.execute("INSERT INTO log (type, message_id, recipient, date, message) VALUES \
        ('{type}', '{message_id}', '{recipient}', '{curtimestr}', '{message}')".format(
        type=typemsg,
        message_id=message_id,
        recipient=recipient,
        curtimestr=self.__getSQLTime(), 
        message=message.decode('utf-8')))
    
    def __getSQLTime(self):
        curtime = datetime.datetime.now()
        curtimestr = curtime.strftime("%Y-%m-%d %H:%M:%S")
        return curtimestr
    
    def __converDate(self, message_date):
        return strftime("%Y-%m-%d %H:%M:%S", parsedate(message_date))
    
    def __convertSubject(self, message_subject):
        if not message_subject[0][1]:
            return message_subject[0][0]
        else:
            return message_subject[0][0].decode(message_subject[0][1])
    
    def checkEmail(self):
        checkfrom = self.config['checkfrom']['checkfrom'] # адреса важных email
        date = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y") # проверка за последние 2 дня
        
        for cfbox in checkfrom.split(','): # проверка непрочитанных писем с важных адресов
            typ, data = self.M.search(None, '(UNSEEN FROM "{checkfrom}")'.format(checkfrom=cfbox, date=date)) # получение писем
            
            if typ == 'OK' and data[0]: # если все ОК и есть письма
                for num in data[0].decode('utf-8').split():
                    typ, data = self.M.fetch(num, '(RFC822)')
                    maildict = email.message_from_string(data[0][1].decode('utf-8'))
                    typ, data = self.M.store(num,'-FLAGS', '\Seen')
                    message_id = maildict['Message-ID']
                    message_subject = decode_header(maildict['Subject'])
                    message_date = maildict['Date']
                    checkMail = self.__checkMailDB(message_id)
                    if not checkMail:
                        self.__sendMessage(message_id, cfbox, message_subject, message_date)
#            elif data[0] == None:
#                print ("No messages")
                
        self.sqliteconn.commit()
    
if __name__ == "__main__":
    checkEmailSMS = CheckEmailSMS()
    checkEmailSMS.checkEmail()







