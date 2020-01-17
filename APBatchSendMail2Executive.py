import logging
from os import path
import smtplib
import urllib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pyodbc
import pandas as pd
from sqlalchemy import create_engine
from config import MAIL_SENDER, MAIL_BODY, MAIL_SUBJECT
import codecs


class ConnectDB:
    def __init__(self):
        ''' Constructor for this class. '''
        self._connection = pyodbc.connect(
            'Driver={ODBC Driver 17 for SQL Server};Server=192.168.2.58;Database=db_iconcrm_fusion;uid=iconuser;pwd=P@ssw0rd;')
        self._cursor = self._connection.cursor()

    def query(self, query):
        try:
            result = self._cursor.execute(query)
        except Exception as e:
            logging.error('error execting query "{}", error: {}'.format(query, e))
            return None
        finally:
            return result

    def update(self, sqlStatement):
        try:
            self._cursor.execute(sqlStatement)
        except Exception as e:
            logging.error('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            self._cursor.commit()

    def exec_sp(self, sqlStatement, params):
        try:
            self._cursor.execute(sqlStatement, params)
        except Exception as e:
            logging.error('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            self._cursor.commit()

    def exec_spRet(self, sqlStatement, params):
        try:
            result = self._cursor.execute(sqlStatement, params)
        except Exception as e:
            print('error execting Statement "{}", error: {}'.format(sqlStatement, e))
            return None
        finally:
            return result

    def __del__(self):
        self._cursor.close()


def send_email(subject, message, from_email, to_email=None, attachment=None):
    """
    :param subject: email subject
    :param message: Body content of the email (string), can be HTML/CSS or plain text
    :param from_email: Email address from where the email is sent
    :param to_email: List of email recipients, example: ["a@a.com", "b@b.com"]
    :param attachment: List of attachments, exmaple: ["file1.txt", "file2.txt"]
    """
    if attachment is None:
        attachment = []
    if to_email is None:
        to_email = []
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = ", ".join(to_email)
    msg.attach(MIMEText(message, 'html'))

    for f in attachment:
        with open(f, 'rb') as a_file:
            basename = path.basename(f)
            part = MIMEApplication(a_file.read(), Name=basename)

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename
        msg.attach(part)

    email = smtplib.SMTP('apmail.apthai.com', 25)
    email.sendmail(from_email, to_email, msg.as_string())
    email.quit()
    return


def getHeaderTable(p_parm: int = None):
    strSQL = """
    EXEC [dbo].[sp_proc_mail_llbw_H] @p_parm = ?
    """
    params = (p_parm,)

    myConnDB = ConnectDB()
    result_set = myConnDB.exec_spRet(strSQL, params=params)
    returnVal = []
    for row in result_set:
        # print(row.head_subject)
        returnVal.append(row.head_subject)
        returnVal.append(row.head_curr_quarter)
        returnVal.append(row.head_curr_quarter_m1)
        returnVal.append(row.head_curr_quarter_m2)
        returnVal.append(row.head_curr_quarter_m3)
        returnVal.append(row.head_curr_qtd_w)
        returnVal.append(row.head_curr_month)
        returnVal.append(row.head_curr_month_w1)
        returnVal.append(row.head_curr_month_w2)
        returnVal.append(row.head_curr_month_w3)
        returnVal.append(row.head_curr_month_w4)
        returnVal.append(row.head_curr_month_w5)

    return returnVal


def readhtml(p_parm: int = None):
    f = None
    if p_parm == 1:
        f = codecs.open("templates/template_sec1_w5.html", 'r')
    if p_parm == 2:
        f = codecs.open("templates/template_sec2_w5.html", 'r')
    if p_parm == 3:
        f = codecs.open("templates/template_sec3_w5.html", 'r')
    return f.read()


def readhtmlw4(p_parm: int = None):
    f = None
    if p_parm == 1:
        f = codecs.open("templates/template_sec1_w4.html", 'r')
    if p_parm == 2:
        f = codecs.open("templates/template_sec2_w4.html", 'r')
    if p_parm == 3:
        f = codecs.open("templates/template_sec3_w4.html", 'r')
    return f.read()


def main():
    # params = 'Driver={ODBC Driver 17 for SQL Server};Server=192.168.2.58;Database=db_iconcrm_fusion;uid=iconuser;pwd' \
    #          '=P@ssw0rd; '
    # params = urllib.parse.quote_plus(params)
    # db = create_engine('mssql+pyodbc:///?odbc_connect=%s' % params, fast_executemany=True)

    receivers = ['suchat_s@apthai.com']
    subject = MAIL_SUBJECT
    # bodyMsg = MAIL_BODY
    bodyMsg = generateHTML()
    sender = MAIL_SENDER

    attachedFile = []

    # Send Email to Customer
    send_email(subject, bodyMsg, sender, receivers, attachedFile)


def generateHTML():

    strSQL = """
        SELECT COUNT(*) AS total_record
    	FROM dbo.BI_Calendar_Week WITH(NOLOCK)
    	WHERE 1=1
    	AND DATEPART(YEAR, StartDate) = DATEPART(YEAR,GETDATE())
    	AND DATEPART(QUARTER, StartDate) = DATEPART(QUARTER,GETDATE())
    	AND DATEPART(MONTH, StartDate) = DATEPART(MONTH,GETDATE())
        """

    myConnDB = ConnectDB()
    result_set = myConnDB.query(strSQL)
    returnVal = []
    for row in result_set:
        returnVal.append(row.total_record)

    total_week = returnVal[0]
    if total_week == 5:
        return readhtml(1) \
               .format(getHeaderTable(0)[0], getHeaderTable(0)[1], getHeaderTable(0)[2],
                       getHeaderTable(0)[3], getHeaderTable(0)[4], getHeaderTable(0)[5],
                       getHeaderTable(0)[6], getHeaderTable(0)[7], getHeaderTable(0)[8],
                       getHeaderTable(0)[9], getHeaderTable(0)[10], getHeaderTable(0)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(1)[0], getHeaderTable(1)[1], getHeaderTable(1)[2],
                       getHeaderTable(1)[3], getHeaderTable(1)[4], getHeaderTable(1)[5],
                       getHeaderTable(1)[6], getHeaderTable(1)[7], getHeaderTable(1)[8],
                       getHeaderTable(1)[9], getHeaderTable(1)[10], getHeaderTable(1)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(2)[0], getHeaderTable(2)[1], getHeaderTable(2)[2],
                       getHeaderTable(2)[3], getHeaderTable(2)[4], getHeaderTable(2)[5],
                       getHeaderTable(2)[6], getHeaderTable(2)[7], getHeaderTable(2)[8],
                       getHeaderTable(2)[9], getHeaderTable(2)[10], getHeaderTable(2)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                       getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                       getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                       getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(4)[0], getHeaderTable(4)[1], getHeaderTable(4)[2],
                       getHeaderTable(4)[3], getHeaderTable(4)[4], getHeaderTable(4)[5],
                       getHeaderTable(4)[6], getHeaderTable(4)[7], getHeaderTable(4)[8],
                       getHeaderTable(4)[9], getHeaderTable(4)[10], getHeaderTable(4)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(5)[0], getHeaderTable(5)[1], getHeaderTable(5)[2],
                       getHeaderTable(5)[3], getHeaderTable(5)[4], getHeaderTable(5)[5],
                       getHeaderTable(5)[6], getHeaderTable(5)[7], getHeaderTable(5)[8],
                       getHeaderTable(5)[9], getHeaderTable(5)[10], getHeaderTable(5)[11]) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(6)[0], getHeaderTable(6)[1], getHeaderTable(6)[2],
                       getHeaderTable(6)[3], getHeaderTable(6)[4], getHeaderTable(6)[5],
                       getHeaderTable(6)[6], getHeaderTable(6)[7], getHeaderTable(6)[8],
                       getHeaderTable(6)[9], getHeaderTable(6)[10], getHeaderTable(6)[11]) \
           + "<br />" + readhtml(2) \
               .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                       getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                       getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                       getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
           + "<br />" + readhtml(3) \
               .format(getHeaderTable(7)[0], getHeaderTable(7)[1], getHeaderTable(7)[2],
                       getHeaderTable(7)[3], getHeaderTable(7)[4], getHeaderTable(7)[5],
                       getHeaderTable(7)[6], getHeaderTable(7)[7], getHeaderTable(7)[8],
                       getHeaderTable(7)[9], getHeaderTable(7)[10], getHeaderTable(7)[11]) \
           + "<br />"
    else:
        return readhtmlw4(1) \
                   .format(getHeaderTable(0)[0], getHeaderTable(0)[1], getHeaderTable(0)[2],
                           getHeaderTable(0)[3], getHeaderTable(0)[4], getHeaderTable(0)[5],
                           getHeaderTable(0)[6], getHeaderTable(0)[7], getHeaderTable(0)[8],
                           getHeaderTable(0)[9], getHeaderTable(0)[10], getHeaderTable(0)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(1)[0], getHeaderTable(1)[1], getHeaderTable(1)[2],
                           getHeaderTable(1)[3], getHeaderTable(1)[4], getHeaderTable(1)[5],
                           getHeaderTable(1)[6], getHeaderTable(1)[7], getHeaderTable(1)[8],
                           getHeaderTable(1)[9], getHeaderTable(1)[10], getHeaderTable(1)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(2)[0], getHeaderTable(2)[1], getHeaderTable(2)[2],
                           getHeaderTable(2)[3], getHeaderTable(2)[4], getHeaderTable(2)[5],
                           getHeaderTable(2)[6], getHeaderTable(2)[7], getHeaderTable(2)[8],
                           getHeaderTable(2)[9], getHeaderTable(2)[10], getHeaderTable(2)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                           getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                           getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                           getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(4)[0], getHeaderTable(4)[1], getHeaderTable(4)[2],
                           getHeaderTable(4)[3], getHeaderTable(4)[4], getHeaderTable(4)[5],
                           getHeaderTable(4)[6], getHeaderTable(4)[7], getHeaderTable(4)[8],
                           getHeaderTable(4)[9], getHeaderTable(4)[10], getHeaderTable(4)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(5)[0], getHeaderTable(5)[1], getHeaderTable(5)[2],
                           getHeaderTable(5)[3], getHeaderTable(5)[4], getHeaderTable(5)[5],
                           getHeaderTable(5)[6], getHeaderTable(5)[7], getHeaderTable(5)[8],
                           getHeaderTable(5)[9], getHeaderTable(5)[10], getHeaderTable(5)[11]) \
               + "<br />" + readhtmlw4(1) \
                   .format(getHeaderTable(6)[0], getHeaderTable(6)[1], getHeaderTable(6)[2],
                           getHeaderTable(6)[3], getHeaderTable(6)[4], getHeaderTable(6)[5],
                           getHeaderTable(6)[6], getHeaderTable(6)[7], getHeaderTable(6)[8],
                           getHeaderTable(6)[9], getHeaderTable(6)[10], getHeaderTable(6)[11]) \
               + "<br />" + readhtmlw4(2) \
                   .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                           getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                           getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                           getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
               + "<br />" + readhtmlw4(3) \
                   .format(getHeaderTable(7)[0], getHeaderTable(7)[1], getHeaderTable(7)[2],
                           getHeaderTable(7)[3], getHeaderTable(7)[4], getHeaderTable(7)[5],
                           getHeaderTable(7)[6], getHeaderTable(7)[7], getHeaderTable(7)[8],
                           getHeaderTable(7)[9], getHeaderTable(7)[10], getHeaderTable(7)[11]) \
               + "<br />"


if __name__ == '__main__':
    main()
