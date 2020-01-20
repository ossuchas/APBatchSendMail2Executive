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
from config import MAIL_SENDER,  EST_MAIL_SUBJECT, EST_MAIL_BODY, EST_MAIL_TABLE_HEAD, EST_MAIL_TABLE_COL_HEAD
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


def generateHTML():
    strSQL = """
    SELECT subject_name, qtd_curr_w_ac, qtd_curr_w_diff, qtg_val, curr_m_w1_ac 
	FROM dbo.crm_mail_ll_data WHERE subject_no = 1 ORDER BY tran_id
    """

    myConnDB = ConnectDB()
    result_set = myConnDB.query(strSQL)

    strHTML = ""
    str_tmpHTML = """
    <tr>
    <td align="left"><strong>{0}</strong></td>
    <td align="right" >{1}&nbsp;</td>
    <td align="right">{2}&nbsp;</td>
    <td align="right">{3}&nbsp;</td>
    <td align="right">{4}&nbsp;</td>
    </tr>
    """

    for row in result_set:
        # print(row.subject_name.strip(), row.qtd_curr_w_ac, row.qtd_curr_w_diff, row.qtg_val, row.curr_m_w1_ac )
        strHTML = strHTML.strip() + str_tmpHTML.format(row.subject_name.strip(), f"{row.qtd_curr_w_ac:,.0f}",
                                                       f"{row.qtd_curr_w_diff:,.0f}",
                                                       f"{row.qtg_val:,.0f}", f"{row.curr_m_w1_ac:,.0f}")

    return """
    <body>{0}{1}{2}</body>
    """.format(EST_MAIL_TABLE_HEAD, EST_MAIL_TABLE_COL_HEAD, strHTML)


def readHTMLFile(p_parm: int = None):
    f = None
    if p_parm == 1:
        f = codecs.open("templates/template_est_rev.html", 'r')
    if p_parm == 2:
        f = codecs.open("templates/template_est_rev_byproj.html", 'r')

    return f.read()


def main():
    # receivers = ['suchat_s@apthai.com', 'apichaya@apthai.com', 'jintana_i@apthai.com', 'polwaritpakorn@apthai.com']
    receivers = ['suchat_s@apthai.com']
    subject = EST_MAIL_SUBJECT
    # bodyMsg = generateHTML()
    # bodyMsg = EST_MAIL_BODY
    bodyMsg = "{0} <br /> {1}".format(readHTMLFile(1), readHTMLFile(2))
    sender = MAIL_SENDER

    attachedFile = []

    # Send Email to Customer
    send_email(subject, bodyMsg, sender, receivers, attachedFile)


if __name__ == '__main__':
    main()
