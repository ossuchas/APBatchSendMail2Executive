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


def generateHTMLDetl():
    strSQL = """
    SELECT seqn_no, trns_name, ac_q1, ac_q2, ac_q3, ac_q4, qtd_curr_ac, est_curr_d1, est_curr_d2, est_curr_d3,
    est_curr_d4, est_curr_d5, qtd_est_total, ytd_est_ac
    FROM dbo.crm_mail_est_reve
    ORDER BY seqn_no
    """

    myConnDB = ConnectDB()
    result_set = myConnDB.query(strSQL)

    strHTML = ""
    str_tmpHTML = """
    <tr>
        <td align="left">{0}</td>
        <td align="right">{1}&nbsp;</td>
        <td align="right">{2}&nbsp;</td>
        <td align="right">{3}&nbsp;</td>
        <td align="right">{4}&nbsp;</td>
        <td align="right">{5}&nbsp;</td>
        <td align="right">{6}&nbsp;</td>
        <td align="right">{7}&nbsp;</td>
        <td align="right">{8}&nbsp;</td>
        <td align="right">{9}&nbsp;</td>
        <td align="right">{10}&nbsp;</td>
        <td align="right">{11}&nbsp;</td>
        <td align="right">{12}&nbsp;</td>
    </tr>
    """

    str_total_tmpHTML = """
    <tr>
        <td align="right" bgcolor="#FFFFCC"><strong>{0}</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{1}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{2}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{3}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{4}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{5}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{6}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{7}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{8}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{9}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{10}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{11}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{12}&nbsp;</strong></td>
    </tr>
    """

    for row in result_set:
        if row.seqn_no != 3:
            strHTML = strHTML.strip() + str_tmpHTML.format(row.trns_name.strip(),
                                                           f"{row.ac_q1:,.0f}",
                                                           f"{row.ac_q2:,.0f}",
                                                           f"{row.ac_q3:,.0f}",
                                                           f"{row.ac_q4:,.0f}",
                                                           f"{row.qtd_curr_ac:,.0f}",
                                                           f"{row.est_curr_d1:,.0f}",
                                                           f"{row.est_curr_d2:,.0f}",
                                                           f"{row.est_curr_d3:,.0f}",
                                                           f"{row.est_curr_d4:,.0f}",
                                                           f"{row.est_curr_d5:,.0f}",
                                                           f"{row.qtd_est_total:,.0f}",
                                                           f"{row.ytd_est_ac:,.0f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.trns_name.strip(),
                                                                 f"{row.ac_q1:,.0f}",
                                                                 f"{row.ac_q2:,.0f}",
                                                                 f"{row.ac_q3:,.0f}",
                                                                 f"{row.ac_q4:,.0f}",
                                                                 f"{row.qtd_curr_ac:,.0f}",
                                                                 f"{row.est_curr_d1:,.0f}",
                                                                 f"{row.est_curr_d2:,.0f}",
                                                                 f"{row.est_curr_d3:,.0f}",
                                                                 f"{row.est_curr_d4:,.0f}",
                                                                 f"{row.est_curr_d5:,.0f}",
                                                                 f"{row.qtd_est_total:,.0f}",
                                                                 f"{row.ytd_est_ac:,.0f}"
                                                                 )

    return readHTMLFile(1).format(strHTML)


def generateHTMLTotal():
    strSQL = """
    SELECT seqn_no, trns_name, ll_curr_q_tg, qtd_ac, est_curr_totl, qtd_est_totl, qtd_est_diff_ll_curr_q_tg
    FROM dbo.crm_mail_est_reve_totl
    ORDER BY seqn_no
    """

    myConnDB = ConnectDB()
    result_set = myConnDB.query(strSQL)

    strHTML = ""
    str_tmpHTML = """
    <tr>
        <td align="left">{0}</td>
        <td align="right">{1}&nbsp;</td>
        <td align="right">{2}&nbsp;</td>
        <td align="right">{3}&nbsp;</td>
        <td align="right">{4}&nbsp;</td>
        <td align="right">{5}&nbsp;</td>
    </tr>
    """

    str_total_tmpHTML = """
    <tr>
        <td align="right" bgcolor="#FFFFCC"><strong>{0}</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{1}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{2}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{3}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{4}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{5}&nbsp;</strong></td>
    </tr>
    """

    for row in result_set:
        if row.seqn_no != 7:
            strHTML = strHTML.strip() + str_tmpHTML.format(row.trns_name.strip(),
                                                           f"{row.ll_curr_q_tg:,.0f}",
                                                           f"{row.qtd_ac:,.0f}",
                                                           f"{row.est_curr_totl:,.0f}",
                                                           f"{row.qtd_est_totl:,.0f}",
                                                           f"{row.qtd_est_diff_ll_curr_q_tg:,.0f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.trns_name.strip(),
                                                           f"{row.ll_curr_q_tg:,.0f}",
                                                           f"{row.qtd_ac:,.0f}",
                                                           f"{row.est_curr_totl:,.0f}",
                                                           f"{row.qtd_est_totl:,.0f}",
                                                           f"{row.qtd_est_diff_ll_curr_q_tg:,.0f}"
                                                           )

    return readHTMLFile(2).format(strHTML)


def generateHTMLbyProj():
    strSQL = """
    SELECT seqn_no, project_name, curr_q_m1_ac_u, curr_q_m1_ac_vol, curr_q_m2_ac_u, curr_q_m2_ac_vol, curr_q_m3_ac_u,
    curr_q_m3_ac_vol, qtd_curr_ac_u, qtd_curr_q_ac_vol
    FROM dbo.crm_mail_est_reve_byproj
    ORDER BY seqn_no
    """

    myConnDB = ConnectDB()
    result_set = myConnDB.query(strSQL)

    strHTML = ""
    str_tmpHTML = """
    <tr>
        <td align="left">{0}</td>
        <td align="right" >{1}&nbsp;</td>
        <td align="right" >{2}&nbsp;</td>
        <td align="right">{3}&nbsp;</td>
        <td align="right">{4}&nbsp;</td>
        <td align="right">{5}&nbsp;</td>
        <td align="right">{6}&nbsp;</td>
        <td align="right">{7}&nbsp;</td>
        <td align="right">{8}&nbsp;</td>
    </tr>
    """

    str_total_tmpHTML = """
    <tr>
        <td align="right" bgcolor="#FFFFCC"><strong>{0}</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{1}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{2}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{3}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{4}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{5}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{6}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{7}&nbsp;</strong></td>
        <td align="right" bgcolor="#FFFFCC"><strong>{8}&nbsp;</strong></td>
    </tr>
    """

    for row in result_set:
        if row.seqn_no != 6:
            strHTML = strHTML.strip() + str_tmpHTML.format(row.project_name.strip(),
                                                           f"{row.curr_q_m1_ac_u:,.0f}",
                                                           f"{row.curr_q_m1_ac_vol:,.0f}",
                                                           f"{row.curr_q_m2_ac_u:,.0f}",
                                                           f"{row.curr_q_m2_ac_vol:,.0f}",
                                                           f"{row.curr_q_m3_ac_u:,.0f}",
                                                           f"{row.curr_q_m3_ac_vol:,.0f}",
                                                           f"{row.qtd_curr_ac_u:,.0f}",
                                                           f"{row.qtd_curr_q_ac_vol:,.0f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.project_name.strip(),
                                                           f"{row.curr_q_m1_ac_u:,.0f}",
                                                           f"{row.curr_q_m1_ac_vol:,.0f}",
                                                           f"{row.curr_q_m2_ac_u:,.0f}",
                                                           f"{row.curr_q_m2_ac_vol:,.0f}",
                                                           f"{row.curr_q_m3_ac_u:,.0f}",
                                                           f"{row.curr_q_m3_ac_vol:,.0f}",
                                                           f"{row.qtd_curr_ac_u:,.0f}",
                                                           f"{row.qtd_curr_q_ac_vol:,.0f}"
                                                           )

    return readHTMLFile(3).format(strHTML)


def readHTMLFile(p_parm: int = None):
    f = None
    if p_parm == 1:
        f = codecs.open("templates/template_est_rev.html", 'r')
    if p_parm == 2:
        f = codecs.open("templates/template_est_rev_total.html", 'r')
    if p_parm == 3:
        f = codecs.open("templates/template_est_rev_byproj.html", 'r')

    return f.read()


def main():
    # receivers = ['suchat_s@apthai.com', 'pimonwan@apthai.com', 'pornnapa@apthai.com']
    receivers = ['suchat_s@apthai.com']
    subject = EST_MAIL_SUBJECT
    # bodyMsg = generateHTMLbyProj()
    # bodyMsg = EST_MAIL_BODY
    bodyMsg = "{0} <br /> {1} <br /> {2}".format(generateHTMLDetl(), generateHTMLTotal(), generateHTMLbyProj())
    sender = MAIL_SENDER

    attachedFile = []

    # Send Email to Customer
    send_email(subject, bodyMsg, sender, receivers, attachedFile)


if __name__ == '__main__':
    main()
