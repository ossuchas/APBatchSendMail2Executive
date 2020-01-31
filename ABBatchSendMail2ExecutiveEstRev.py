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


def getHeaderTable():
    strSQL = """
    EXEC [dbo].[sp_proc_mail_est_reve_DH]
    """

    myConnDB = ConnectDB()
    result_set = myConnDB.exec_spRet(strSQL, params=())
    returnVal = []
    for row in result_set:
        returnVal.append(row.head_curr_quarter)
        returnVal.append(row.head_est_d1)
        returnVal.append(row.head_est_d2)
        returnVal.append(row.head_est_d3)
        returnVal.append(row.head_est_d4)
        returnVal.append(row.head_est_d5)
        returnVal.append(row.head_m1)
        returnVal.append(row.head_m2)
        returnVal.append(row.head_m3)

    return returnVal


def generateHTMLDetl():
    strSQL = """
    SELECT seqn_no, trns_name, ac_q1, qtd_curr_ac, est_curr_d1, est_curr_d2, est_curr_d3,
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
    </tr>
    """

    for row in result_set:
        if row.seqn_no != 3:
            strHTML = strHTML.strip() + str_tmpHTML.format(row.trns_name.strip(),
                                                           f"{row.ac_q1:,.02f}",
                                                           f"{row.qtd_curr_ac:,.02f}",
                                                           f"{row.est_curr_d1:,.02f}",
                                                           f"{row.est_curr_d2:,.02f}",
                                                           f"{row.est_curr_d3:,.02f}",
                                                           f"{row.est_curr_d4:,.02f}",
                                                           f"{row.est_curr_d5:,.02f}",
                                                           f"{row.qtd_est_total:,.02f}",
                                                           f"{row.ytd_est_ac:,.02f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.trns_name.strip(),
                                                                 f"{row.ac_q1:,.02f}",
                                                                 f"{row.qtd_curr_ac:,.02f}",
                                                                 f"{row.est_curr_d1:,.02f}",
                                                                 f"{row.est_curr_d2:,.02f}",
                                                                 f"{row.est_curr_d3:,.02f}",
                                                                 f"{row.est_curr_d4:,.02f}",
                                                                 f"{row.est_curr_d5:,.02f}",
                                                                 f"{row.qtd_est_total:,.02f}",
                                                                 f"{row.ytd_est_ac:,.02f}"
                                                                 )

    return readHTMLFile(1).format(getHeaderTable()[0],
                                  getHeaderTable()[1],
                                  getHeaderTable()[2],
                                  getHeaderTable()[3],
                                  getHeaderTable()[4],
                                  getHeaderTable()[5],
                                  getHeaderTable()[0],
                                  strHTML)


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
                                                           f"{row.ll_curr_q_tg:,.02f}",
                                                           f"{row.qtd_ac:,.02f}",
                                                           f"{row.est_curr_totl:,.02f}",
                                                           f"{row.qtd_est_totl:,.02f}",
                                                           f"{row.qtd_est_diff_ll_curr_q_tg:,.02f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.trns_name.strip(),
                                                           f"{row.ll_curr_q_tg:,.02f}",
                                                           f"{row.qtd_ac:,.02f}",
                                                           f"{row.est_curr_totl:,.02f}",
                                                           f"{row.qtd_est_totl:,.02f}",
                                                           f"{row.qtd_est_diff_ll_curr_q_tg:,.02f}"
                                                           )

    return readHTMLFile(2).format(strHTML, getHeaderTable()[0])


def generateHTMLbyProj():
    strSQL = """
    SELECT a.seqn_no, a.project_name, a.curr_q_m1_ac_u, a.curr_q_m1_ac_vol, a.curr_q_m2_ac_u,
       a.curr_q_m2_ac_vol, a.curr_q_m3_ac_u, a.curr_q_m3_ac_vol, a.qtd_curr_ac_u, a.qtd_curr_q_ac_vol
       FROM dbo.crm_mail_est_reve_byproj a WITH(NOLOCK), dbo.crm_mail_est_reve_conf t WITH(NOLOCK)
       WHERE a.projectid = t.projectid
       AND CAST(t.effc_date AS DATE) <= CAST(GETDATE() AS DATE)
       AND (t.expr_date IS NULL OR CAST(t.expr_date AS DATE) >= CAST(GETDATE() AS DATE))
       ORDER BY a.seqn_no
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
                                                           f"{row.curr_q_m1_ac_u:,.02f}",
                                                           f"{row.curr_q_m1_ac_vol:,.02f}",
                                                           f"{row.curr_q_m2_ac_u:,.02f}",
                                                           f"{row.curr_q_m2_ac_vol:,.02f}",
                                                           f"{row.curr_q_m3_ac_u:,.02f}",
                                                           f"{row.curr_q_m3_ac_vol:,.02f}",
                                                           f"{row.qtd_curr_ac_u:,.02f}",
                                                           f"{row.qtd_curr_q_ac_vol:,.02f}"
                                                           )
        else:
            strHTML = strHTML.strip() + str_total_tmpHTML.format(row.project_name.strip(),
                                                           f"{row.curr_q_m1_ac_u:,.02f}",
                                                           f"{row.curr_q_m1_ac_vol:,.02f}",
                                                           f"{row.curr_q_m2_ac_u:,.02f}",
                                                           f"{row.curr_q_m2_ac_vol:,.02f}",
                                                           f"{row.curr_q_m3_ac_u:,.02f}",
                                                           f"{row.curr_q_m3_ac_vol:,.02f}",
                                                           f"{row.qtd_curr_ac_u:,.02f}",
                                                           f"{row.qtd_curr_q_ac_vol:,.02f}"
                                                           )

    return readHTMLFile(3).format(strHTML,
                                  getHeaderTable()[0],
                                  getHeaderTable()[6],
                                  getHeaderTable()[7],
                                  getHeaderTable()[8]
                                  )


def readHTMLFile(p_parm: int = None):
    f = None
    if p_parm == 1:
        f = codecs.open("templates/template_est_rev.html", 'r')
    if p_parm == 2:
        f = codecs.open("templates/template_est_rev_total.html", 'r')
    if p_parm == 3:
        f = codecs.open("templates/template_est_rev_byproj.html", 'r')

    return f.read()


def refreshDataLastUpdate():
    strSQL = """
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '1', @p_option = ''
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '2', @p_option = ''
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '3', @p_option = 'AP'
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '3', @p_option = 'JV'
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '4', @p_option = 'AP'
    EXEC dbo.sp_proc_mail_est_reve_detl @p_type = '4', @p_option = 'JV'
    EXEC dbo.sp_proc_mail_est_reve_totl
    EXEC dbo.sp_proc_mail_est_reve_byproj
    """

    myConnDB = ConnectDB()
    myConnDB.exec_sp(strSQL, params=())


def main():
    # Refresh Data
    refreshDataLastUpdate()

    receivers = ['suchat_s@apthai.com', 'pimonwan@apthai.com', 'pornnapa@apthai.com',
                 'tanonchai@apthai.com', 'polwaritpakorn@apthai.com', 'jintana_i@apthai.com',
                 'apichaya@apthai.com']
    # receivers = ['suchat_s@apthai.com']
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
