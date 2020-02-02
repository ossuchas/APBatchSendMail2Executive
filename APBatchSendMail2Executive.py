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


def refreshDataLastUpdate():
    strSQL = """
    EXEC [dbo].[sp_proc_mail_llbw_all]
    """

    myConnDB = ConnectDB()
    myConnDB.exec_sp(strSQL, params=())


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


def getDatafromTable(sub_no: int = None):
    strSQL = """
    SELECT * FROM dbo.crm_mail_ll_data WHERE subject_no = {} ORDER BY tran_id
    """.format(sub_no)

    params = 'Driver={ODBC Driver 17 for SQL Server};Server=192.168.2.58;Database=db_iconcrm_fusion;uid=iconuser;pwd' \
             '=P@ssw0rd; '
    params = urllib.parse.quote_plus(params)
    db = create_engine('mssql+pyodbc:///?odbc_connect=%s' % params, fast_executemany=True)

    df = pd.read_sql(sql=strSQL, con=db)
    # print(df)
    # print(df['qtd_curr_w_ac'][1])

    return df


def main():
    # Refresh Data
    refreshDataLastUpdate()

    # receivers = ['suchat_s@apthai.com', 'apichaya@apthai.com', 'jintana_i@apthai.com', 'polwaritpakorn@apthai.com',
    #              'tanonchai@apthai.com', 'teerapat_s@apthai.com']

    # receivers = ['suchat_s@apthai.com', 'apichaya@apthai.com', 'jintana_i@apthai.com', 'polwaritpakorn@apthai.com',
    #              'tanonchai@apthai.com', 'teerapat_s@apthai.com', 'laddawan_v@apthai.com', 'woraphan_c@apthai.com',
    #              'raweewan_p@apthai.com', 'srisakul_p@apthai.com', 'jatuporn_p@apthai.com']
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

    df1 = getDatafromTable(sub_no=1)
    df2 = getDatafromTable(sub_no=2)
    df3 = getDatafromTable(sub_no=3)
    df4 = getDatafromTable(sub_no=4)
    df5 = getDatafromTable(sub_no=5)
    df6 = getDatafromTable(sub_no=6)
    df7 = getDatafromTable(sub_no=7)
    df8 = getDatafromTable(sub_no=8)
    df9 = getDatafromTable(sub_no=9)

    total_week = returnVal[0]
    if total_week >= 4:
        return readhtml(1) \
               .format(getHeaderTable(0)[0], getHeaderTable(0)[1], getHeaderTable(0)[2],
                       getHeaderTable(0)[3], getHeaderTable(0)[4], getHeaderTable(0)[5],
                       getHeaderTable(0)[6], getHeaderTable(0)[7], getHeaderTable(0)[8],
                       getHeaderTable(0)[9], getHeaderTable(0)[10], getHeaderTable(0)[11],
                       # SDH
                       f"{df1['curr_q_q1tg'][0]:,.0f}", f"{df1['curr_q_q2tg'][0]:,.0f}",
                       f"{df1['curr_q_q3tg'][0]:,.0f}", f"{df1['curr_q_total'][0]:,.0f}",
                       f"{df1['qtd_curr_w_tg'][0]:,.0f}", f"{df1['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][0]:,.0f}", f"{df1['qtg_val'][0]:,.0f}",
                       f"{df1['curr_m_w1_tg'][0]:,.0f}", f"{df1['curr_m_w1_ac'][0]:,.0f}",
                       f"{df1['curr_m_w2_tg'][0]:,.0f}", f"{df1['curr_m_w2_ac'][0]:,.0f}",
                       f"{df1['curr_m_w3_tg'][0]:,.0f}", f"{df1['curr_m_w3_ac'][0]:,.0f}",
                       f"{df1['curr_m_w4_tg'][0]:,.0f}", f"{df1['curr_m_w4_ac'][0]:,.0f}",
                       f"{df1['curr_m_w5_tg'][0]:,.0f}", f"{df1['curr_m_w5_ac'][0]:,.0f}",
                       f"{df1['full_month_tg'][0]:,.0f}", f"{df1['full_month_ac'][0]:,.0f}",
                       f"{df1['full_month_dff'][0]:,.0f}",
                       f"{df1['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df1['curr_m_w1_flag'][0]}", f"{df1['curr_m_w2_flag'][0]}",
                       f"{df1['curr_m_w3_flag'][0]}", f"{df1['curr_m_w4_flag'][0]}",
                       f"{df1['curr_m_w5_flag'][0]}", f"{df1['full_month_flag'][0]}",
                       # TH
                       f"{df1['curr_q_q1tg'][1]:,.0f}", f"{df1['curr_q_q2tg'][1]:,.0f}",
                       f"{df1['curr_q_q3tg'][1]:,.0f}", f"{df1['curr_q_total'][1]:,.0f}",
                       f"{df1['qtd_curr_w_tg'][1]:,.0f}", f"{df1['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][1]:,.0f}", f"{df1['qtg_val'][1]:,.0f}",
                       f"{df1['curr_m_w1_tg'][1]:,.0f}", f"{df1['curr_m_w1_ac'][1]:,.0f}",
                       f"{df1['curr_m_w2_tg'][1]:,.0f}", f"{df1['curr_m_w2_ac'][1]:,.0f}",
                       f"{df1['curr_m_w3_tg'][1]:,.0f}", f"{df1['curr_m_w3_ac'][1]:,.0f}",
                       f"{df1['curr_m_w4_tg'][1]:,.0f}", f"{df1['curr_m_w4_ac'][1]:,.0f}",
                       f"{df1['curr_m_w5_tg'][1]:,.0f}", f"{df1['curr_m_w5_ac'][1]:,.0f}",
                       f"{df1['full_month_tg'][1]:,.0f}", f"{df1['full_month_ac'][1]:,.0f}",
                       f"{df1['full_month_dff'][1]:,.0f}", f"{df1['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df1['curr_m_w1_flag'][1]}", f"{df1['curr_m_w2_flag'][1]}", f"{df1['curr_m_w3_flag'][1]}",
                       f"{df1['curr_m_w4_flag'][1]}", f"{df1['curr_m_w5_flag'][1]}", f"{df1['full_month_flag'][1]}",
                       # TH BKM
                       f"{df1['curr_q_q1tg'][2]:,.0f}", f"{df1['curr_q_q2tg'][2]:,.0f}", f"{df1['curr_q_q3tg'][2]:,.0f}",
                       f"{df1['curr_q_total'][2]:,.0f}", f"{df1['qtd_curr_w_tg'][2]:,.0f}", f"{df1['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][2]:,.0f}", f"{df1['qtg_val'][2]:,.0f}", f"{df1['curr_m_w1_tg'][2]:,.0f}",
                       f"{df1['curr_m_w1_ac'][2]:,.0f}", f"{df1['curr_m_w2_tg'][2]:,.0f}", f"{df1['curr_m_w2_ac'][2]:,.0f}",
                       f"{df1['curr_m_w3_tg'][2]:,.0f}", f"{df1['curr_m_w3_ac'][2]:,.0f}", f"{df1['curr_m_w4_tg'][2]:,.0f}",
                       f"{df1['curr_m_w4_ac'][2]:,.0f}", f"{df1['curr_m_w5_tg'][2]:,.0f}", f"{df1['curr_m_w5_ac'][2]:,.0f}",
                       f"{df1['full_month_tg'][2]:,.0f}", f"{df1['full_month_ac'][2]:,.0f}", f"{df1['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df1['qtd_curr_w_flag'][2]}", f"{df1['curr_m_w1_flag'][2]}", f"{df1['curr_m_w2_flag'][2]}",
                       f"{df1['curr_m_w3_flag'][2]}", f"{df1['curr_m_w4_flag'][2]}", f"{df1['curr_m_w5_flag'][2]}",
                       f"{df1['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df1['curr_q_q1tg'][3]:,.0f}", f"{df1['curr_q_q2tg'][3]:,.0f}", f"{df1['curr_q_q3tg'][3]:,.0f}",
                       f"{df1['curr_q_total'][3]:,.0f}", f"{df1['qtd_curr_w_tg'][3]:,.0f}", f"{df1['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][3]:,.0f}", f"{df1['qtg_val'][3]:,.0f}", f"{df1['curr_m_w1_tg'][3]:,.0f}",
                       f"{df1['curr_m_w1_ac'][3]:,.0f}", f"{df1['curr_m_w2_tg'][3]:,.0f}", f"{df1['curr_m_w2_ac'][3]:,.0f}",
                       f"{df1['curr_m_w3_tg'][3]:,.0f}", f"{df1['curr_m_w3_ac'][3]:,.0f}", f"{df1['curr_m_w4_tg'][3]:,.0f}",
                       f"{df1['curr_m_w4_ac'][3]:,.0f}", f"{df1['curr_m_w5_tg'][3]:,.0f}", f"{df1['curr_m_w5_ac'][3]:,.0f}",
                       f"{df1['full_month_tg'][3]:,.0f}", f"{df1['full_month_ac'][3]:,.0f}", f"{df1['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df1['qtd_curr_w_flag'][3]}", f"{df1['curr_m_w1_flag'][3]}", f"{df1['curr_m_w2_flag'][3]}",
                       f"{df1['curr_m_w3_flag'][3]}", f"{df1['curr_m_w4_flag'][3]}", f"{df1['curr_m_w5_flag'][3]}",
                       f"{df1['full_month_flag'][3]}",
                       # CD1
                       f"{df1['curr_q_q1tg'][4]:,.0f}", f"{df1['curr_q_q2tg'][4]:,.0f}", f"{df1['curr_q_q3tg'][4]:,.0f}",
                       f"{df1['curr_q_total'][4]:,.0f}", f"{df1['qtd_curr_w_tg'][4]:,.0f}", f"{df1['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][4]:,.0f}", f"{df1['qtg_val'][4]:,.0f}", f"{df1['curr_m_w1_tg'][4]:,.0f}",
                       f"{df1['curr_m_w1_ac'][4]:,.0f}", f"{df1['curr_m_w2_tg'][4]:,.0f}", f"{df1['curr_m_w2_ac'][4]:,.0f}",
                       f"{df1['curr_m_w3_tg'][4]:,.0f}", f"{df1['curr_m_w3_ac'][4]:,.0f}", f"{df1['curr_m_w4_tg'][4]:,.0f}",
                       f"{df1['curr_m_w4_ac'][4]:,.0f}", f"{df1['curr_m_w5_tg'][4]:,.0f}", f"{df1['curr_m_w5_ac'][4]:,.0f}",
                       f"{df1['full_month_tg'][4]:,.0f}", f"{df1['full_month_ac'][4]:,.0f}", f"{df1['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df1['qtd_curr_w_flag'][4]}", f"{df1['curr_m_w1_flag'][4]}", f"{df1['curr_m_w2_flag'][4]}",
                       f"{df1['curr_m_w3_flag'][4]}", f"{df1['curr_m_w4_flag'][4]}", f"{df1['curr_m_w5_flag'][4]}",
                       f"{df1['full_month_flag'][4]}",
                       # CD2
                       f"{df1['curr_q_q1tg'][5]:,.0f}", f"{df1['curr_q_q2tg'][5]:,.0f}", f"{df1['curr_q_q3tg'][5]:,.0f}",
                       f"{df1['curr_q_total'][5]:,.0f}", f"{df1['qtd_curr_w_tg'][5]:,.0f}", f"{df1['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][5]:,.0f}", f"{df1['qtg_val'][5]:,.0f}", f"{df1['curr_m_w1_tg'][5]:,.0f}",
                       f"{df1['curr_m_w1_ac'][5]:,.0f}", f"{df1['curr_m_w2_tg'][5]:,.0f}", f"{df1['curr_m_w2_ac'][5]:,.0f}",
                       f"{df1['curr_m_w3_tg'][5]:,.0f}", f"{df1['curr_m_w3_ac'][5]:,.0f}", f"{df1['curr_m_w4_tg'][5]:,.0f}",
                       f"{df1['curr_m_w4_ac'][5]:,.0f}", f"{df1['curr_m_w5_tg'][5]:,.0f}", f"{df1['curr_m_w5_ac'][5]:,.0f}",
                       f"{df1['full_month_tg'][5]:,.0f}", f"{df1['full_month_ac'][5]:,.0f}", f"{df1['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df1['qtd_curr_w_flag'][5]}", f"{df1['curr_m_w1_flag'][5]}", f"{df1['curr_m_w2_flag'][5]}",
                       f"{df1['curr_m_w3_flag'][5]}", f"{df1['curr_m_w4_flag'][5]}", f"{df1['curr_m_w5_flag'][5]}",
                       f"{df1['full_month_flag'][5]}",
                       # Total
                       f"{df1['curr_q_q1tg'][6]:,.0f}", f"{df1['curr_q_q2tg'][6]:,.0f}", f"{df1['curr_q_q3tg'][6]:,.0f}",
                       f"{df1['curr_q_total'][6]:,.0f}", f"{df1['qtd_curr_w_tg'][6]:,.0f}", f"{df1['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df1['qtd_curr_w_diff'][6]:,.0f}", f"{df1['qtg_val'][6]:,.0f}", f"{df1['curr_m_w1_tg'][6]:,.0f}",
                       f"{df1['curr_m_w1_ac'][6]:,.0f}", f"{df1['curr_m_w2_tg'][6]:,.0f}", f"{df1['curr_m_w2_ac'][6]:,.0f}",
                       f"{df1['curr_m_w3_tg'][6]:,.0f}", f"{df1['curr_m_w3_ac'][6]:,.0f}", f"{df1['curr_m_w4_tg'][6]:,.0f}",
                       f"{df1['curr_m_w4_ac'][6]:,.0f}", f"{df1['curr_m_w5_tg'][6]:,.0f}", f"{df1['curr_m_w5_ac'][6]:,.0f}",
                       f"{df1['full_month_tg'][6]:,.0f}", f"{df1['full_month_ac'][6]:,.0f}", f"{df1['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df1['qtd_curr_w_flag'][6]}", f"{df1['curr_m_w1_flag'][6]}", f"{df1['curr_m_w2_flag'][6]}",
                       f"{df1['curr_m_w3_flag'][6]}", f"{df1['curr_m_w4_flag'][6]}", f"{df1['curr_m_w5_flag'][6]}",
                       f"{df1['full_month_flag'][6]}" ) \
        + "<br />" + readhtml(1) \
               .format(getHeaderTable(1)[0], getHeaderTable(1)[1], getHeaderTable(1)[2],
                       getHeaderTable(1)[3], getHeaderTable(1)[4], getHeaderTable(1)[5],
                       getHeaderTable(1)[6], getHeaderTable(1)[7], getHeaderTable(1)[8],
                       getHeaderTable(1)[9], getHeaderTable(1)[10], getHeaderTable(1)[11],
                       # SDH
                       f"{df2['curr_q_q1tg'][0]:,.0f}", f"{df2['curr_q_q2tg'][0]:,.0f}",
                       f"{df2['curr_q_q3tg'][0]:,.0f}", f"{df2['curr_q_total'][0]:,.0f}",
                       f"{df2['qtd_curr_w_tg'][0]:,.0f}", f"{df2['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][0]:,.0f}", f"{df2['qtg_val'][0]:,.0f}",
                       f"{df2['curr_m_w1_tg'][0]:,.0f}", f"{df2['curr_m_w1_ac'][0]:,.0f}",
                       f"{df2['curr_m_w2_tg'][0]:,.0f}", f"{df2['curr_m_w2_ac'][0]:,.0f}",
                       f"{df2['curr_m_w3_tg'][0]:,.0f}", f"{df2['curr_m_w3_ac'][0]:,.0f}",
                       f"{df2['curr_m_w4_tg'][0]:,.0f}", f"{df2['curr_m_w4_ac'][0]:,.0f}",
                       f"{df2['curr_m_w5_tg'][0]:,.0f}", f"{df2['curr_m_w5_ac'][0]:,.0f}",
                       f"{df2['full_month_tg'][0]:,.0f}", f"{df2['full_month_ac'][0]:,.0f}",
                       f"{df2['full_month_dff'][0]:,.0f}",
                       f"{df2['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df2['curr_m_w1_flag'][0]}", f"{df2['curr_m_w2_flag'][0]}",
                       f"{df2['curr_m_w3_flag'][0]}", f"{df2['curr_m_w4_flag'][0]}",
                       f"{df2['curr_m_w5_flag'][0]}", f"{df2['full_month_flag'][0]}",
                       # TH
                       f"{df2['curr_q_q1tg'][1]:,.0f}", f"{df2['curr_q_q2tg'][1]:,.0f}",
                       f"{df2['curr_q_q3tg'][1]:,.0f}", f"{df2['curr_q_total'][1]:,.0f}",
                       f"{df2['qtd_curr_w_tg'][1]:,.0f}", f"{df2['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][1]:,.0f}", f"{df2['qtg_val'][1]:,.0f}",
                       f"{df2['curr_m_w1_tg'][1]:,.0f}", f"{df2['curr_m_w1_ac'][1]:,.0f}",
                       f"{df2['curr_m_w2_tg'][1]:,.0f}", f"{df2['curr_m_w2_ac'][1]:,.0f}",
                       f"{df2['curr_m_w3_tg'][1]:,.0f}", f"{df2['curr_m_w3_ac'][1]:,.0f}",
                       f"{df2['curr_m_w4_tg'][1]:,.0f}", f"{df2['curr_m_w4_ac'][1]:,.0f}",
                       f"{df2['curr_m_w5_tg'][1]:,.0f}", f"{df2['curr_m_w5_ac'][1]:,.0f}",
                       f"{df2['full_month_tg'][1]:,.0f}", f"{df2['full_month_ac'][1]:,.0f}",
                       f"{df2['full_month_dff'][1]:,.0f}", f"{df2['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df2['curr_m_w1_flag'][1]}", f"{df2['curr_m_w2_flag'][1]}", f"{df2['curr_m_w3_flag'][1]}",
                       f"{df2['curr_m_w4_flag'][1]}", f"{df2['curr_m_w5_flag'][1]}", f"{df2['full_month_flag'][1]}",
                       # TH BKM
                       f"{df2['curr_q_q1tg'][2]:,.0f}", f"{df2['curr_q_q2tg'][2]:,.0f}",
                       f"{df2['curr_q_q3tg'][2]:,.0f}",
                       f"{df2['curr_q_total'][2]:,.0f}", f"{df2['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df2['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][2]:,.0f}", f"{df2['qtg_val'][2]:,.0f}",
                       f"{df2['curr_m_w1_tg'][2]:,.0f}",
                       f"{df2['curr_m_w1_ac'][2]:,.0f}", f"{df2['curr_m_w2_tg'][2]:,.0f}",
                       f"{df2['curr_m_w2_ac'][2]:,.0f}",
                       f"{df2['curr_m_w3_tg'][2]:,.0f}", f"{df2['curr_m_w3_ac'][2]:,.0f}",
                       f"{df2['curr_m_w4_tg'][2]:,.0f}",
                       f"{df2['curr_m_w4_ac'][2]:,.0f}", f"{df2['curr_m_w5_tg'][2]:,.0f}",
                       f"{df2['curr_m_w5_ac'][2]:,.0f}",
                       f"{df2['full_month_tg'][2]:,.0f}", f"{df2['full_month_ac'][2]:,.0f}",
                       f"{df2['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df2['qtd_curr_w_flag'][2]}", f"{df2['curr_m_w1_flag'][2]}", f"{df2['curr_m_w2_flag'][2]}",
                       f"{df2['curr_m_w3_flag'][2]}", f"{df2['curr_m_w4_flag'][2]}", f"{df2['curr_m_w5_flag'][2]}",
                       f"{df2['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df2['curr_q_q1tg'][3]:,.0f}", f"{df2['curr_q_q2tg'][3]:,.0f}",
                       f"{df2['curr_q_q3tg'][3]:,.0f}",
                       f"{df2['curr_q_total'][3]:,.0f}", f"{df2['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df2['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][3]:,.0f}", f"{df2['qtg_val'][3]:,.0f}",
                       f"{df2['curr_m_w1_tg'][3]:,.0f}",
                       f"{df2['curr_m_w1_ac'][3]:,.0f}", f"{df2['curr_m_w2_tg'][3]:,.0f}",
                       f"{df2['curr_m_w2_ac'][3]:,.0f}",
                       f"{df2['curr_m_w3_tg'][3]:,.0f}", f"{df2['curr_m_w3_ac'][3]:,.0f}",
                       f"{df2['curr_m_w4_tg'][3]:,.0f}",
                       f"{df2['curr_m_w4_ac'][3]:,.0f}", f"{df2['curr_m_w5_tg'][3]:,.0f}",
                       f"{df2['curr_m_w5_ac'][3]:,.0f}",
                       f"{df2['full_month_tg'][3]:,.0f}", f"{df2['full_month_ac'][3]:,.0f}",
                       f"{df2['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df2['qtd_curr_w_flag'][3]}", f"{df2['curr_m_w1_flag'][3]}", f"{df2['curr_m_w2_flag'][3]}",
                       f"{df2['curr_m_w3_flag'][3]}", f"{df2['curr_m_w4_flag'][3]}", f"{df2['curr_m_w5_flag'][3]}",
                       f"{df2['full_month_flag'][3]}",
                       # CD1
                       f"{df2['curr_q_q1tg'][4]:,.0f}", f"{df2['curr_q_q2tg'][4]:,.0f}",
                       f"{df2['curr_q_q3tg'][4]:,.0f}",
                       f"{df2['curr_q_total'][4]:,.0f}", f"{df2['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df2['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][4]:,.0f}", f"{df2['qtg_val'][4]:,.0f}",
                       f"{df2['curr_m_w1_tg'][4]:,.0f}",
                       f"{df2['curr_m_w1_ac'][4]:,.0f}", f"{df2['curr_m_w2_tg'][4]:,.0f}",
                       f"{df2['curr_m_w2_ac'][4]:,.0f}",
                       f"{df2['curr_m_w3_tg'][4]:,.0f}", f"{df2['curr_m_w3_ac'][4]:,.0f}",
                       f"{df2['curr_m_w4_tg'][4]:,.0f}",
                       f"{df2['curr_m_w4_ac'][4]:,.0f}", f"{df2['curr_m_w5_tg'][4]:,.0f}",
                       f"{df2['curr_m_w5_ac'][4]:,.0f}",
                       f"{df2['full_month_tg'][4]:,.0f}", f"{df2['full_month_ac'][4]:,.0f}",
                       f"{df2['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df2['qtd_curr_w_flag'][4]}", f"{df2['curr_m_w1_flag'][4]}", f"{df2['curr_m_w2_flag'][4]}",
                       f"{df2['curr_m_w3_flag'][4]}", f"{df2['curr_m_w4_flag'][4]}", f"{df2['curr_m_w5_flag'][4]}",
                       f"{df2['full_month_flag'][4]}",
                       # CD2
                       f"{df2['curr_q_q1tg'][5]:,.0f}", f"{df2['curr_q_q2tg'][5]:,.0f}",
                       f"{df2['curr_q_q3tg'][5]:,.0f}",
                       f"{df2['curr_q_total'][5]:,.0f}", f"{df2['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df2['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][5]:,.0f}", f"{df2['qtg_val'][5]:,.0f}",
                       f"{df2['curr_m_w1_tg'][5]:,.0f}",
                       f"{df2['curr_m_w1_ac'][5]:,.0f}", f"{df2['curr_m_w2_tg'][5]:,.0f}",
                       f"{df2['curr_m_w2_ac'][5]:,.0f}",
                       f"{df2['curr_m_w3_tg'][5]:,.0f}", f"{df2['curr_m_w3_ac'][5]:,.0f}",
                       f"{df2['curr_m_w4_tg'][5]:,.0f}",
                       f"{df2['curr_m_w4_ac'][5]:,.0f}", f"{df2['curr_m_w5_tg'][5]:,.0f}",
                       f"{df2['curr_m_w5_ac'][5]:,.0f}",
                       f"{df2['full_month_tg'][5]:,.0f}", f"{df2['full_month_ac'][5]:,.0f}",
                       f"{df2['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df2['qtd_curr_w_flag'][5]}", f"{df2['curr_m_w1_flag'][5]}", f"{df2['curr_m_w2_flag'][5]}",
                       f"{df2['curr_m_w3_flag'][5]}", f"{df2['curr_m_w4_flag'][5]}", f"{df2['curr_m_w5_flag'][5]}",
                       f"{df2['full_month_flag'][5]}",
                       # Total
                       f"{df2['curr_q_q1tg'][6]:,.0f}", f"{df2['curr_q_q2tg'][6]:,.0f}",
                       f"{df2['curr_q_q3tg'][6]:,.0f}",
                       f"{df2['curr_q_total'][6]:,.0f}", f"{df2['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df2['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df2['qtd_curr_w_diff'][6]:,.0f}", f"{df2['qtg_val'][6]:,.0f}",
                       f"{df2['curr_m_w1_tg'][6]:,.0f}",
                       f"{df2['curr_m_w1_ac'][6]:,.0f}", f"{df2['curr_m_w2_tg'][6]:,.0f}",
                       f"{df2['curr_m_w2_ac'][6]:,.0f}",
                       f"{df2['curr_m_w3_tg'][6]:,.0f}", f"{df2['curr_m_w3_ac'][6]:,.0f}",
                       f"{df2['curr_m_w4_tg'][6]:,.0f}",
                       f"{df2['curr_m_w4_ac'][6]:,.0f}", f"{df2['curr_m_w5_tg'][6]:,.0f}",
                       f"{df2['curr_m_w5_ac'][6]:,.0f}",
                       f"{df2['full_month_tg'][6]:,.0f}", f"{df2['full_month_ac'][6]:,.0f}",
                       f"{df2['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df2['qtd_curr_w_flag'][6]}", f"{df2['curr_m_w1_flag'][6]}", f"{df2['curr_m_w2_flag'][6]}",
                       f"{df2['curr_m_w3_flag'][6]}", f"{df2['curr_m_w4_flag'][6]}", f"{df2['curr_m_w5_flag'][6]}",
                       f"{df2['full_month_flag'][6]}") \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(2)[0], getHeaderTable(2)[1], getHeaderTable(2)[2],
                       getHeaderTable(2)[3], getHeaderTable(2)[4], getHeaderTable(2)[5],
                       getHeaderTable(2)[6], getHeaderTable(2)[7], getHeaderTable(2)[8],
                       getHeaderTable(2)[9], getHeaderTable(2)[10], getHeaderTable(2)[11],
                       # SDH
                       f"{df3['curr_q_q1tg'][0]:,.0f}", f"{df3['curr_q_q2tg'][0]:,.0f}",
                       f"{df3['curr_q_q3tg'][0]:,.0f}", f"{df3['curr_q_total'][0]:,.0f}",
                       f"{df3['qtd_curr_w_tg'][0]:,.0f}", f"{df3['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][0]:,.0f}", f"{df3['qtg_val'][0]:,.0f}",
                       f"{df3['curr_m_w1_tg'][0]:,.0f}", f"{df3['curr_m_w1_ac'][0]:,.0f}",
                       f"{df3['curr_m_w2_tg'][0]:,.0f}", f"{df3['curr_m_w2_ac'][0]:,.0f}",
                       f"{df3['curr_m_w3_tg'][0]:,.0f}", f"{df3['curr_m_w3_ac'][0]:,.0f}",
                       f"{df3['curr_m_w4_tg'][0]:,.0f}", f"{df3['curr_m_w4_ac'][0]:,.0f}",
                       f"{df3['curr_m_w5_tg'][0]:,.0f}", f"{df3['curr_m_w5_ac'][0]:,.0f}",
                       f"{df3['full_month_tg'][0]:,.0f}", f"{df3['full_month_ac'][0]:,.0f}",
                       f"{df3['full_month_dff'][0]:,.0f}",
                       f"{df3['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df3['curr_m_w1_flag'][0]}", f"{df3['curr_m_w2_flag'][0]}",
                       f"{df3['curr_m_w3_flag'][0]}", f"{df3['curr_m_w4_flag'][0]}",
                       f"{df3['curr_m_w5_flag'][0]}", f"{df3['full_month_flag'][0]}",
                       # TH
                       f"{df3['curr_q_q1tg'][1]:,.0f}", f"{df3['curr_q_q2tg'][1]:,.0f}",
                       f"{df3['curr_q_q3tg'][1]:,.0f}", f"{df3['curr_q_total'][1]:,.0f}",
                       f"{df3['qtd_curr_w_tg'][1]:,.0f}", f"{df3['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][1]:,.0f}", f"{df3['qtg_val'][1]:,.0f}",
                       f"{df3['curr_m_w1_tg'][1]:,.0f}", f"{df3['curr_m_w1_ac'][1]:,.0f}",
                       f"{df3['curr_m_w2_tg'][1]:,.0f}", f"{df3['curr_m_w2_ac'][1]:,.0f}",
                       f"{df3['curr_m_w3_tg'][1]:,.0f}", f"{df3['curr_m_w3_ac'][1]:,.0f}",
                       f"{df3['curr_m_w4_tg'][1]:,.0f}", f"{df3['curr_m_w4_ac'][1]:,.0f}",
                       f"{df3['curr_m_w5_tg'][1]:,.0f}", f"{df3['curr_m_w5_ac'][1]:,.0f}",
                       f"{df3['full_month_tg'][1]:,.0f}", f"{df3['full_month_ac'][1]:,.0f}",
                       f"{df3['full_month_dff'][1]:,.0f}", f"{df3['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df3['curr_m_w1_flag'][1]}", f"{df3['curr_m_w2_flag'][1]}", f"{df3['curr_m_w3_flag'][1]}",
                       f"{df3['curr_m_w4_flag'][1]}", f"{df3['curr_m_w5_flag'][1]}", f"{df3['full_month_flag'][1]}",
                       # TH BKM
                       f"{df3['curr_q_q1tg'][2]:,.0f}", f"{df3['curr_q_q2tg'][2]:,.0f}",
                       f"{df3['curr_q_q3tg'][2]:,.0f}",
                       f"{df3['curr_q_total'][2]:,.0f}", f"{df3['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df3['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][2]:,.0f}", f"{df3['qtg_val'][2]:,.0f}",
                       f"{df3['curr_m_w1_tg'][2]:,.0f}",
                       f"{df3['curr_m_w1_ac'][2]:,.0f}", f"{df3['curr_m_w2_tg'][2]:,.0f}",
                       f"{df3['curr_m_w2_ac'][2]:,.0f}",
                       f"{df3['curr_m_w3_tg'][2]:,.0f}", f"{df3['curr_m_w3_ac'][2]:,.0f}",
                       f"{df3['curr_m_w4_tg'][2]:,.0f}",
                       f"{df3['curr_m_w4_ac'][2]:,.0f}", f"{df3['curr_m_w5_tg'][2]:,.0f}",
                       f"{df3['curr_m_w5_ac'][2]:,.0f}",
                       f"{df3['full_month_tg'][2]:,.0f}", f"{df3['full_month_ac'][2]:,.0f}",
                       f"{df3['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df3['qtd_curr_w_flag'][2]}", f"{df3['curr_m_w1_flag'][2]}", f"{df3['curr_m_w2_flag'][2]}",
                       f"{df3['curr_m_w3_flag'][2]}", f"{df3['curr_m_w4_flag'][2]}", f"{df3['curr_m_w5_flag'][2]}",
                       f"{df3['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df3['curr_q_q1tg'][3]:,.0f}", f"{df3['curr_q_q2tg'][3]:,.0f}",
                       f"{df3['curr_q_q3tg'][3]:,.0f}",
                       f"{df3['curr_q_total'][3]:,.0f}", f"{df3['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df3['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][3]:,.0f}", f"{df3['qtg_val'][3]:,.0f}",
                       f"{df3['curr_m_w1_tg'][3]:,.0f}",
                       f"{df3['curr_m_w1_ac'][3]:,.0f}", f"{df3['curr_m_w2_tg'][3]:,.0f}",
                       f"{df3['curr_m_w2_ac'][3]:,.0f}",
                       f"{df3['curr_m_w3_tg'][3]:,.0f}", f"{df3['curr_m_w3_ac'][3]:,.0f}",
                       f"{df3['curr_m_w4_tg'][3]:,.0f}",
                       f"{df3['curr_m_w4_ac'][3]:,.0f}", f"{df3['curr_m_w5_tg'][3]:,.0f}",
                       f"{df3['curr_m_w5_ac'][3]:,.0f}",
                       f"{df3['full_month_tg'][3]:,.0f}", f"{df3['full_month_ac'][3]:,.0f}",
                       f"{df3['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df3['qtd_curr_w_flag'][3]}", f"{df3['curr_m_w1_flag'][3]}", f"{df3['curr_m_w2_flag'][3]}",
                       f"{df3['curr_m_w3_flag'][3]}", f"{df3['curr_m_w4_flag'][3]}", f"{df3['curr_m_w5_flag'][3]}",
                       f"{df3['full_month_flag'][3]}",
                       # CD1
                       f"{df3['curr_q_q1tg'][4]:,.0f}", f"{df3['curr_q_q2tg'][4]:,.0f}",
                       f"{df3['curr_q_q3tg'][4]:,.0f}",
                       f"{df3['curr_q_total'][4]:,.0f}", f"{df3['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df3['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][4]:,.0f}", f"{df3['qtg_val'][4]:,.0f}",
                       f"{df3['curr_m_w1_tg'][4]:,.0f}",
                       f"{df3['curr_m_w1_ac'][4]:,.0f}", f"{df3['curr_m_w2_tg'][4]:,.0f}",
                       f"{df3['curr_m_w2_ac'][4]:,.0f}",
                       f"{df3['curr_m_w3_tg'][4]:,.0f}", f"{df3['curr_m_w3_ac'][4]:,.0f}",
                       f"{df3['curr_m_w4_tg'][4]:,.0f}",
                       f"{df3['curr_m_w4_ac'][4]:,.0f}", f"{df3['curr_m_w5_tg'][4]:,.0f}",
                       f"{df3['curr_m_w5_ac'][4]:,.0f}",
                       f"{df3['full_month_tg'][4]:,.0f}", f"{df3['full_month_ac'][4]:,.0f}",
                       f"{df3['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df3['qtd_curr_w_flag'][4]}", f"{df3['curr_m_w1_flag'][4]}", f"{df3['curr_m_w2_flag'][4]}",
                       f"{df3['curr_m_w3_flag'][4]}", f"{df3['curr_m_w4_flag'][4]}", f"{df3['curr_m_w5_flag'][4]}",
                       f"{df3['full_month_flag'][4]}",
                       # CD2
                       f"{df3['curr_q_q1tg'][5]:,.0f}", f"{df3['curr_q_q2tg'][5]:,.0f}",
                       f"{df3['curr_q_q3tg'][5]:,.0f}",
                       f"{df3['curr_q_total'][5]:,.0f}", f"{df3['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df3['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][5]:,.0f}", f"{df3['qtg_val'][5]:,.0f}",
                       f"{df3['curr_m_w1_tg'][5]:,.0f}",
                       f"{df3['curr_m_w1_ac'][5]:,.0f}", f"{df3['curr_m_w2_tg'][5]:,.0f}",
                       f"{df3['curr_m_w2_ac'][5]:,.0f}",
                       f"{df3['curr_m_w3_tg'][5]:,.0f}", f"{df3['curr_m_w3_ac'][5]:,.0f}",
                       f"{df3['curr_m_w4_tg'][5]:,.0f}",
                       f"{df3['curr_m_w4_ac'][5]:,.0f}", f"{df3['curr_m_w5_tg'][5]:,.0f}",
                       f"{df3['curr_m_w5_ac'][5]:,.0f}",
                       f"{df3['full_month_tg'][5]:,.0f}", f"{df3['full_month_ac'][5]:,.0f}",
                       f"{df3['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df3['qtd_curr_w_flag'][5]}", f"{df3['curr_m_w1_flag'][5]}", f"{df3['curr_m_w2_flag'][5]}",
                       f"{df3['curr_m_w3_flag'][5]}", f"{df3['curr_m_w4_flag'][5]}", f"{df3['curr_m_w5_flag'][5]}",
                       f"{df3['full_month_flag'][5]}",
                       # Total
                       f"{df3['curr_q_q1tg'][6]:,.0f}", f"{df3['curr_q_q2tg'][6]:,.0f}",
                       f"{df3['curr_q_q3tg'][6]:,.0f}",
                       f"{df3['curr_q_total'][6]:,.0f}", f"{df3['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df3['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df3['qtd_curr_w_diff'][6]:,.0f}", f"{df3['qtg_val'][6]:,.0f}",
                       f"{df3['curr_m_w1_tg'][6]:,.0f}",
                       f"{df3['curr_m_w1_ac'][6]:,.0f}", f"{df3['curr_m_w2_tg'][6]:,.0f}",
                       f"{df3['curr_m_w2_ac'][6]:,.0f}",
                       f"{df3['curr_m_w3_tg'][6]:,.0f}", f"{df3['curr_m_w3_ac'][6]:,.0f}",
                       f"{df3['curr_m_w4_tg'][6]:,.0f}",
                       f"{df3['curr_m_w4_ac'][6]:,.0f}", f"{df3['curr_m_w5_tg'][6]:,.0f}",
                       f"{df3['curr_m_w5_ac'][6]:,.0f}",
                       f"{df3['full_month_tg'][6]:,.0f}", f"{df3['full_month_ac'][6]:,.0f}",
                       f"{df3['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df3['qtd_curr_w_flag'][6]}", f"{df3['curr_m_w1_flag'][6]}", f"{df3['curr_m_w2_flag'][6]}",
                       f"{df3['curr_m_w3_flag'][6]}", f"{df3['curr_m_w4_flag'][6]}", f"{df3['curr_m_w5_flag'][6]}",
                       f"{df3['full_month_flag'][6]}") \
               + "<br />" + readhtml(1) \
               .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                       getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                       getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                       getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11],
                       # SDH
                       f"{df4['curr_q_q1tg'][0]:,.0f}", f"{df4['curr_q_q2tg'][0]:,.0f}",
                       f"{df4['curr_q_q3tg'][0]:,.0f}", f"{df4['curr_q_total'][0]:,.0f}",
                       f"{df4['qtd_curr_w_tg'][0]:,.0f}", f"{df4['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][0]:,.0f}", f"{df4['qtg_val'][0]:,.0f}",
                       f"{df4['curr_m_w1_tg'][0]:,.0f}", f"{df4['curr_m_w1_ac'][0]:,.0f}",
                       f"{df4['curr_m_w2_tg'][0]:,.0f}", f"{df4['curr_m_w2_ac'][0]:,.0f}",
                       f"{df4['curr_m_w3_tg'][0]:,.0f}", f"{df4['curr_m_w3_ac'][0]:,.0f}",
                       f"{df4['curr_m_w4_tg'][0]:,.0f}", f"{df4['curr_m_w4_ac'][0]:,.0f}",
                       f"{df4['curr_m_w5_tg'][0]:,.0f}", f"{df4['curr_m_w5_ac'][0]:,.0f}",
                       f"{df4['full_month_tg'][0]:,.0f}", f"{df4['full_month_ac'][0]:,.0f}",
                       f"{df4['full_month_dff'][0]:,.0f}",
                       f"{df4['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df4['curr_m_w1_flag'][0]}", f"{df4['curr_m_w2_flag'][0]}",
                       f"{df4['curr_m_w3_flag'][0]}", f"{df4['curr_m_w4_flag'][0]}",
                       f"{df4['curr_m_w5_flag'][0]}", f"{df4['full_month_flag'][0]}",
                       # TH
                       f"{df4['curr_q_q1tg'][1]:,.0f}", f"{df4['curr_q_q2tg'][1]:,.0f}",
                       f"{df4['curr_q_q3tg'][1]:,.0f}", f"{df4['curr_q_total'][1]:,.0f}",
                       f"{df4['qtd_curr_w_tg'][1]:,.0f}", f"{df4['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][1]:,.0f}", f"{df4['qtg_val'][1]:,.0f}",
                       f"{df4['curr_m_w1_tg'][1]:,.0f}", f"{df4['curr_m_w1_ac'][1]:,.0f}",
                       f"{df4['curr_m_w2_tg'][1]:,.0f}", f"{df4['curr_m_w2_ac'][1]:,.0f}",
                       f"{df4['curr_m_w3_tg'][1]:,.0f}", f"{df4['curr_m_w3_ac'][1]:,.0f}",
                       f"{df4['curr_m_w4_tg'][1]:,.0f}", f"{df4['curr_m_w4_ac'][1]:,.0f}",
                       f"{df4['curr_m_w5_tg'][1]:,.0f}", f"{df4['curr_m_w5_ac'][1]:,.0f}",
                       f"{df4['full_month_tg'][1]:,.0f}", f"{df4['full_month_ac'][1]:,.0f}",
                       f"{df4['full_month_dff'][1]:,.0f}", f"{df4['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df4['curr_m_w1_flag'][1]}", f"{df4['curr_m_w2_flag'][1]}", f"{df4['curr_m_w3_flag'][1]}",
                       f"{df4['curr_m_w4_flag'][1]}", f"{df4['curr_m_w5_flag'][1]}", f"{df4['full_month_flag'][1]}",
                       # TH BKM
                       f"{df4['curr_q_q1tg'][2]:,.0f}", f"{df4['curr_q_q2tg'][2]:,.0f}",
                       f"{df4['curr_q_q3tg'][2]:,.0f}",
                       f"{df4['curr_q_total'][2]:,.0f}", f"{df4['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df4['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][2]:,.0f}", f"{df4['qtg_val'][2]:,.0f}",
                       f"{df4['curr_m_w1_tg'][2]:,.0f}",
                       f"{df4['curr_m_w1_ac'][2]:,.0f}", f"{df4['curr_m_w2_tg'][2]:,.0f}",
                       f"{df4['curr_m_w2_ac'][2]:,.0f}",
                       f"{df4['curr_m_w3_tg'][2]:,.0f}", f"{df4['curr_m_w3_ac'][2]:,.0f}",
                       f"{df4['curr_m_w4_tg'][2]:,.0f}",
                       f"{df4['curr_m_w4_ac'][2]:,.0f}", f"{df4['curr_m_w5_tg'][2]:,.0f}",
                       f"{df4['curr_m_w5_ac'][2]:,.0f}",
                       f"{df4['full_month_tg'][2]:,.0f}", f"{df4['full_month_ac'][2]:,.0f}",
                       f"{df4['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df4['qtd_curr_w_flag'][2]}", f"{df4['curr_m_w1_flag'][2]}", f"{df4['curr_m_w2_flag'][2]}",
                       f"{df4['curr_m_w3_flag'][2]}", f"{df4['curr_m_w4_flag'][2]}", f"{df4['curr_m_w5_flag'][2]}",
                       f"{df4['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df4['curr_q_q1tg'][3]:,.0f}", f"{df4['curr_q_q2tg'][3]:,.0f}",
                       f"{df4['curr_q_q3tg'][3]:,.0f}",
                       f"{df4['curr_q_total'][3]:,.0f}", f"{df4['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df4['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][3]:,.0f}", f"{df4['qtg_val'][3]:,.0f}",
                       f"{df4['curr_m_w1_tg'][3]:,.0f}",
                       f"{df4['curr_m_w1_ac'][3]:,.0f}", f"{df4['curr_m_w2_tg'][3]:,.0f}",
                       f"{df4['curr_m_w2_ac'][3]:,.0f}",
                       f"{df4['curr_m_w3_tg'][3]:,.0f}", f"{df4['curr_m_w3_ac'][3]:,.0f}",
                       f"{df4['curr_m_w4_tg'][3]:,.0f}",
                       f"{df4['curr_m_w4_ac'][3]:,.0f}", f"{df4['curr_m_w5_tg'][3]:,.0f}",
                       f"{df4['curr_m_w5_ac'][3]:,.0f}",
                       f"{df4['full_month_tg'][3]:,.0f}", f"{df4['full_month_ac'][3]:,.0f}",
                       f"{df4['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df4['qtd_curr_w_flag'][3]}", f"{df4['curr_m_w1_flag'][3]}", f"{df4['curr_m_w2_flag'][3]}",
                       f"{df4['curr_m_w3_flag'][3]}", f"{df4['curr_m_w4_flag'][3]}", f"{df4['curr_m_w5_flag'][3]}",
                       f"{df4['full_month_flag'][3]}",
                       # CD1
                       f"{df4['curr_q_q1tg'][4]:,.0f}", f"{df4['curr_q_q2tg'][4]:,.0f}",
                       f"{df4['curr_q_q3tg'][4]:,.0f}",
                       f"{df4['curr_q_total'][4]:,.0f}", f"{df4['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df4['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][4]:,.0f}", f"{df4['qtg_val'][4]:,.0f}",
                       f"{df4['curr_m_w1_tg'][4]:,.0f}",
                       f"{df4['curr_m_w1_ac'][4]:,.0f}", f"{df4['curr_m_w2_tg'][4]:,.0f}",
                       f"{df4['curr_m_w2_ac'][4]:,.0f}",
                       f"{df4['curr_m_w3_tg'][4]:,.0f}", f"{df4['curr_m_w3_ac'][4]:,.0f}",
                       f"{df4['curr_m_w4_tg'][4]:,.0f}",
                       f"{df4['curr_m_w4_ac'][4]:,.0f}", f"{df4['curr_m_w5_tg'][4]:,.0f}",
                       f"{df4['curr_m_w5_ac'][4]:,.0f}",
                       f"{df4['full_month_tg'][4]:,.0f}", f"{df4['full_month_ac'][4]:,.0f}",
                       f"{df4['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df4['qtd_curr_w_flag'][4]}", f"{df4['curr_m_w1_flag'][4]}", f"{df4['curr_m_w2_flag'][4]}",
                       f"{df4['curr_m_w3_flag'][4]}", f"{df4['curr_m_w4_flag'][4]}", f"{df4['curr_m_w5_flag'][4]}",
                       f"{df4['full_month_flag'][4]}",
                       # CD2
                       f"{df4['curr_q_q1tg'][5]:,.0f}", f"{df4['curr_q_q2tg'][5]:,.0f}",
                       f"{df4['curr_q_q3tg'][5]:,.0f}",
                       f"{df4['curr_q_total'][5]:,.0f}", f"{df4['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df4['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][5]:,.0f}", f"{df4['qtg_val'][5]:,.0f}",
                       f"{df4['curr_m_w1_tg'][5]:,.0f}",
                       f"{df4['curr_m_w1_ac'][5]:,.0f}", f"{df4['curr_m_w2_tg'][5]:,.0f}",
                       f"{df4['curr_m_w2_ac'][5]:,.0f}",
                       f"{df4['curr_m_w3_tg'][5]:,.0f}", f"{df4['curr_m_w3_ac'][5]:,.0f}",
                       f"{df4['curr_m_w4_tg'][5]:,.0f}",
                       f"{df4['curr_m_w4_ac'][5]:,.0f}", f"{df4['curr_m_w5_tg'][5]:,.0f}",
                       f"{df4['curr_m_w5_ac'][5]:,.0f}",
                       f"{df4['full_month_tg'][5]:,.0f}", f"{df4['full_month_ac'][5]:,.0f}",
                       f"{df4['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df4['qtd_curr_w_flag'][5]}", f"{df4['curr_m_w1_flag'][5]}", f"{df4['curr_m_w2_flag'][5]}",
                       f"{df4['curr_m_w3_flag'][5]}", f"{df4['curr_m_w4_flag'][5]}", f"{df4['curr_m_w5_flag'][5]}",
                       f"{df4['full_month_flag'][5]}",
                       # Total
                       f"{df4['curr_q_q1tg'][6]:,.0f}", f"{df4['curr_q_q2tg'][6]:,.0f}",
                       f"{df4['curr_q_q3tg'][6]:,.0f}",
                       f"{df4['curr_q_total'][6]:,.0f}", f"{df4['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df4['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df4['qtd_curr_w_diff'][6]:,.0f}", f"{df4['qtg_val'][6]:,.0f}",
                       f"{df4['curr_m_w1_tg'][6]:,.0f}",
                       f"{df4['curr_m_w1_ac'][6]:,.0f}", f"{df4['curr_m_w2_tg'][6]:,.0f}",
                       f"{df4['curr_m_w2_ac'][6]:,.0f}",
                       f"{df4['curr_m_w3_tg'][6]:,.0f}", f"{df4['curr_m_w3_ac'][6]:,.0f}",
                       f"{df4['curr_m_w4_tg'][6]:,.0f}",
                       f"{df4['curr_m_w4_ac'][6]:,.0f}", f"{df4['curr_m_w5_tg'][6]:,.0f}",
                       f"{df4['curr_m_w5_ac'][6]:,.0f}",
                       f"{df4['full_month_tg'][6]:,.0f}", f"{df4['full_month_ac'][6]:,.0f}",
                       f"{df4['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df4['qtd_curr_w_flag'][6]}", f"{df4['curr_m_w1_flag'][6]}", f"{df4['curr_m_w2_flag'][6]}",
                       f"{df4['curr_m_w3_flag'][6]}", f"{df4['curr_m_w4_flag'][6]}", f"{df4['curr_m_w5_flag'][6]}",
                       f"{df4['full_month_flag'][6]}") \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(4)[0], getHeaderTable(4)[1], getHeaderTable(4)[2],
                       getHeaderTable(4)[3], getHeaderTable(4)[4], getHeaderTable(4)[5],
                       getHeaderTable(4)[6], getHeaderTable(4)[7], getHeaderTable(4)[8],
                       getHeaderTable(4)[9], getHeaderTable(4)[10], getHeaderTable(4)[11],
                       # SDH
                       f"{df5['curr_q_q1tg'][0]:,.0f}", f"{df5['curr_q_q2tg'][0]:,.0f}",
                       f"{df5['curr_q_q3tg'][0]:,.0f}", f"{df5['curr_q_total'][0]:,.0f}",
                       f"{df5['qtd_curr_w_tg'][0]:,.0f}", f"{df5['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][0]:,.0f}", f"{df5['qtg_val'][0]:,.0f}",
                       f"{df5['curr_m_w1_tg'][0]:,.0f}", f"{df5['curr_m_w1_ac'][0]:,.0f}",
                       f"{df5['curr_m_w2_tg'][0]:,.0f}", f"{df5['curr_m_w2_ac'][0]:,.0f}",
                       f"{df5['curr_m_w3_tg'][0]:,.0f}", f"{df5['curr_m_w3_ac'][0]:,.0f}",
                       f"{df5['curr_m_w4_tg'][0]:,.0f}", f"{df5['curr_m_w4_ac'][0]:,.0f}",
                       f"{df5['curr_m_w5_tg'][0]:,.0f}", f"{df5['curr_m_w5_ac'][0]:,.0f}",
                       f"{df5['full_month_tg'][0]:,.0f}", f"{df5['full_month_ac'][0]:,.0f}",
                       f"{df5['full_month_dff'][0]:,.0f}",
                       f"{df5['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df5['curr_m_w1_flag'][0]}", f"{df5['curr_m_w2_flag'][0]}",
                       f"{df5['curr_m_w3_flag'][0]}", f"{df5['curr_m_w4_flag'][0]}",
                       f"{df5['curr_m_w5_flag'][0]}", f"{df5['full_month_flag'][0]}",
                       # TH
                       f"{df5['curr_q_q1tg'][1]:,.0f}", f"{df5['curr_q_q2tg'][1]:,.0f}",
                       f"{df5['curr_q_q3tg'][1]:,.0f}", f"{df5['curr_q_total'][1]:,.0f}",
                       f"{df5['qtd_curr_w_tg'][1]:,.0f}", f"{df5['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][1]:,.0f}", f"{df5['qtg_val'][1]:,.0f}",
                       f"{df5['curr_m_w1_tg'][1]:,.0f}", f"{df5['curr_m_w1_ac'][1]:,.0f}",
                       f"{df5['curr_m_w2_tg'][1]:,.0f}", f"{df5['curr_m_w2_ac'][1]:,.0f}",
                       f"{df5['curr_m_w3_tg'][1]:,.0f}", f"{df5['curr_m_w3_ac'][1]:,.0f}",
                       f"{df5['curr_m_w4_tg'][1]:,.0f}", f"{df5['curr_m_w4_ac'][1]:,.0f}",
                       f"{df5['curr_m_w5_tg'][1]:,.0f}", f"{df5['curr_m_w5_ac'][1]:,.0f}",
                       f"{df5['full_month_tg'][1]:,.0f}", f"{df5['full_month_ac'][1]:,.0f}",
                       f"{df5['full_month_dff'][1]:,.0f}", f"{df5['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df5['curr_m_w1_flag'][1]}", f"{df5['curr_m_w2_flag'][1]}", f"{df5['curr_m_w3_flag'][1]}",
                       f"{df5['curr_m_w4_flag'][1]}", f"{df5['curr_m_w5_flag'][1]}", f"{df5['full_month_flag'][1]}",
                       # TH BKM
                       f"{df5['curr_q_q1tg'][2]:,.0f}", f"{df5['curr_q_q2tg'][2]:,.0f}",
                       f"{df5['curr_q_q3tg'][2]:,.0f}",
                       f"{df5['curr_q_total'][2]:,.0f}", f"{df5['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df5['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][2]:,.0f}", f"{df5['qtg_val'][2]:,.0f}",
                       f"{df5['curr_m_w1_tg'][2]:,.0f}",
                       f"{df5['curr_m_w1_ac'][2]:,.0f}", f"{df5['curr_m_w2_tg'][2]:,.0f}",
                       f"{df5['curr_m_w2_ac'][2]:,.0f}",
                       f"{df5['curr_m_w3_tg'][2]:,.0f}", f"{df5['curr_m_w3_ac'][2]:,.0f}",
                       f"{df5['curr_m_w4_tg'][2]:,.0f}",
                       f"{df5['curr_m_w4_ac'][2]:,.0f}", f"{df5['curr_m_w5_tg'][2]:,.0f}",
                       f"{df5['curr_m_w5_ac'][2]:,.0f}",
                       f"{df5['full_month_tg'][2]:,.0f}", f"{df5['full_month_ac'][2]:,.0f}",
                       f"{df5['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df5['qtd_curr_w_flag'][2]}", f"{df5['curr_m_w1_flag'][2]}", f"{df5['curr_m_w2_flag'][2]}",
                       f"{df5['curr_m_w3_flag'][2]}", f"{df5['curr_m_w4_flag'][2]}", f"{df5['curr_m_w5_flag'][2]}",
                       f"{df5['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df5['curr_q_q1tg'][3]:,.0f}", f"{df5['curr_q_q2tg'][3]:,.0f}",
                       f"{df5['curr_q_q3tg'][3]:,.0f}",
                       f"{df5['curr_q_total'][3]:,.0f}", f"{df5['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df5['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][3]:,.0f}", f"{df5['qtg_val'][3]:,.0f}",
                       f"{df5['curr_m_w1_tg'][3]:,.0f}",
                       f"{df5['curr_m_w1_ac'][3]:,.0f}", f"{df5['curr_m_w2_tg'][3]:,.0f}",
                       f"{df5['curr_m_w2_ac'][3]:,.0f}",
                       f"{df5['curr_m_w3_tg'][3]:,.0f}", f"{df5['curr_m_w3_ac'][3]:,.0f}",
                       f"{df5['curr_m_w4_tg'][3]:,.0f}",
                       f"{df5['curr_m_w4_ac'][3]:,.0f}", f"{df5['curr_m_w5_tg'][3]:,.0f}",
                       f"{df5['curr_m_w5_ac'][3]:,.0f}",
                       f"{df5['full_month_tg'][3]:,.0f}", f"{df5['full_month_ac'][3]:,.0f}",
                       f"{df5['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df5['qtd_curr_w_flag'][3]}", f"{df5['curr_m_w1_flag'][3]}", f"{df5['curr_m_w2_flag'][3]}",
                       f"{df5['curr_m_w3_flag'][3]}", f"{df5['curr_m_w4_flag'][3]}", f"{df5['curr_m_w5_flag'][3]}",
                       f"{df5['full_month_flag'][3]}",
                       # CD1
                       f"{df5['curr_q_q1tg'][4]:,.0f}", f"{df5['curr_q_q2tg'][4]:,.0f}",
                       f"{df5['curr_q_q3tg'][4]:,.0f}",
                       f"{df5['curr_q_total'][4]:,.0f}", f"{df5['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df5['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][4]:,.0f}", f"{df5['qtg_val'][4]:,.0f}",
                       f"{df5['curr_m_w1_tg'][4]:,.0f}",
                       f"{df5['curr_m_w1_ac'][4]:,.0f}", f"{df5['curr_m_w2_tg'][4]:,.0f}",
                       f"{df5['curr_m_w2_ac'][4]:,.0f}",
                       f"{df5['curr_m_w3_tg'][4]:,.0f}", f"{df5['curr_m_w3_ac'][4]:,.0f}",
                       f"{df5['curr_m_w4_tg'][4]:,.0f}",
                       f"{df5['curr_m_w4_ac'][4]:,.0f}", f"{df5['curr_m_w5_tg'][4]:,.0f}",
                       f"{df5['curr_m_w5_ac'][4]:,.0f}",
                       f"{df5['full_month_tg'][4]:,.0f}", f"{df5['full_month_ac'][4]:,.0f}",
                       f"{df5['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df5['qtd_curr_w_flag'][4]}", f"{df5['curr_m_w1_flag'][4]}", f"{df5['curr_m_w2_flag'][4]}",
                       f"{df5['curr_m_w3_flag'][4]}", f"{df5['curr_m_w4_flag'][4]}", f"{df5['curr_m_w5_flag'][4]}",
                       f"{df5['full_month_flag'][4]}",
                       # CD2
                       f"{df5['curr_q_q1tg'][5]:,.0f}", f"{df5['curr_q_q2tg'][5]:,.0f}",
                       f"{df5['curr_q_q3tg'][5]:,.0f}",
                       f"{df5['curr_q_total'][5]:,.0f}", f"{df5['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df5['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][5]:,.0f}", f"{df5['qtg_val'][5]:,.0f}",
                       f"{df5['curr_m_w1_tg'][5]:,.0f}",
                       f"{df5['curr_m_w1_ac'][5]:,.0f}", f"{df5['curr_m_w2_tg'][5]:,.0f}",
                       f"{df5['curr_m_w2_ac'][5]:,.0f}",
                       f"{df5['curr_m_w3_tg'][5]:,.0f}", f"{df5['curr_m_w3_ac'][5]:,.0f}",
                       f"{df5['curr_m_w4_tg'][5]:,.0f}",
                       f"{df5['curr_m_w4_ac'][5]:,.0f}", f"{df5['curr_m_w5_tg'][5]:,.0f}",
                       f"{df5['curr_m_w5_ac'][5]:,.0f}",
                       f"{df5['full_month_tg'][5]:,.0f}", f"{df5['full_month_ac'][5]:,.0f}",
                       f"{df5['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df5['qtd_curr_w_flag'][5]}", f"{df5['curr_m_w1_flag'][5]}", f"{df5['curr_m_w2_flag'][5]}",
                       f"{df5['curr_m_w3_flag'][5]}", f"{df5['curr_m_w4_flag'][5]}", f"{df5['curr_m_w5_flag'][5]}",
                       f"{df5['full_month_flag'][5]}",
                       # Total
                       f"{df5['curr_q_q1tg'][6]:,.0f}", f"{df5['curr_q_q2tg'][6]:,.0f}",
                       f"{df5['curr_q_q3tg'][6]:,.0f}",
                       f"{df5['curr_q_total'][6]:,.0f}", f"{df5['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df5['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df5['qtd_curr_w_diff'][6]:,.0f}", f"{df5['qtg_val'][6]:,.0f}",
                       f"{df5['curr_m_w1_tg'][6]:,.0f}",
                       f"{df5['curr_m_w1_ac'][6]:,.0f}", f"{df5['curr_m_w2_tg'][6]:,.0f}",
                       f"{df5['curr_m_w2_ac'][6]:,.0f}",
                       f"{df5['curr_m_w3_tg'][6]:,.0f}", f"{df5['curr_m_w3_ac'][6]:,.0f}",
                       f"{df5['curr_m_w4_tg'][6]:,.0f}",
                       f"{df5['curr_m_w4_ac'][6]:,.0f}", f"{df5['curr_m_w5_tg'][6]:,.0f}",
                       f"{df5['curr_m_w5_ac'][6]:,.0f}",
                       f"{df5['full_month_tg'][6]:,.0f}", f"{df5['full_month_ac'][6]:,.0f}",
                       f"{df5['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df5['qtd_curr_w_flag'][6]}", f"{df5['curr_m_w1_flag'][6]}", f"{df5['curr_m_w2_flag'][6]}",
                       f"{df5['curr_m_w3_flag'][6]}", f"{df5['curr_m_w4_flag'][6]}", f"{df5['curr_m_w5_flag'][6]}",
                       f"{df5['full_month_flag'][6]}" ) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(5)[0], getHeaderTable(5)[1], getHeaderTable(5)[2],
                       getHeaderTable(5)[3], getHeaderTable(5)[4], getHeaderTable(5)[5],
                       getHeaderTable(5)[6], getHeaderTable(5)[7], getHeaderTable(5)[8],
                       getHeaderTable(5)[9], getHeaderTable(5)[10], getHeaderTable(5)[11],
                       # SDH
                       f"{df6['curr_q_q1tg'][0]:,.0f}", f"{df6['curr_q_q2tg'][0]:,.0f}",
                       f"{df6['curr_q_q3tg'][0]:,.0f}", f"{df6['curr_q_total'][0]:,.0f}",
                       f"{df6['qtd_curr_w_tg'][0]:,.0f}", f"{df6['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][0]:,.0f}", f"{df6['qtg_val'][0]:,.0f}",
                       f"{df6['curr_m_w1_tg'][0]:,.0f}", f"{df6['curr_m_w1_ac'][0]:,.0f}",
                       f"{df6['curr_m_w2_tg'][0]:,.0f}", f"{df6['curr_m_w2_ac'][0]:,.0f}",
                       f"{df6['curr_m_w3_tg'][0]:,.0f}", f"{df6['curr_m_w3_ac'][0]:,.0f}",
                       f"{df6['curr_m_w4_tg'][0]:,.0f}", f"{df6['curr_m_w4_ac'][0]:,.0f}",
                       f"{df6['curr_m_w5_tg'][0]:,.0f}", f"{df6['curr_m_w5_ac'][0]:,.0f}",
                       f"{df6['full_month_tg'][0]:,.0f}", f"{df6['full_month_ac'][0]:,.0f}",
                       f"{df6['full_month_dff'][0]:,.0f}",
                       f"{df6['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df6['curr_m_w1_flag'][0]}", f"{df6['curr_m_w2_flag'][0]}",
                       f"{df6['curr_m_w3_flag'][0]}", f"{df6['curr_m_w4_flag'][0]}",
                       f"{df6['curr_m_w5_flag'][0]}", f"{df6['full_month_flag'][0]}",
                       # TH
                       f"{df6['curr_q_q1tg'][1]:,.0f}", f"{df6['curr_q_q2tg'][1]:,.0f}",
                       f"{df6['curr_q_q3tg'][1]:,.0f}", f"{df6['curr_q_total'][1]:,.0f}",
                       f"{df6['qtd_curr_w_tg'][1]:,.0f}", f"{df6['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][1]:,.0f}", f"{df6['qtg_val'][1]:,.0f}",
                       f"{df6['curr_m_w1_tg'][1]:,.0f}", f"{df6['curr_m_w1_ac'][1]:,.0f}",
                       f"{df6['curr_m_w2_tg'][1]:,.0f}", f"{df6['curr_m_w2_ac'][1]:,.0f}",
                       f"{df6['curr_m_w3_tg'][1]:,.0f}", f"{df6['curr_m_w3_ac'][1]:,.0f}",
                       f"{df6['curr_m_w4_tg'][1]:,.0f}", f"{df6['curr_m_w4_ac'][1]:,.0f}",
                       f"{df6['curr_m_w5_tg'][1]:,.0f}", f"{df6['curr_m_w5_ac'][1]:,.0f}",
                       f"{df6['full_month_tg'][1]:,.0f}", f"{df6['full_month_ac'][1]:,.0f}",
                       f"{df6['full_month_dff'][1]:,.0f}", f"{df6['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df6['curr_m_w1_flag'][1]}", f"{df6['curr_m_w2_flag'][1]}", f"{df6['curr_m_w3_flag'][1]}",
                       f"{df6['curr_m_w4_flag'][1]}", f"{df6['curr_m_w5_flag'][1]}", f"{df6['full_month_flag'][1]}",
                       # TH BKM
                       f"{df6['curr_q_q1tg'][2]:,.0f}", f"{df6['curr_q_q2tg'][2]:,.0f}",
                       f"{df6['curr_q_q3tg'][2]:,.0f}",
                       f"{df6['curr_q_total'][2]:,.0f}", f"{df6['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df6['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][2]:,.0f}", f"{df6['qtg_val'][2]:,.0f}",
                       f"{df6['curr_m_w1_tg'][2]:,.0f}",
                       f"{df6['curr_m_w1_ac'][2]:,.0f}", f"{df6['curr_m_w2_tg'][2]:,.0f}",
                       f"{df6['curr_m_w2_ac'][2]:,.0f}",
                       f"{df6['curr_m_w3_tg'][2]:,.0f}", f"{df6['curr_m_w3_ac'][2]:,.0f}",
                       f"{df6['curr_m_w4_tg'][2]:,.0f}",
                       f"{df6['curr_m_w4_ac'][2]:,.0f}", f"{df6['curr_m_w5_tg'][2]:,.0f}",
                       f"{df6['curr_m_w5_ac'][2]:,.0f}",
                       f"{df6['full_month_tg'][2]:,.0f}", f"{df6['full_month_ac'][2]:,.0f}",
                       f"{df6['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df6['qtd_curr_w_flag'][2]}", f"{df6['curr_m_w1_flag'][2]}", f"{df6['curr_m_w2_flag'][2]}",
                       f"{df6['curr_m_w3_flag'][2]}", f"{df6['curr_m_w4_flag'][2]}", f"{df6['curr_m_w5_flag'][2]}",
                       f"{df6['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df6['curr_q_q1tg'][3]:,.0f}", f"{df6['curr_q_q2tg'][3]:,.0f}",
                       f"{df6['curr_q_q3tg'][3]:,.0f}",
                       f"{df6['curr_q_total'][3]:,.0f}", f"{df6['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df6['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][3]:,.0f}", f"{df6['qtg_val'][3]:,.0f}",
                       f"{df6['curr_m_w1_tg'][3]:,.0f}",
                       f"{df6['curr_m_w1_ac'][3]:,.0f}", f"{df6['curr_m_w2_tg'][3]:,.0f}",
                       f"{df6['curr_m_w2_ac'][3]:,.0f}",
                       f"{df6['curr_m_w3_tg'][3]:,.0f}", f"{df6['curr_m_w3_ac'][3]:,.0f}",
                       f"{df6['curr_m_w4_tg'][3]:,.0f}",
                       f"{df6['curr_m_w4_ac'][3]:,.0f}", f"{df6['curr_m_w5_tg'][3]:,.0f}",
                       f"{df6['curr_m_w5_ac'][3]:,.0f}",
                       f"{df6['full_month_tg'][3]:,.0f}", f"{df6['full_month_ac'][3]:,.0f}",
                       f"{df6['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df6['qtd_curr_w_flag'][3]}", f"{df6['curr_m_w1_flag'][3]}", f"{df6['curr_m_w2_flag'][3]}",
                       f"{df6['curr_m_w3_flag'][3]}", f"{df6['curr_m_w4_flag'][3]}", f"{df6['curr_m_w5_flag'][3]}",
                       f"{df6['full_month_flag'][3]}",
                       # CD1
                       f"{df6['curr_q_q1tg'][4]:,.0f}", f"{df6['curr_q_q2tg'][4]:,.0f}",
                       f"{df6['curr_q_q3tg'][4]:,.0f}",
                       f"{df6['curr_q_total'][4]:,.0f}", f"{df6['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df6['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][4]:,.0f}", f"{df6['qtg_val'][4]:,.0f}",
                       f"{df6['curr_m_w1_tg'][4]:,.0f}",
                       f"{df6['curr_m_w1_ac'][4]:,.0f}", f"{df6['curr_m_w2_tg'][4]:,.0f}",
                       f"{df6['curr_m_w2_ac'][4]:,.0f}",
                       f"{df6['curr_m_w3_tg'][4]:,.0f}", f"{df6['curr_m_w3_ac'][4]:,.0f}",
                       f"{df6['curr_m_w4_tg'][4]:,.0f}",
                       f"{df6['curr_m_w4_ac'][4]:,.0f}", f"{df6['curr_m_w5_tg'][4]:,.0f}",
                       f"{df6['curr_m_w5_ac'][4]:,.0f}",
                       f"{df6['full_month_tg'][4]:,.0f}", f"{df6['full_month_ac'][4]:,.0f}",
                       f"{df6['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df6['qtd_curr_w_flag'][4]}", f"{df6['curr_m_w1_flag'][4]}", f"{df6['curr_m_w2_flag'][4]}",
                       f"{df6['curr_m_w3_flag'][4]}", f"{df6['curr_m_w4_flag'][4]}", f"{df6['curr_m_w5_flag'][4]}",
                       f"{df6['full_month_flag'][4]}",
                       # CD2
                       f"{df6['curr_q_q1tg'][5]:,.0f}", f"{df6['curr_q_q2tg'][5]:,.0f}",
                       f"{df6['curr_q_q3tg'][5]:,.0f}",
                       f"{df6['curr_q_total'][5]:,.0f}", f"{df6['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df6['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][5]:,.0f}", f"{df6['qtg_val'][5]:,.0f}",
                       f"{df6['curr_m_w1_tg'][5]:,.0f}",
                       f"{df6['curr_m_w1_ac'][5]:,.0f}", f"{df6['curr_m_w2_tg'][5]:,.0f}",
                       f"{df6['curr_m_w2_ac'][5]:,.0f}",
                       f"{df6['curr_m_w3_tg'][5]:,.0f}", f"{df6['curr_m_w3_ac'][5]:,.0f}",
                       f"{df6['curr_m_w4_tg'][5]:,.0f}",
                       f"{df6['curr_m_w4_ac'][5]:,.0f}", f"{df6['curr_m_w5_tg'][5]:,.0f}",
                       f"{df6['curr_m_w5_ac'][5]:,.0f}",
                       f"{df6['full_month_tg'][5]:,.0f}", f"{df6['full_month_ac'][5]:,.0f}",
                       f"{df6['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df6['qtd_curr_w_flag'][5]}", f"{df6['curr_m_w1_flag'][5]}", f"{df6['curr_m_w2_flag'][5]}",
                       f"{df6['curr_m_w3_flag'][5]}", f"{df6['curr_m_w4_flag'][5]}", f"{df6['curr_m_w5_flag'][5]}",
                       f"{df6['full_month_flag'][5]}",
                       # Total
                       f"{df6['curr_q_q1tg'][6]:,.0f}", f"{df6['curr_q_q2tg'][6]:,.0f}",
                       f"{df6['curr_q_q3tg'][6]:,.0f}",
                       f"{df6['curr_q_total'][6]:,.0f}", f"{df6['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df6['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df6['qtd_curr_w_diff'][6]:,.0f}", f"{df6['qtg_val'][6]:,.0f}",
                       f"{df6['curr_m_w1_tg'][6]:,.0f}",
                       f"{df6['curr_m_w1_ac'][6]:,.0f}", f"{df6['curr_m_w2_tg'][6]:,.0f}",
                       f"{df6['curr_m_w2_ac'][6]:,.0f}",
                       f"{df6['curr_m_w3_tg'][6]:,.0f}", f"{df6['curr_m_w3_ac'][6]:,.0f}",
                       f"{df6['curr_m_w4_tg'][6]:,.0f}",
                       f"{df6['curr_m_w4_ac'][6]:,.0f}", f"{df6['curr_m_w5_tg'][6]:,.0f}",
                       f"{df6['curr_m_w5_ac'][6]:,.0f}",
                       f"{df6['full_month_tg'][6]:,.0f}", f"{df6['full_month_ac'][6]:,.0f}",
                       f"{df6['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df6['qtd_curr_w_flag'][6]}", f"{df6['curr_m_w1_flag'][6]}", f"{df6['curr_m_w2_flag'][6]}",
                       f"{df6['curr_m_w3_flag'][6]}", f"{df6['curr_m_w4_flag'][6]}", f"{df6['curr_m_w5_flag'][6]}",
                       f"{df6['full_month_flag'][6]}" ) \
           + "<br />" + readhtml(1) \
               .format(getHeaderTable(6)[0], getHeaderTable(6)[1], getHeaderTable(6)[2],
                       getHeaderTable(6)[3], getHeaderTable(6)[4], getHeaderTable(6)[5],
                       getHeaderTable(6)[6], getHeaderTable(6)[7], getHeaderTable(6)[8],
                       getHeaderTable(6)[9], getHeaderTable(6)[10], getHeaderTable(6)[11],
                       # SDH
                       f"{df7['curr_q_q1tg'][0]:,.0f}", f"{df7['curr_q_q2tg'][0]:,.0f}",
                       f"{df7['curr_q_q3tg'][0]:,.0f}", f"{df7['curr_q_total'][0]:,.0f}",
                       f"{df7['qtd_curr_w_tg'][0]:,.0f}", f"{df7['qtd_curr_w_ac'][0]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][0]:,.0f}", f"{df7['qtg_val'][0]:,.0f}",
                       f"{df7['curr_m_w1_tg'][0]:,.0f}", f"{df7['curr_m_w1_ac'][0]:,.0f}",
                       f"{df7['curr_m_w2_tg'][0]:,.0f}", f"{df7['curr_m_w2_ac'][0]:,.0f}",
                       f"{df7['curr_m_w3_tg'][0]:,.0f}", f"{df7['curr_m_w3_ac'][0]:,.0f}",
                       f"{df7['curr_m_w4_tg'][0]:,.0f}", f"{df7['curr_m_w4_ac'][0]:,.0f}",
                       f"{df7['curr_m_w5_tg'][0]:,.0f}", f"{df7['curr_m_w5_ac'][0]:,.0f}",
                       f"{df7['full_month_tg'][0]:,.0f}", f"{df7['full_month_ac'][0]:,.0f}",
                       f"{df7['full_month_dff'][0]:,.0f}",
                       f"{df7['qtd_curr_w_flag'][0]}",  # SDH Flag Color
                       f"{df7['curr_m_w1_flag'][0]}", f"{df7['curr_m_w2_flag'][0]}",
                       f"{df7['curr_m_w3_flag'][0]}", f"{df7['curr_m_w4_flag'][0]}",
                       f"{df7['curr_m_w5_flag'][0]}", f"{df7['full_month_flag'][0]}",
                       # TH
                       f"{df7['curr_q_q1tg'][1]:,.0f}", f"{df7['curr_q_q2tg'][1]:,.0f}",
                       f"{df7['curr_q_q3tg'][1]:,.0f}", f"{df7['curr_q_total'][1]:,.0f}",
                       f"{df7['qtd_curr_w_tg'][1]:,.0f}", f"{df7['qtd_curr_w_ac'][1]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][1]:,.0f}", f"{df7['qtg_val'][1]:,.0f}",
                       f"{df7['curr_m_w1_tg'][1]:,.0f}", f"{df7['curr_m_w1_ac'][1]:,.0f}",
                       f"{df7['curr_m_w2_tg'][1]:,.0f}", f"{df7['curr_m_w2_ac'][1]:,.0f}",
                       f"{df7['curr_m_w3_tg'][1]:,.0f}", f"{df7['curr_m_w3_ac'][1]:,.0f}",
                       f"{df7['curr_m_w4_tg'][1]:,.0f}", f"{df7['curr_m_w4_ac'][1]:,.0f}",
                       f"{df7['curr_m_w5_tg'][1]:,.0f}", f"{df7['curr_m_w5_ac'][1]:,.0f}",
                       f"{df7['full_month_tg'][1]:,.0f}", f"{df7['full_month_ac'][1]:,.0f}",
                       f"{df7['full_month_dff'][1]:,.0f}", f"{df7['qtd_curr_w_flag'][1]}",
                       # TH Flag Color
                       f"{df7['curr_m_w1_flag'][1]}", f"{df7['curr_m_w2_flag'][1]}", f"{df7['curr_m_w3_flag'][1]}",
                       f"{df7['curr_m_w4_flag'][1]}", f"{df7['curr_m_w5_flag'][1]}", f"{df7['full_month_flag'][1]}",
                       # TH BKM
                       f"{df7['curr_q_q1tg'][2]:,.0f}", f"{df7['curr_q_q2tg'][2]:,.0f}",
                       f"{df7['curr_q_q3tg'][2]:,.0f}",
                       f"{df7['curr_q_total'][2]:,.0f}", f"{df7['qtd_curr_w_tg'][2]:,.0f}",
                       f"{df7['qtd_curr_w_ac'][2]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][2]:,.0f}", f"{df7['qtg_val'][2]:,.0f}",
                       f"{df7['curr_m_w1_tg'][2]:,.0f}",
                       f"{df7['curr_m_w1_ac'][2]:,.0f}", f"{df7['curr_m_w2_tg'][2]:,.0f}",
                       f"{df7['curr_m_w2_ac'][2]:,.0f}",
                       f"{df7['curr_m_w3_tg'][2]:,.0f}", f"{df7['curr_m_w3_ac'][2]:,.0f}",
                       f"{df7['curr_m_w4_tg'][2]:,.0f}",
                       f"{df7['curr_m_w4_ac'][2]:,.0f}", f"{df7['curr_m_w5_tg'][2]:,.0f}",
                       f"{df7['curr_m_w5_ac'][2]:,.0f}",
                       f"{df7['full_month_tg'][2]:,.0f}", f"{df7['full_month_ac'][2]:,.0f}",
                       f"{df7['full_month_dff'][2]:,.0f}",
                       # TH BKM Color
                       f"{df7['qtd_curr_w_flag'][2]}", f"{df7['curr_m_w1_flag'][2]}", f"{df7['curr_m_w2_flag'][2]}",
                       f"{df7['curr_m_w3_flag'][2]}", f"{df7['curr_m_w4_flag'][2]}", f"{df7['curr_m_w5_flag'][2]}",
                       f"{df7['full_month_flag'][2]}",
                       # TH PLENO
                       f"{df7['curr_q_q1tg'][3]:,.0f}", f"{df7['curr_q_q2tg'][3]:,.0f}",
                       f"{df7['curr_q_q3tg'][3]:,.0f}",
                       f"{df7['curr_q_total'][3]:,.0f}", f"{df7['qtd_curr_w_tg'][3]:,.0f}",
                       f"{df7['qtd_curr_w_ac'][3]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][3]:,.0f}", f"{df7['qtg_val'][3]:,.0f}",
                       f"{df7['curr_m_w1_tg'][3]:,.0f}",
                       f"{df7['curr_m_w1_ac'][3]:,.0f}", f"{df7['curr_m_w2_tg'][3]:,.0f}",
                       f"{df7['curr_m_w2_ac'][3]:,.0f}",
                       f"{df7['curr_m_w3_tg'][3]:,.0f}", f"{df7['curr_m_w3_ac'][3]:,.0f}",
                       f"{df7['curr_m_w4_tg'][3]:,.0f}",
                       f"{df7['curr_m_w4_ac'][3]:,.0f}", f"{df7['curr_m_w5_tg'][3]:,.0f}",
                       f"{df7['curr_m_w5_ac'][3]:,.0f}",
                       f"{df7['full_month_tg'][3]:,.0f}", f"{df7['full_month_ac'][3]:,.0f}",
                       f"{df7['full_month_dff'][3]:,.0f}",
                       # TH Pleno Color
                       f"{df7['qtd_curr_w_flag'][3]}", f"{df7['curr_m_w1_flag'][3]}", f"{df7['curr_m_w2_flag'][3]}",
                       f"{df7['curr_m_w3_flag'][3]}", f"{df7['curr_m_w4_flag'][3]}", f"{df7['curr_m_w5_flag'][3]}",
                       f"{df7['full_month_flag'][3]}",
                       # CD1
                       f"{df7['curr_q_q1tg'][4]:,.0f}", f"{df7['curr_q_q2tg'][4]:,.0f}",
                       f"{df7['curr_q_q3tg'][4]:,.0f}",
                       f"{df7['curr_q_total'][4]:,.0f}", f"{df7['qtd_curr_w_tg'][4]:,.0f}",
                       f"{df7['qtd_curr_w_ac'][4]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][4]:,.0f}", f"{df7['qtg_val'][4]:,.0f}",
                       f"{df7['curr_m_w1_tg'][4]:,.0f}",
                       f"{df7['curr_m_w1_ac'][4]:,.0f}", f"{df7['curr_m_w2_tg'][4]:,.0f}",
                       f"{df7['curr_m_w2_ac'][4]:,.0f}",
                       f"{df7['curr_m_w3_tg'][4]:,.0f}", f"{df7['curr_m_w3_ac'][4]:,.0f}",
                       f"{df7['curr_m_w4_tg'][4]:,.0f}",
                       f"{df7['curr_m_w4_ac'][4]:,.0f}", f"{df7['curr_m_w5_tg'][4]:,.0f}",
                       f"{df7['curr_m_w5_ac'][4]:,.0f}",
                       f"{df7['full_month_tg'][4]:,.0f}", f"{df7['full_month_ac'][4]:,.0f}",
                       f"{df7['full_month_dff'][4]:,.0f}",
                       # CD1 Color
                       f"{df7['qtd_curr_w_flag'][4]}", f"{df7['curr_m_w1_flag'][4]}", f"{df7['curr_m_w2_flag'][4]}",
                       f"{df7['curr_m_w3_flag'][4]}", f"{df7['curr_m_w4_flag'][4]}", f"{df7['curr_m_w5_flag'][4]}",
                       f"{df7['full_month_flag'][4]}",
                       # CD2
                       f"{df7['curr_q_q1tg'][5]:,.0f}", f"{df7['curr_q_q2tg'][5]:,.0f}",
                       f"{df7['curr_q_q3tg'][5]:,.0f}",
                       f"{df7['curr_q_total'][5]:,.0f}", f"{df7['qtd_curr_w_tg'][5]:,.0f}",
                       f"{df7['qtd_curr_w_ac'][5]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][5]:,.0f}", f"{df7['qtg_val'][5]:,.0f}",
                       f"{df7['curr_m_w1_tg'][5]:,.0f}",
                       f"{df7['curr_m_w1_ac'][5]:,.0f}", f"{df7['curr_m_w2_tg'][5]:,.0f}",
                       f"{df7['curr_m_w2_ac'][5]:,.0f}",
                       f"{df7['curr_m_w3_tg'][5]:,.0f}", f"{df7['curr_m_w3_ac'][5]:,.0f}",
                       f"{df7['curr_m_w4_tg'][5]:,.0f}",
                       f"{df7['curr_m_w4_ac'][5]:,.0f}", f"{df7['curr_m_w5_tg'][5]:,.0f}",
                       f"{df7['curr_m_w5_ac'][5]:,.0f}",
                       f"{df7['full_month_tg'][5]:,.0f}", f"{df7['full_month_ac'][5]:,.0f}",
                       f"{df7['full_month_dff'][5]:,.0f}",
                       # CD2 Color
                       f"{df7['qtd_curr_w_flag'][5]}", f"{df7['curr_m_w1_flag'][5]}", f"{df7['curr_m_w2_flag'][5]}",
                       f"{df7['curr_m_w3_flag'][5]}", f"{df7['curr_m_w4_flag'][5]}", f"{df7['curr_m_w5_flag'][5]}",
                       f"{df7['full_month_flag'][5]}",
                       # Total
                       f"{df7['curr_q_q1tg'][6]:,.0f}", f"{df7['curr_q_q2tg'][6]:,.0f}",
                       f"{df7['curr_q_q3tg'][6]:,.0f}",
                       f"{df7['curr_q_total'][6]:,.0f}", f"{df7['qtd_curr_w_tg'][6]:,.0f}",
                       f"{df7['qtd_curr_w_ac'][6]:,.0f}",
                       f"{df7['qtd_curr_w_diff'][6]:,.0f}", f"{df7['qtg_val'][6]:,.0f}",
                       f"{df7['curr_m_w1_tg'][6]:,.0f}",
                       f"{df7['curr_m_w1_ac'][6]:,.0f}", f"{df7['curr_m_w2_tg'][6]:,.0f}",
                       f"{df7['curr_m_w2_ac'][6]:,.0f}",
                       f"{df7['curr_m_w3_tg'][6]:,.0f}", f"{df7['curr_m_w3_ac'][6]:,.0f}",
                       f"{df7['curr_m_w4_tg'][6]:,.0f}",
                       f"{df7['curr_m_w4_ac'][6]:,.0f}", f"{df7['curr_m_w5_tg'][6]:,.0f}",
                       f"{df7['curr_m_w5_ac'][6]:,.0f}",
                       f"{df7['full_month_tg'][6]:,.0f}", f"{df7['full_month_ac'][6]:,.0f}",
                       f"{df7['full_month_dff'][6]:,.0f}",
                       # Total Color
                       f"{df7['qtd_curr_w_flag'][6]}", f"{df7['curr_m_w1_flag'][6]}", f"{df7['curr_m_w2_flag'][6]}",
                       f"{df7['curr_m_w3_flag'][6]}", f"{df7['curr_m_w4_flag'][6]}", f"{df7['curr_m_w5_flag'][6]}",
                       f"{df7['full_month_flag'][6]}" ) \
            + "<br />" + readhtml(2) \
            .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
                    getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
                    getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
                    getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11],
                    # Condo AP
                    f"{df8['curr_q_q1tg'][0]:,.0f}", f"{df8['curr_q_q2tg'][0]:,.0f}",
                    f"{df8['curr_q_q3tg'][0]:,.0f}", f"{df8['curr_q_total'][0]:,.0f}",
                    f"{df8['qtd_curr_w_tg'][0]:,.0f}", f"{df8['qtd_curr_w_ac'][0]:,.0f}",
                    f"{df8['qtd_curr_w_diff'][0]:,.0f}", f"{df8['qtg_val'][0]:,.0f}",
                    f"{df8['curr_m_w1_tg'][0]:,.0f}", f"{df8['curr_m_w1_ac'][0]:,.0f}",
                    f"{df8['curr_m_w2_tg'][0]:,.0f}", f"{df8['curr_m_w2_ac'][0]:,.0f}",
                    f"{df8['curr_m_w3_tg'][0]:,.0f}", f"{df8['curr_m_w3_ac'][0]:,.0f}",
                    f"{df8['curr_m_w4_tg'][0]:,.0f}", f"{df8['curr_m_w4_ac'][0]:,.0f}",
                    f"{df8['curr_m_w5_tg'][0]:,.0f}", f"{df8['curr_m_w5_ac'][0]:,.0f}",
                    f"{df8['full_month_tg'][0]:,.0f}", f"{df8['full_month_ac'][0]:,.0f}",
                    f"{df8['full_month_dff'][0]:,.0f}",
                    # Condo AP Flag Color
                    f"{df8['qtd_curr_w_flag'][0]}",
                    f"{df8['curr_m_w1_flag'][0]}", f"{df8['curr_m_w2_flag'][0]}",
                    f"{df8['curr_m_w3_flag'][0]}", f"{df8['curr_m_w4_flag'][0]}",
                    f"{df8['curr_m_w5_flag'][0]}", f"{df8['full_month_flag'][0]}",
                    # CD1
                    f"{df8['curr_q_q1tg'][1]:,.0f}", f"{df8['curr_q_q2tg'][1]:,.0f}",
                    f"{df8['curr_q_q3tg'][1]:,.0f}", f"{df8['curr_q_total'][1]:,.0f}",
                    f"{df8['qtd_curr_w_tg'][1]:,.0f}", f"{df8['qtd_curr_w_ac'][1]:,.0f}",
                    f"{df8['qtd_curr_w_diff'][1]:,.0f}", f"{df8['qtg_val'][1]:,.0f}",
                    f"{df8['curr_m_w1_tg'][1]:,.0f}", f"{df8['curr_m_w1_ac'][1]:,.0f}",
                    f"{df8['curr_m_w2_tg'][1]:,.0f}", f"{df8['curr_m_w2_ac'][1]:,.0f}",
                    f"{df8['curr_m_w3_tg'][1]:,.0f}", f"{df8['curr_m_w3_ac'][1]:,.0f}",
                    f"{df8['curr_m_w4_tg'][1]:,.0f}", f"{df8['curr_m_w4_ac'][1]:,.0f}",
                    f"{df8['curr_m_w5_tg'][1]:,.0f}", f"{df8['curr_m_w5_ac'][1]:,.0f}",
                    f"{df8['full_month_tg'][1]:,.0f}", f"{df8['full_month_ac'][1]:,.0f}",
                    f"{df8['full_month_dff'][1]:,.0f}",
                    # CD1 Flag Color
                    f"{df8['qtd_curr_w_flag'][1]}",
                    f"{df8['curr_m_w1_flag'][1]}", f"{df8['curr_m_w2_flag'][1]}", f"{df8['curr_m_w3_flag'][1]}",
                    f"{df8['curr_m_w4_flag'][1]}", f"{df8['curr_m_w5_flag'][1]}", f"{df8['full_month_flag'][1]}",
                    # CD2
                    f"{df8['curr_q_q1tg'][2]:,.0f}", f"{df8['curr_q_q2tg'][2]:,.0f}", f"{df8['curr_q_q3tg'][2]:,.0f}",
                    f"{df8['curr_q_total'][2]:,.0f}", f"{df8['qtd_curr_w_tg'][2]:,.0f}",
                    f"{df8['qtd_curr_w_ac'][2]:,.0f}",
                    f"{df8['qtd_curr_w_diff'][2]:,.0f}", f"{df8['qtg_val'][2]:,.0f}", f"{df8['curr_m_w1_tg'][2]:,.0f}",
                    f"{df8['curr_m_w1_ac'][2]:,.0f}", f"{df8['curr_m_w2_tg'][2]:,.0f}",
                    f"{df8['curr_m_w2_ac'][2]:,.0f}",
                    f"{df8['curr_m_w3_tg'][2]:,.0f}", f"{df8['curr_m_w3_ac'][2]:,.0f}",
                    f"{df8['curr_m_w4_tg'][2]:,.0f}",
                    f"{df8['curr_m_w4_ac'][2]:,.0f}", f"{df8['curr_m_w5_tg'][2]:,.0f}",
                    f"{df8['curr_m_w5_ac'][2]:,.0f}",
                    f"{df8['full_month_tg'][2]:,.0f}", f"{df8['full_month_ac'][2]:,.0f}",
                    f"{df8['full_month_dff'][2]:,.0f}",
                    # CD2 Color
                    f"{df8['qtd_curr_w_flag'][2]}", f"{df8['curr_m_w1_flag'][2]}", f"{df8['curr_m_w2_flag'][2]}",
                    f"{df8['curr_m_w3_flag'][2]}", f"{df8['curr_m_w4_flag'][2]}", f"{df8['curr_m_w5_flag'][2]}",
                    f"{df8['full_month_flag'][2]}",
                    # Total
                    f"{df8['curr_q_q1tg'][3]:,.0f}", f"{df8['curr_q_q2tg'][3]:,.0f}", f"{df8['curr_q_q3tg'][3]:,.0f}",
                    f"{df8['curr_q_total'][3]:,.0f}", f"{df8['qtd_curr_w_tg'][3]:,.0f}",
                    f"{df8['qtd_curr_w_ac'][3]:,.0f}",
                    f"{df8['qtd_curr_w_diff'][3]:,.0f}", f"{df8['qtg_val'][3]:,.0f}", f"{df8['curr_m_w1_tg'][3]:,.0f}",
                    f"{df8['curr_m_w1_ac'][3]:,.0f}", f"{df8['curr_m_w2_tg'][3]:,.0f}",
                    f"{df8['curr_m_w2_ac'][3]:,.0f}",
                    f"{df8['curr_m_w3_tg'][3]:,.0f}", f"{df8['curr_m_w3_ac'][3]:,.0f}",
                    f"{df8['curr_m_w4_tg'][3]:,.0f}",
                    f"{df8['curr_m_w4_ac'][3]:,.0f}", f"{df8['curr_m_w5_tg'][3]:,.0f}",
                    f"{df8['curr_m_w5_ac'][3]:,.0f}",
                    f"{df8['full_month_tg'][3]:,.0f}", f"{df8['full_month_ac'][3]:,.0f}",
                    f"{df8['full_month_dff'][3]:,.0f}",
                    # Total Color
                    f"{df8['qtd_curr_w_flag'][3]}", f"{df8['curr_m_w1_flag'][3]}", f"{df8['curr_m_w2_flag'][3]}",
                    f"{df8['curr_m_w3_flag'][3]}", f"{df8['curr_m_w4_flag'][3]}", f"{df8['curr_m_w5_flag'][3]}",
                    f"{df8['full_month_flag'][3]}" ) \
        + "<br />" + readhtml(3) \
            .format(getHeaderTable(7)[0], getHeaderTable(7)[1], getHeaderTable(7)[2],
                    getHeaderTable(7)[3], getHeaderTable(7)[4], getHeaderTable(7)[5],
                    getHeaderTable(7)[6], getHeaderTable(7)[7], getHeaderTable(7)[8],
                    getHeaderTable(7)[9], getHeaderTable(7)[10], getHeaderTable(7)[11],
                    # Pre-sale
                    f"{df9['curr_q_q1tg'][0]:,.0f}", f"{df9['curr_q_q2tg'][0]:,.0f}",
                    f"{df9['curr_q_q3tg'][0]:,.0f}", f"{df9['curr_q_total'][0]:,.0f}",
                    f"{df9['qtd_curr_w_tg'][0]:,.0f}", f"{df9['qtd_curr_w_ac'][0]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][0]:,.0f}", f"{df9['qtg_val'][0]:,.0f}",
                    f"{df9['curr_m_w1_tg'][0]:,.0f}", f"{df9['curr_m_w1_ac'][0]:,.0f}",
                    f"{df9['curr_m_w2_tg'][0]:,.0f}", f"{df9['curr_m_w2_ac'][0]:,.0f}",
                    f"{df9['curr_m_w3_tg'][0]:,.0f}", f"{df9['curr_m_w3_ac'][0]:,.0f}",
                    f"{df9['curr_m_w4_tg'][0]:,.0f}", f"{df9['curr_m_w4_ac'][0]:,.0f}",
                    f"{df9['curr_m_w5_tg'][0]:,.0f}", f"{df9['curr_m_w5_ac'][0]:,.0f}",
                    f"{df9['full_month_tg'][0]:,.0f}", f"{df9['full_month_ac'][0]:,.0f}",
                    f"{df9['full_month_dff'][0]:,.0f}",
                    # Pre-sale Flag Color
                    f"{df9['qtd_curr_w_flag'][0]}",
                    f"{df9['curr_m_w1_flag'][0]}", f"{df9['curr_m_w2_flag'][0]}",
                    f"{df9['curr_m_w3_flag'][0]}", f"{df9['curr_m_w4_flag'][0]}",
                    f"{df9['curr_m_w5_flag'][0]}", f"{df9['full_month_flag'][0]}",
                    # CD1
                    f"{df9['curr_q_q1tg'][1]:,.0f}", f"{df9['curr_q_q2tg'][1]:,.0f}",
                    f"{df9['curr_q_q3tg'][1]:,.0f}", f"{df9['curr_q_total'][1]:,.0f}",
                    f"{df9['qtd_curr_w_tg'][1]:,.0f}", f"{df9['qtd_curr_w_ac'][1]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][1]:,.0f}", f"{df9['qtg_val'][1]:,.0f}",
                    f"{df9['curr_m_w1_tg'][1]:,.0f}", f"{df9['curr_m_w1_ac'][1]:,.0f}",
                    f"{df9['curr_m_w2_tg'][1]:,.0f}", f"{df9['curr_m_w2_ac'][1]:,.0f}",
                    f"{df9['curr_m_w3_tg'][1]:,.0f}", f"{df9['curr_m_w3_ac'][1]:,.0f}",
                    f"{df9['curr_m_w4_tg'][1]:,.0f}", f"{df9['curr_m_w4_ac'][1]:,.0f}",
                    f"{df9['curr_m_w5_tg'][1]:,.0f}", f"{df9['curr_m_w5_ac'][1]:,.0f}",
                    f"{df9['full_month_tg'][1]:,.0f}", f"{df9['full_month_ac'][1]:,.0f}",
                    f"{df9['full_month_dff'][1]:,.0f}",
                    # CD1 Flag Color
                    f"{df9['qtd_curr_w_flag'][1]}",
                    f"{df9['curr_m_w1_flag'][1]}", f"{df9['curr_m_w2_flag'][1]}", f"{df9['curr_m_w3_flag'][1]}",
                    f"{df9['curr_m_w4_flag'][1]}", f"{df9['curr_m_w5_flag'][1]}", f"{df9['full_month_flag'][1]}",
                    # CD2
                    f"{df9['curr_q_q1tg'][2]:,.0f}", f"{df9['curr_q_q2tg'][2]:,.0f}", f"{df9['curr_q_q3tg'][2]:,.0f}",
                    f"{df9['curr_q_total'][2]:,.0f}", f"{df9['qtd_curr_w_tg'][2]:,.0f}",
                    f"{df9['qtd_curr_w_ac'][2]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][2]:,.0f}", f"{df9['qtg_val'][2]:,.0f}", f"{df9['curr_m_w1_tg'][2]:,.0f}",
                    f"{df9['curr_m_w1_ac'][2]:,.0f}", f"{df9['curr_m_w2_tg'][2]:,.0f}",
                    f"{df9['curr_m_w2_ac'][2]:,.0f}",
                    f"{df9['curr_m_w3_tg'][2]:,.0f}", f"{df9['curr_m_w3_ac'][2]:,.0f}",
                    f"{df9['curr_m_w4_tg'][2]:,.0f}",
                    f"{df9['curr_m_w4_ac'][2]:,.0f}", f"{df9['curr_m_w5_tg'][2]:,.0f}",
                    f"{df9['curr_m_w5_ac'][2]:,.0f}",
                    f"{df9['full_month_tg'][2]:,.0f}", f"{df9['full_month_ac'][2]:,.0f}",
                    f"{df9['full_month_dff'][2]:,.0f}",
                    # CD2 Color
                    f"{df9['qtd_curr_w_flag'][2]}", f"{df9['curr_m_w1_flag'][2]}", f"{df9['curr_m_w2_flag'][2]}",
                    f"{df9['curr_m_w3_flag'][2]}", f"{df9['curr_m_w4_flag'][2]}", f"{df9['curr_m_w5_flag'][2]}",
                    f"{df9['full_month_flag'][2]}",
                    # Transfer
                    f"{df9['curr_q_q1tg'][3]:,.0f}", f"{df9['curr_q_q2tg'][3]:,.0f}", f"{df9['curr_q_q3tg'][3]:,.0f}",
                    f"{df9['curr_q_total'][3]:,.0f}", f"{df9['qtd_curr_w_tg'][3]:,.0f}",
                    f"{df9['qtd_curr_w_ac'][3]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][3]:,.0f}", f"{df9['qtg_val'][3]:,.0f}", f"{df9['curr_m_w1_tg'][3]:,.0f}",
                    f"{df9['curr_m_w1_ac'][3]:,.0f}", f"{df9['curr_m_w2_tg'][3]:,.0f}",
                    f"{df9['curr_m_w2_ac'][3]:,.0f}",
                    f"{df9['curr_m_w3_tg'][3]:,.0f}", f"{df9['curr_m_w3_ac'][3]:,.0f}",
                    f"{df9['curr_m_w4_tg'][3]:,.0f}",
                    f"{df9['curr_m_w4_ac'][3]:,.0f}", f"{df9['curr_m_w5_tg'][3]:,.0f}",
                    f"{df9['curr_m_w5_ac'][3]:,.0f}",
                    f"{df9['full_month_tg'][3]:,.0f}", f"{df9['full_month_ac'][3]:,.0f}",
                    f"{df9['full_month_dff'][3]:,.0f}",
                    # Transfer Color
                    f"{df9['qtd_curr_w_flag'][3]}", f"{df9['curr_m_w1_flag'][3]}", f"{df9['curr_m_w2_flag'][3]}",
                    f"{df9['curr_m_w3_flag'][3]}", f"{df9['curr_m_w4_flag'][3]}", f"{df9['curr_m_w5_flag'][3]}",
                    f"{df9['full_month_flag'][3]}",
                    # CD1
                    f"{df9['curr_q_q1tg'][4]:,.0f}", f"{df9['curr_q_q2tg'][4]:,.0f}", f"{df9['curr_q_q3tg'][4]:,.0f}",
                    f"{df9['curr_q_total'][4]:,.0f}", f"{df9['qtd_curr_w_tg'][4]:,.0f}",
                    f"{df9['qtd_curr_w_ac'][4]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][4]:,.0f}", f"{df9['qtg_val'][4]:,.0f}", f"{df9['curr_m_w1_tg'][4]:,.0f}",
                    f"{df9['curr_m_w1_ac'][4]:,.0f}", f"{df9['curr_m_w2_tg'][4]:,.0f}",
                    f"{df9['curr_m_w2_ac'][4]:,.0f}",
                    f"{df9['curr_m_w3_tg'][4]:,.0f}", f"{df9['curr_m_w3_ac'][4]:,.0f}",
                    f"{df9['curr_m_w4_tg'][4]:,.0f}",
                    f"{df9['curr_m_w4_ac'][4]:,.0f}", f"{df9['curr_m_w5_tg'][4]:,.0f}",
                    f"{df9['curr_m_w5_ac'][4]:,.0f}",
                    f"{df9['full_month_tg'][4]:,.0f}", f"{df9['full_month_ac'][4]:,.0f}",
                    f"{df9['full_month_dff'][4]:,.0f}",
                    # CD1 Color
                    f"{df9['qtd_curr_w_flag'][4]}", f"{df9['curr_m_w1_flag'][4]}", f"{df9['curr_m_w2_flag'][4]}",
                    f"{df9['curr_m_w3_flag'][4]}", f"{df9['curr_m_w4_flag'][4]}", f"{df9['curr_m_w5_flag'][4]}",
                    f"{df9['full_month_flag'][4]}",
                    # CD2
                    f"{df9['curr_q_q1tg'][5]:,.0f}", f"{df9['curr_q_q2tg'][5]:,.0f}", f"{df9['curr_q_q3tg'][5]:,.0f}",
                    f"{df9['curr_q_total'][5]:,.0f}", f"{df9['qtd_curr_w_tg'][5]:,.0f}",
                    f"{df9['qtd_curr_w_ac'][5]:,.0f}",
                    f"{df9['qtd_curr_w_diff'][5]:,.0f}", f"{df9['qtg_val'][5]:,.0f}", f"{df9['curr_m_w1_tg'][5]:,.0f}",
                    f"{df9['curr_m_w1_ac'][5]:,.0f}", f"{df9['curr_m_w2_tg'][5]:,.0f}",
                    f"{df9['curr_m_w2_ac'][5]:,.0f}",
                    f"{df9['curr_m_w3_tg'][5]:,.0f}", f"{df9['curr_m_w3_ac'][5]:,.0f}",
                    f"{df9['curr_m_w4_tg'][5]:,.0f}",
                    f"{df9['curr_m_w4_ac'][5]:,.0f}", f"{df9['curr_m_w5_tg'][5]:,.0f}",
                    f"{df9['curr_m_w5_ac'][5]:,.0f}",
                    f"{df9['full_month_tg'][5]:,.0f}", f"{df9['full_month_ac'][5]:,.0f}",
                    f"{df9['full_month_dff'][5]:,.0f}",
                    # CD2 Color
                    f"{df9['qtd_curr_w_flag'][5]}", f"{df9['curr_m_w1_flag'][5]}", f"{df9['curr_m_w2_flag'][5]}",
                    f"{df9['curr_m_w3_flag'][5]}", f"{df9['curr_m_w4_flag'][5]}", f"{df9['curr_m_w5_flag'][5]}",
                    f"{df9['full_month_flag'][5]}") \
           # + "<br />" + readhtml(2) \
           #     .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
           #             getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
           #             getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
           #             getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
           # + "<br />" + readhtml(3) \
           #     .format(getHeaderTable(7)[0], getHeaderTable(7)[1], getHeaderTable(7)[2],
           #             getHeaderTable(7)[3], getHeaderTable(7)[4], getHeaderTable(7)[5],
           #             getHeaderTable(7)[6], getHeaderTable(7)[7], getHeaderTable(7)[8],
           #             getHeaderTable(7)[9], getHeaderTable(7)[10], getHeaderTable(7)[11]) \
           # + "<br />"
    else:
        return readhtmlw4(1) \
                   .format(getHeaderTable(0)[0], getHeaderTable(0)[1], getHeaderTable(0)[2],
                           getHeaderTable(0)[3], getHeaderTable(0)[4], getHeaderTable(0)[5],
                           getHeaderTable(0)[6], getHeaderTable(0)[7], getHeaderTable(0)[8],
                           getHeaderTable(0)[9], getHeaderTable(0)[10], getHeaderTable(0)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(1)[0], getHeaderTable(1)[1], getHeaderTable(1)[2],
               #             getHeaderTable(1)[3], getHeaderTable(1)[4], getHeaderTable(1)[5],
               #             getHeaderTable(1)[6], getHeaderTable(1)[7], getHeaderTable(1)[8],
               #             getHeaderTable(1)[9], getHeaderTable(1)[10], getHeaderTable(1)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(2)[0], getHeaderTable(2)[1], getHeaderTable(2)[2],
               #             getHeaderTable(2)[3], getHeaderTable(2)[4], getHeaderTable(2)[5],
               #             getHeaderTable(2)[6], getHeaderTable(2)[7], getHeaderTable(2)[8],
               #             getHeaderTable(2)[9], getHeaderTable(2)[10], getHeaderTable(2)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
               #             getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
               #             getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
               #             getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(4)[0], getHeaderTable(4)[1], getHeaderTable(4)[2],
               #             getHeaderTable(4)[3], getHeaderTable(4)[4], getHeaderTable(4)[5],
               #             getHeaderTable(4)[6], getHeaderTable(4)[7], getHeaderTable(4)[8],
               #             getHeaderTable(4)[9], getHeaderTable(4)[10], getHeaderTable(4)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(5)[0], getHeaderTable(5)[1], getHeaderTable(5)[2],
               #             getHeaderTable(5)[3], getHeaderTable(5)[4], getHeaderTable(5)[5],
               #             getHeaderTable(5)[6], getHeaderTable(5)[7], getHeaderTable(5)[8],
               #             getHeaderTable(5)[9], getHeaderTable(5)[10], getHeaderTable(5)[11]) \
               # + "<br />" + readhtmlw4(1) \
               #     .format(getHeaderTable(6)[0], getHeaderTable(6)[1], getHeaderTable(6)[2],
               #             getHeaderTable(6)[3], getHeaderTable(6)[4], getHeaderTable(6)[5],
               #             getHeaderTable(6)[6], getHeaderTable(6)[7], getHeaderTable(6)[8],
               #             getHeaderTable(6)[9], getHeaderTable(6)[10], getHeaderTable(6)[11]) \
               # + "<br />" + readhtmlw4(2) \
               #     .format(getHeaderTable(3)[0], getHeaderTable(3)[1], getHeaderTable(3)[2],
               #             getHeaderTable(3)[3], getHeaderTable(3)[4], getHeaderTable(3)[5],
               #             getHeaderTable(3)[6], getHeaderTable(3)[7], getHeaderTable(3)[8],
               #             getHeaderTable(3)[9], getHeaderTable(3)[10], getHeaderTable(3)[11]) \
               # + "<br />" + readhtmlw4(3) \
               #     .format(getHeaderTable(7)[0], getHeaderTable(7)[1], getHeaderTable(7)[2],
               #             getHeaderTable(7)[3], getHeaderTable(7)[4], getHeaderTable(7)[5],
               #             getHeaderTable(7)[6], getHeaderTable(7)[7], getHeaderTable(7)[8],
               #             getHeaderTable(7)[9], getHeaderTable(7)[10], getHeaderTable(7)[11]) \
               # + "<br />"


if __name__ == '__main__':
    main()
