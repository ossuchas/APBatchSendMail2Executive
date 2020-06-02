from datetime import datetime
# Mail Setting
# Production
MAIL_TO=["anuphong@apthai.com","pichet@apthai.com","vittakarn@apthai.com","visanu@apthai.com","pamorn@apthai.com","kamolthip@apthai.com","prajark@apthai.com","worrapong@apthai.com","wason@apthai.com","opas@apthai.com","ratchayud@apthai.com","tippawan_s@apthai.com","somchai_w@apthai.com","Risk Management Riskmanagement@apthai.com"]
MAIL_CC=["kultipa_t@apthai.com","plotch@apthai.com","bhunnapa@apthai.com","napaporn_ja@apthai.com","woraphan_c@apthai.com","laddawan_v@apthai.com","raweewan_p@apthai.com","srisakul_p@apthai.com","thirata_t@apthai.com","wipawan_k@apthai.com","panwarin_j@apthai.com","jatuporn_p@apthai.com","suchat_s@apthai.com","jintana_i@apthai.com","apichaya@apthai.com","tanonchai@apthai.com"]
# Dev
<<<<<<< HEAD
MAIL_TO=["apichaya@apthai.com", "polwaritpakorn@apthai.com"]
MAIL_CC=["suchat_s@apthai.com", "jintana_i@apthai.com"]
=======
#MAIL_TO=["apichaya@apthai.com", "suchat_s@apthai.com", "jintana_i@apthai.com","tanonchai@apthai.com","polwaritpakorn@apthai.com"]
#MAIL_CC=["suchat_s@apthai.com"]
>>>>>>> 7e10f8c2aaffc650b1e232383c976e8eadb656b4
MAIL_SENDER = "crmconsult@apthai.com"
MAIL_SUBJECT = f"Weekly Summary Lead Lag at [{datetime.now().strftime('%d/%m/%Y')}]"
MAIL_BODY = """
    """


# Mail Settting Estimate Revenue
EST_MAIL_SUBJECT = "Weekly Summary Estimate Revenue"
EST_MAIL_BODY = """
<body>
<table cellpadding="0" cellspacing="0" border="1" style="font-family:AP;font-size: 8.5pt;border: 1px solid black;border-collapse: collapse;">
  <col width="166" />
  <col width="61" span="3" />
  <col width="70" />
  <tr>
    <td width="166" rowspan="3" align="center" bgcolor="#CCCCCC"><strong>Transfer</strong></td>
    <td colspan="8" align="center" bgcolor="#66FFFF"><strong>Q1</strong></td>
    <td width="70" rowspan="3" align="center" bgcolor="#6699FF"><strong>QTD</strong></td>
    <td colspan="6" align="center" bgcolor="#FFCCFF"><strong>Estimate Income</strong></td>
  </tr>
  <tr>
    <td colspan="2" align="center" bgcolor="#66FFFF"><strong>Jan</strong></td>
    <td colspan="2" align="center" bgcolor="#66FFFF"><strong>Feb</strong></td>
    <td colspan="2" align="center" bgcolor="#66FFFF"><strong>Mar</strong></td>
    <td colspan="2" align="center" bgcolor="#66FFFF"><strong>Total Q1</strong></td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est</strong>#1</td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est</strong>#2</td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est</strong>#3</td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est</strong>#4</td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est#5</strong></td>
    <td width="70" rowspan="2" align="center" bgcolor="#FFCCFF"><strong>Est QTD</strong></td>
  </tr>
  <tr>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Target</strong></td>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Actual</strong></td>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Target</strong></td>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Actual</strong></td>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Target</strong></td>
    <td width="61" align="center" bgcolor="#66FFFF"><strong>Actual</strong></td>
    <td width="70" align="center" bgcolor="#66FFFF"><strong>Target</strong></td>
    <td width="70" align="center" bgcolor="#66FFFF"><strong>Actual</strong></td>
  </tr>
  <tr>
    <td align="left"><strong>AP (Excl.JV)</strong></td>
    <td align="right" >&nbsp;</td>
    <td align="right" >&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>JV (100%)</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><p><em><strong>Total</strong></em></p></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left" bgcolor="#999999"><p>&nbsp;</p></td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG1-SDH</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG2-BKM &amp;Pleno</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG3-Condo1 (Excl.JV)</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG4-Condo2 (Excl.JV)</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG3-JV (100%)</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="left"><strong>BG4-JV (100%)</strong></td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td align="right">&nbsp;</td>
  </tr>
</table>
</body>
    """
EST_MAIL_TABLE_HEAD = """
<table cellpadding="0" cellspacing="0" border="1" style="font-family:AP;font-size: 8.5pt;border: 1px solid black;border-collapse: collapse;">
  <col width="166" />
  <col width="61" span="3" />
  <col width="70" />
  """
EST_MAIL_TABLE_COL_HEAD = """
<tr>
    <td width="166" rowspan="3" align="center" bgcolor="#CCCCCC"><strong>Transfer</strong></td>
    <td colspan="4" align="center" bgcolor="#CCCCCC"><strong>Q1</strong></td>
  </tr>
  <tr>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Jan</strong></td>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Feb</strong></td>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Mar</strong></td>
    <td width="70" rowspan="2" align="right" bgcolor="#CCCCCC"><strong>Total Q1</strong></td>
  </tr>
  <tr>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Target</strong></td>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Target</strong></td>
    <td width="61" align="center" bgcolor="#CCCCCC"><strong>Target</strong></td>
  </tr>
"""
