import os
# Mail Setting
MAIL_SENDER = "noreply@apthai.com"
MAIL_SUBJECT = "Weekly Summary Lead Lag"
MAIL_BODY = """
<body>
<table border="1" cellpadding="0" cellspacing="0" style="font-family:Calibri;font-size: 9pt;">
  <col width="166" />
  <col width="61" span="3" />
  <col width="70" />
  <tr>
    <td align="right" width="166">&nbsp;</td>
    <td colspan="4" align="center">Q4</td>
    <td colspan="3" align="center">QTD</td>
    <td width="70" rowspan="2" align="center">QTG</td>
    <td colspan="8" align="center">Dec</td>
  </tr>
  <tr>
    <td rowspan="2">Pre-sales</td>
    <td width="61" align="center">Oct</td>
    <td width="61" align="center">Nov</td>
    <td width="61" align="center">Dec</td>
    <td width="70" rowspan="2" align="center">Total </td>
    <td width="70" align="center">Target</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Var</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td colspan="3" align="center">Full Month</td>
  </tr>
  <tr>
    <td width="61" align="center"> Target </td>
    <td width="61" align="center"> Target </td>
    <td width="61" align="center"> Target </td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">&nbsp;</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Target</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Target</td>
    <td width="70" align="center">Actual</td>
    <td width="70" align="center">Var</td>
  </tr>
  <tr>
    <td align="left">BG-SDH </td>
    <td align="right">854</td>
    <td align="right">718</td>
    <td align="right">525</td>
    <td align="right">2,098</td>
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
    <td align="left">BG-TH</td>
    <td align="right">1,264</td>
    <td align="right">1,336</td>
    <td align="right">742</td>
    <td align="right">3,342</td>
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
    <td align="left">TH1 - BKM</td>
    <td align="right">560</td>
    <td align="right">631</td>
    <td align="right">387</td>
    <td align="right">1,577</td>
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
    <td align="left">TH2 - Pleno</td>
    <td align="right">704</td>
    <td align="right">705</td>
    <td align="right">355</td>
    <td align="right">1,765</td>
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
    <td align="left">BG-CD1 (AP+JV100%) </td>
    <td align="right">436</td>
    <td align="right">1,237</td>
    <td align="right">552</td>
    <td align="right">2,224</td>
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
    <td align="left">BG-CD2 (AP+JV100%)</td>
    <td align="right">472</td>
    <td align="right">452</td>
    <td align="right">301</td>
    <td align="right">1,225</td>
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
    <td align="left">Total</td>
    <td align="right">3,026</td>
    <td align="right">3,743</td>
    <td align="right">2,120</td>
    <td align="right">8,888</td>
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



