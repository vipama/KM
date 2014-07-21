<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<script language="JavaScript" src="../../_java/main.js"></script>
<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">

<body leftmargin="0" topmargin="0"> 
<center>
    <table width="780" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
              <td><img src="images/qs_fda.gif" width="780" height="150"></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td><!--#include file="checksession.asp"--></td>
                  </tr>
                </table></td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr>
          <td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*** ผู้ที่สนใจต้องการเข้าดูเอกสารระบบคุณภาพของ อย. สามารถเข้าไปได้ที่เว็บไซต์ <a href="http://elib.fda.moph.go.th/kmfda/default.asp" target="_blank">FDA KM</a> หัวข้อ &quot;ระบบจัดการเอกสารและข้อมูล&quot; ซึ่งจำกัดให้เฉพาะบุคลากรที่มี e-mail address ของ อย. คือ <span style="font-weight: bold">@fda.moph.go.th</span> เท่านั้น ที่สามารถเข้าใช้งานได้ โดยจะต้องสมัครสมาชิกเพื่อขอรับรหัสผ่าน (Password) .ในการเข้าใช้งานระบบนะคะ ...*** </font></td>
        </tr></table>
  </td></tr>
        <tr>
              <td background="images/qs_fda2.gif" height="50" valign="bottom">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
<%

set rs_counter=server.CreateObject("ADODB.recordset")
rs_counter.open "Select * From tabUniIP ",con,1,3


if session("check")="" then
session("check")="x"
rs_counter.addnew
rs_counter("ip")=request.ServerVariables("REMOTE_HOST")
rs_counter("date")=date
rs_counter.update
end if
rs_counter.close

%>
                      <a href="javascript:openWin('counter.asp','counter',500,500)" >Uni IP <img src="../../_images/folder_stats.gif" width="30" height="30" border=0></a></td>
  </tr>
</table>

			  </td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
