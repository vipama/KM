<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
'===========================Login Member=====================================
Email=BlockSqlConjection(request("Email"))
password=BlockSqlConjection(request("password"))
login=request("login")
chk_login = Request("chk_login")
qs_chk = Request.Form("qs_chk")
if login="1"or chk_login="1"  then
		set rs_login=server.createobject("ADODB.recordset")
		rs_login.open "Select * from member Where Email='"&Email&"' And KMPassword='"&password&"' and ConfirmStatus=1",con,1,1
		if rs_login.recordcount<>0 then
				session("member")=rs_login("Email")
				session("act") = rs_login("ACT")
				if  Patchdate2(rs_login("LastLogin"),0,"EN")<>Patchdate2(Date,0,"EN") Then call AddScoreBoard(session("member"),"logincount")
					con.execute("Update Member Set LastLogin=now() Where Email='"&session("member")&"'")
					con.execute("INSERT INTO TabMemberLog (memberid,[date]) VALUES ("&rs_login("id")&",'"&FormatDate(Patchdate2(date,0,"EN"),2)&"')")
					message=""
				else
					message="รหัสผ่านไม่ถูกต้อง"
				end if		
				
				if qs_chk = "QS"	 and rs_login.recordcount<>0 then
					'Response.Redirect "http://filing.fda.moph.go.th/kmfda/default.asp?page=doc"	
					Response.Redirect "http://filing.fda.moph.go.th/kmfda/_block/qos/"
				else
					Response.Redirect "http://filing.fda.moph.go.th/kmfda/_block/qos/"
				end if
		closerecord(rs_login)
elseif login="0" or chk_login="0" then
session("login")=0
session.Contents.Remove("member")
Response.Redirect "default.asp"
End if 
'==============================end of login====================================
%>
<script language="JavaScript" src="../../_java/main.js"></script>
<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">

<body leftmargin="0" topmargin="0" bgcolor="#fffe8c"> 
<center>
    <table width="780" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
              <td><img src="images/qs_fda.gif" width="780" height="150"></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%
						if Isempty(session("member")) = True  then  	
				%>
				<tr><td  bgcolor="#3ed0fd" >&nbsp;&nbsp;&nbsp;<font size="2" face="Ms Sans Serif" color="#3300ff"><strong>Login</strong></font>&nbsp;&nbsp;&nbsp;&nbsp;<font size="2" face="Ms Sans Serif" color="#3300ff"><strong>ผู้ที่สนใจต้องการเข้าดูเอกสารระบบคุณภาพของ อย. สามารถล๊อกอินได้ที่นี่</strong></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="1">
				  <input type="hidden" name="qs_chk" value="QS">
                  <tr>
                    <td bgcolor="#3ed0fd" ><font size="2" face="Ms Sans Serif" color="#3300ff">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E-Mail&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input name="Email" type="text" size=20 value="yourname@fda.moph.go.th" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff">&nbsp;รหัสผ่าน :&nbsp;&nbsp;&nbsp;<input name="Password" type="password" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff"  size=20></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit"  value="เข้าสู่ระบบ">&nbsp;&nbsp;&nbsp;<input type="button" value="สมัครสมาชิก"onclick="javascript:{window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos/register.asp';}" style="cursor:pointer; cursor:hand" ></td>
                  </tr>
                  </form>
				  <tr><td>&nbsp;</td></tr>
				  <% else%>
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="0">
				  <input type="hidden" name="qs_chk" value="QS">
                  <tr>
                    <td bgcolor="#3ed0fd" align="right" ><font size="2" face="Ms Sans Serif" color="#3300ff"><input type="submit"  value="ออกจากระบบ">&nbsp;&nbsp;&nbsp;<input type="button" value="ระบบคุณภาพของ อย."onclick="javascript:{window.location.href='http://filing.fda.moph.go.th/kmfda/default.asp?page=doc';}" style="cursor:pointer; cursor:hand" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                  </form>
				  <tr><td>&nbsp;</td></tr>
				  <% end if %>
                  <tr>
                    <td><br /><!--#include file="checksession.asp"--></td>
                  </tr>
                </table></td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <font size="2" face="Ms Sans Serif" color="#3300ff">"เฉพาะข้าราชการ อย. ที่ยังไม่มีรหัสผ่านในการเข้าดูเอกสารคุณภาพขอให้แจ้งชื่อ นามสกุล มาที่่ library@fda.moph.go.th  และท่านจะได้รับการแจ้งรหัสผ่านกลับทางอีเมล์ที่แจ้งไว้"</font>
  <!--<font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*** ผู้ที่สนใจต้องการเข้าดูเอกสารระบบคุณภาพของ อย. สามารถเข้าไปได้ที่เว็บไซต์ <a href="http://elib.fda.moph.go.th/kmfda/default.asp" target="_blank">FDA KM</a> หัวข้อ "ระบบจัดการเอกสารและข้อมูล" ซึ่งจำกัดให้เฉพาะบุคลากรที่มี e-mail address ของ อย. คือ @fda.moph.go.th เท่านั้น ที่สามารถเข้าใช้งานได้ โดยจะต้องสมัครสมาชิกเพื่อขอรับรหัสผ่าน (Password) .ในการเข้าใช้งานระบบนะคะ  ...*** --></font></td>
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
</center>
</body>
