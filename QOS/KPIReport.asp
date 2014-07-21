<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim Dateddmmyyyy
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
Datemmddyyyy1=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
if isEmpty(session("member")) = True then
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>KPI Report</title>
<style type="text/css">
<!--
.style1 {
font-size:14px;
font-family:Arial, Helvetica, sans-serif;
}
.style2 {
font-size:22px;
font-family:Arial, Helvetica, sans-serif;
font-weight:600;
}
.style3 {
font-size:16px;
font-family:Arial, Helvetica, sans-serif;
font-weight:100;
}
-->
</style>
</head>

<body bgcolor="#000000">
<table width="80%" border="0" cellspacing="0" cellpadding="1" align="center" ><tr><td class="style3" align="center"><font style="font-size:24px; color:#FFFFFF">รายงานการวิเคราะห์ระดับความสำเร็จของการพัฒนาระบบคุณภาพ </font></td></tr></table><br />
<table width="80%" border="1" cellspacing="0" cellpadding="1" align="center" bgcolor="#666666">
  <tr bgcolor="#CCCCCC">
    <td width="80%" rowspan="2" class="style3" ><div align="center">สำนัก / กอง / กลุ่ม</div></td>
    <td width="20%" colspan="5" class="style3" ><div align="center">ระดับความสำเร็จ</div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td align="center" class="style3">1</td>
    <td align="center" class="style3">2</td>
    <td align="center" class="style3">3</td>
    <td align="center" class="style3">4</td>
    <td align="center" class="style3">5</td>
  </tr>
  <%
  dim chkLevel,flagset,flagColor
  set recallDepart =  Server.CreateObject("ADODB.RECORDSET")
  sql_alldepart = "select * from Tb_Department order by  D_Numberlist"
  recallDepart.open sql_allDepart,ConQS,1,3
  While not recallDepart.EOF
  getDepartName = recallDepart("D_Name")
  %>
  <tr>
    <td class="style3" bgcolor="#FFFF99">&nbsp;&nbsp;<%=getDepartName%></td>
    <%
	chkLevel = CheckLevelSuccess1(recallDepart("D_Id"))
	if chkLevel > 0 then
		flagColor="bgcolor=""#00FF00"" "
		flagset = "&#149;"
	else
		flagColor="bgcolor=""#FFFF99"""
		flagset="&nbsp;"
	end if
	%>
    <td align="center" <%=flagColor%>><%=flagset%></td>
     <%
	chkLevel = CheckLevelSuccess2(recallDepart("D_Id"))
	if chkLevel > 0 then
		flagColor="bgcolor=""#00FF00"" "
		flagset = "&#149;"
	else
		flagColor="bgcolor=""#FFFF99"""
		flagset="&nbsp;"
	end if
	%>
    <td align="center" <%=flagColor%>><%=flagset%></td>
	<%
	'-------------------------------------------------------------------Block check Level 3-------------------------------------------------
	chkLevel = CheckLevelSuccess2(recallDepart("D_Id"))
	if chkLevel > 0 then
		chkLevel3 = CheckLevelSuccess3(recallDepart("D_Id")) 
		if chkLevel3 > 0 then
			flagColor="bgcolor=""#00FF00"" "
			flagset = "&#149;"
		else
			flagColor="bgcolor=""#FFFF99"""
			flagset="&nbsp;"
		end if	
	else
		flagColor="bgcolor=""#FFFF99"""
		flagset="&nbsp;"
	end if
	'------------------------------------------------------------------------------------------------------------------------------------------
	%>
    <td align="center" <%=flagColor%>><%=flagset%></td>
     <%
	'-------------------------------------------------------------------Block check Level 4-------------------------------------------------
	chkLevel = CheckLevelSuccess2(recallDepart("D_Id"))
	if chkLevel > 0 then
		chkLevel3 = CheckLevelSuccess3(recallDepart("D_Id")) 
		if chkLevel3 > 0 then
			chkLevel4 = CheckLevelSuccess4(recallDepart("D_Id")) 
			if chkLevel4 > 0 then
				flagColor="bgcolor=""#00FF00"" "
				flagset = "&#149;"
			else
				flagColor="bgcolor=""#FFFF99"""
				flagset="&nbsp;"
			end if
		else
			flagColor="bgcolor=""#FFFF99"""
			flagset="&nbsp;"
		end if	
	else
		flagColor="bgcolor=""#FFFF99"""
		flagset="&nbsp;"
	end if
	'------------------------------------------------------------------------------------------------------------------------------------------
	%>
    <td align="center" <%=flagColor%>><%=flagset%></td>
     <%
	'-------------------------------------------------------------------Block check Level 5-------------------------------------------------
	chkLevel = CheckLevelSuccess2(recallDepart("D_Id"))
	if chkLevel > 0 then
		chkLevel3 = CheckLevelSuccess3(recallDepart("D_Id")) 
		if chkLevel3 > 0 then
			chkLevel4 = CheckLevelSuccess4(recallDepart("D_Id")) 
			if chkLevel4 > 0 then
				flagColor="bgcolor=""#00FF00"" "
				flagset = "&#149;"
			else
				flagColor="bgcolor=""#FFFF99"""
				flagset="&nbsp;"
			end if
		else
			flagColor="bgcolor=""#FFFF99"""
			flagset="&nbsp;"
		end if	
	else
		flagColor="bgcolor=""#FFFF99"""
		flagset="&nbsp;"
	end if
	'------------------------------------------------------------------------------------------------------------------------------------------
	%>
    <td align="center" <%=flagColor%>><%=flagset%></td>
  </tr>
  <%
  recallDepart.MoveNext
  Wend
  %>
</table>
</body>
</html>
