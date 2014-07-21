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
<title>รายงานการวิเคราะห์กระบวนการตาม Core Support Process</title>
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
<body>
<%
dim getAllDepartCore,getAllDepartSupport
set  RecAllDepartCore = Server.CreateObject("ADODB.RECORDSET")
set  RecAllDepartSupport = Server.CreateObject("ADODB.RECORDSET")
Sql_AllDepartCore = "select  count(M_Id) as Core  from  Tb_Manual  where M_Main=1 "
Sql_AllDepartSupport = "select  count(M_Id) as Support  from  Tb_Manual  where M_Reserve=1 "

'----------------------------------------get core----------------------------------------------
 RecAllDepartCore.open Sql_AllDepartCore,ConQS,1,3
 while not RecAllDepartCore.EOF
 getAllDepartCore =  RecAllDepartCore("Core")
 RecAllDepartCore.MoveNext
 wend
 'response.write getAllDepartCore
 '--------------------------------------------------------------------------------------------------
 
 '----------------------------------------get support----------------------------------------------
 RecAllDepartSupport.open Sql_AllDepartSupport,ConQS,1,3
 while not RecAllDepartSupport.EOF
 getAllDepartSupport =  RecAllDepartSupport("Support")
 RecAllDepartSupport.MoveNext
 wend
 'response.write getAllDepartSupport
 '--------------------------------------------------------------------------------------------------
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#666666">
  <tr>
    <td align="center"  bgcolor="#FFFF99"><table width="85%" border="0" cellspacing="0" cellpadding="2">
      <tr >
        <td align="center" valign="baseline">Core</td>
        <td align="center" valign="baseline">Support</td>
      </tr>
      <tr bgcolor="#FFFF99">
        <td  valign="bottom" align="center"><%=getAllDepartCore%><br><img  src="images/core.jpg" height="<%=getAllDepartCore%>" width="40"></td>
        <td valign="bottom" align="center"><%=getAllDepartSupport%><br><img  src="images/support.jpg" height="<%=getAllDepartSupport%>" width="40"></td>
      </tr>
      </table></td>
      <tr>
      <%
	  sql_allDepart = "select * from Tb_Department  order by  D_Numberlist ASC"
	  set recalldepart = Server.CreateObject("ADODB.RECORDSET")
	  recalldepart.open sql_allDepart,ConQS,1,3
	  while not recalldepart.EOF
	  	getIdDepart = recalldepart("D_Id") 
	  %>
      <td colspan="2">
      
      <table width="80%" border="0" cellspacing="0" cellpadding="2">
      <tr >
        <td align="center" valign="baseline">Core</td>
        <td align="center" valign="baseline">Support</td>
      </tr>
      <tr bgcolor="#FFFF99">
        <td  valign="bottom" align="center"><%=getAllDepartCore%><br><img  src="images/core.jpg" height="<%=getAllDepartCore%>" width="20"></td>
        <td valign="bottom" align="center"><%=getAllDepartSupport%><br><img  src="images/support.jpg" height="<%=getAllDepartSupport%>" width="20"></td>
      </tr>
      </table>
      
      </td>
      <%
	  recalldepart.MoveNext
	  wend
	  %>
      </tr>
    </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td >&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<br />

</body>
</html>
