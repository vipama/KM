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
'dim getAllDepartCore,getAllDepartSupport
'set  RecAllDepartCore = Server.CreateObject("ADODB.RECORDSET")
'set  RecAllDepartSupport = Server.CreateObject("ADODB.RECORDSET")
'Sql_AllDepartCore = "select  count(M_Id) as Core  from  Tb_Manual  where M_Main=1 "
'Sql_AllDepartSupport = "select  count(M_Id) as Support  from  Tb_Manual  where M_Reserve=1 "

'----------------------------------------get core----------------------------------------------
'RecAllDepartCore.open Sql_AllDepartCore,ConQS,1,3
' while not RecAllDepartCore.EOF
' getAllDepartCore =  RecAllDepartCore("Core")
' RecAllDepartCore.MoveNext
' wend
' response.write getAllDepartCore
 '--------------------------------------------------------------------------------------------------
 
 '----------------------------------------get support----------------------------------------------
'RecAllDepartSupport.open Sql_AllDepartSupport,ConQS,1,3
' while not RecAllDepartSupport.EOF
' getAllDepartSupport =  RecAllDepartSupport("Support")
' RecAllDepartSupport.MoveNext
' wend
' response.write getAllDepartSupport
 '--------------------------------------------------------------------------------------------------
%>
<!--<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#666666">
  <tr>
    <td width="12%" align="center"  bgcolor="#FFFF99"><table width="90%" border="0" cellspacing="0" cellpadding="2">
      <tr >
        <td align="center" valign="baseline">Core</td>
        <td align="center" valign="baseline">Support</td>
      </tr>
      <tr bgcolor="#FFFF99">
        <td  valign="bottom" align="center"><%'=getAllDepartCore%><br><img  src="images/core.jpg" height="<%'=getAllDepartCore%>" width="40"></td>
        <td valign="bottom" align="center"><%'=getAllDepartSupport%><br><img  src="images/support.jpg" height="<%'=getAllDepartSupport%>" width="40"></td>
      </tr>
      <tr>
      <td colspan="2">คณะกรรมการอาหารและยา</td>
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
-->
<div align="center" class="style2">รายงานการวิเคราะห์กระบวนการตามกระบวนการหลักและกระบวนการสนับสนุน</div><br />
<table width="100%" border="1" cellspacing="0" cellpadding="2">
  <tr bgcolor="#999999">
    <td width="60%" align="center" class="style3">สำนัก / กอง / กลุ่ม</td>
    <td width="20%" align="center" class="style3">กระบวนการหลัก</td>
    <td width="20%" align="center" class="style3">กระบวนการสนับสนุน</td>
  </tr>
  <tr bgcolor="#FFFF99">
<td  align="left" class="style3" >กองผลิตภัณฑ์</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
</tr>
  <%
  dim  countsumcore,countsumsupport
  countsumcore=0
  countsumsupport=0
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='0' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
    <td align="center" class="style3">
    <%
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
	%>
    		<a onClick="javascript:{ window.open('PopupManual.asp?Did=<%=RecDepart("D_Id")%>&Tid=PC','_blank','toolbar=yes, scrollbars=yes, resizable=yes, top=500, left=500, width=650, height=300');}" style="cursor:pointer; cursor:hand; color:#0000FF">
			<% 
	end if
			response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			countsumcore = countsumcore+ GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
			%>
            </a>
    <%
	end if
	%>
    </td>
    <td align="center" class="style3">
    		<% 
			'response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") 
			'countsumsupport = countsumsupport+ GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
			%>
            <%
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
	%>
    		<a onClick="javascript:{ window.open('PopupManual.asp?Did=<%=RecDepart("D_Id")%>&Tid=PS','_blank','toolbar=yes, scrollbars=yes, resizable=yes, top=500, left=500, width=650, height=300');}" style="cursor:pointer; cursor:hand; color:#0000FF">
			<% 
	end if
			response.write GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
			countsumsupport = countsumsupport+ GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
	if GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
			%>
            </a>
    <%
	end if
	%>
    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3" >รวม</td>
<td align="center" class="style3"><%=countsumcore%></td>
<td align="center" class="style3"><%=countsumsupport%></td>
</tr>
 <tr  bgcolor="#FFFF99">
<td  align="left" class="style3">กองสนับสนุน</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
</tr>
 <%
  countsumcore=0
  countsumsupport=0
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='1' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
    <td align="center" class="style3">
    		<% 
			'response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			'countsumcore = countsumcore+ GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			%>
            <%
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
	%>
    		<a onClick="javascript:{ window.open('PopupManual.asp?Did=<%=RecDepart("D_Id")%>&Tid=PC','_blank','toolbar=yes, scrollbars=yes, resizable=yes, top=500, left=500, width=650, height=300');}" style="cursor:pointer; cursor:hand; color:#0000FF">
			<% 
	end if
			response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			countsumcore = countsumcore+ GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
			%>
            </a>
    <%
	end if
	%>
    </td>
    <td align="center" class="style3">
    		<% 
			'response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") 
			'countsumsupport = countsumsupport+ GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
			%>
            <%
	if GetCountRowQS("Tb_Manual","M_Id"," where M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
	%>
    		<a onClick="javascript:{ window.open('PopupManual.asp?Did=<%=RecDepart("D_Id")%>&Tid=PS','_blank','toolbar=yes, scrollbars=yes, resizable=yes, top=500, left=500, width=650, height=300');}" style="cursor:pointer; cursor:hand; color:#0000FF">
			<% 
	end if
			response.write GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
			countsumsupport = countsumsupport+ GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'")
	if GetCountRowQS("Tb_Manual","M_Id"," where  M_Reserve=1 and D_Id='"&RecDepart("D_Id")&"'") > 0 then
			%>
            </a>
    <%
	end if
	%>
    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3">รวม</td>
<td align="center" class="style3"><%=countsumcore%></td>
<td align="center" class="style3"><%=countsumsupport%></td>
</tr>
</table>
</body>
</html>
