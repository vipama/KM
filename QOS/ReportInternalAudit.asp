<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim Dateddmmyyyy
Dateddmmyyyy=Now()
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>InternalAudit Report</title>
<style type="text/css">
<!--
.style1 {
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
.style2 {
font-size:22px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
</head>

<body>
              <%
				
				if isEmpty(Request.QueryString("did")) = true then
					 if isEmpty(Request.Form("hidDid")) = false then
						getDid=Request.Form("hidDid")
					 else
						response.write "<script  language=""javascript""> "
						response.write "alert('did not available!  \n Please try again!');"
						response.write "window.location.href=""default.asp"";"
						response.write "</script > "
					 end if
				else
					getDid=Request.QueryString("did")
				end if
				if isEmpty(Request.QueryString("AN") ) = true then
					sqlAN = ""
				else
					getAN=Request.QueryString("AN")
					sqlAN = " and  No_Car_par='"&getAN&"' "
				end if
			 %>
<br /> 
              <table width="100%" border="2" align="center" cellpadding="3" cellspacing="0">
              <tr bgcolor="#999999"><td colspan="5" class="style2" height="35"><b><%=getDepartmentname(getDid)%></b></td></tr>
  <tr bgcolor="#CCCCCC">
    <td width="25%" align="center" class="style1" height="35"><b>Type of Report</b></td>
    <td width="15%" align="center" class="style1"><b>Audit No.</b></td>
    <td width="15%" align="center" class="style1"><b>SOP No.</b></td>
    <td width="40%" align="center" class="style1"><b>SOP Name.</b></td>
    <td width="10%" align="center" class="style1"><b>&nbsp;</b></td>
  </tr>
  <%
		countloop=1
	
  		set RecshowQS = Server.CreateObject("ADODB.RECORDSET")
		set RecshowQSR = Server.CreateObject("ADODB.RECORDSET")
  		SQL = "select * from Tb_Internalaudit where Audit_Depart='"&getDid&"' and  Audit_Year='"&(year(Dateddmmyyyy)+543)&"' "&sqlAN&" order by ID Desc ,Audit_DocType ASC   "
		'response.write SQL
		RecshowQS.open SQL,ConQS,1,3
		
		while not RecshowQS.EOF
		getDoctype = RecshowQS("Audit_Doctype")
		getNoCarPar = RecshowQS("No_Car_Par")
		getM_Code = RecshowQS("M_Code")
		getM_Name = RecshowQS("M_Name")
  %>
  <tr>
    <td class="style1">
	<%
	if getDoctype = "NC" then 
	response.write "ใบขอให้ปฎิบัติการแก้ไข "
	elseif getDoctype = "OBS" then
	response.write "ใบขอให้ปฏิบัติการป้องกัน "
	elseif getDoctype = "C" then
	response.write "รายงานการตรวจติดตามคุณภาพภายใน "
	end if
	 %>
    </td>
     <td class="style1"><%=getNoCarPar%></td>
     <td class="style1"><%=getM_Code%></td>
    <td class="style1">
    <% 
	response.write getM_Name
	%></td>
    <td align="center"><% if getDoctype = "C" then %>
      <input type="button"  value="พิมพ์" style="width:60; height:22; font-size:14px; vertical-align:middle" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportC.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPar%>&MC=<%=getM_Code%>','_blank');}"/>
      <% else %>
      <input type="button"  value="พิมพ์"  style="width:60; height:22; font-size:14px; vertical-align:middle"  onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportCAR.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPar%>&MC=<%=getM_Code%>','_blank');}"/>
      <% end if %>
    </td>
  </tr>
  <%
  countloop=countloop+1
  RecshowQS.MoveNext
  wend
  %>
              </table><br />
<table width="100%"  align="center">
<tr><td width="100%"><input type="button"  value="ย้อนกลับ" style="width:80; height:35; font-size:18px" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/InternalAudit.asp','_self');}"/>&nbsp;&nbsp;&nbsp;<input type="button"  value="พิมพ์หน้านี้" onclick="javascript:{window.print();}" style="width:90; height:35; font-size:18px" /></td></tr></table>

</body>
</html>
