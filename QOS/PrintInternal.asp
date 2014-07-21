<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>Report</title>
<style type="text/css">
<!--
.style1 {
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			
			window.location.href="PrintInternal.asp?id="+val;
		}else{
			
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="PrintInternal.asp?id="+strUser;
		}
		
}
</script>
</head>

<body>
<%
dim getSOP_Code
dim Dateddmmyyyy
Dateddmmyyyy=Now()
getSOP_Code = Request.Form("hidMC")
'response.write getSOP_Code&"sdfsdfsdf"
if isEmpty(getSOP_Code) = false then
%>
<table width="70%" border="2" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td width="85%" align="center"><font style="font-size:24px;"><b>รายการ</b></font></td>
    <td width="15%">&nbsp;</td>
  </tr>
  
  <%
  		dim countloop,mkReport
		countloop=1
	
  		set RecshowQS = Server.CreateObject("ADODB.RECORDSET")
		set RecshowQSR = Server.CreateObject("ADODB.RECORDSET")
  		SQL = "select * from Tb_Internalaudit where M_Code='"&getSOP_Code&"' and   Audit_Year='"&(year(Dateddmmyyyy)+543)&"' order by ID asc ,Audit_DocType ASC   "
		'response.write SQL
		RecshowQS.open SQL,ConQS,1,3
		
		while not RecshowQS.EOF
		getDoctype = RecshowQS("Audit_Doctype")
		getNoCarPer = RecshowQS("No_Car_Par")
		getM_Code = RecshowQS("M_Code")
		getM_Name = RecshowQS("M_Name")
		if countloop = 1 then
  %>
  <tr>
  <td><font style="font-size:22px;"><b><%=getM_Code%>&nbsp;&nbsp;&nbsp;<%=getM_Name%></b></font></td>
  <td>&nbsp;</td>
  </tr>
  <% end if%>
  <tr>
    <td class="style1"><b><% 
	if getDoctype = "NC" then 
	response.write "ใบขอให้ปฎิบัติการแก้ไข CAR No."
	elseif getDoctype = "OBS" then
	response.write "ใบขอให้ปฏิบัติการป้องกัน PAR No."
	elseif getDoctype = "C" then
	response.write "รายงานการตรวจติดตามคุณภาพภายใน Audit No."
	end if
	response.write "&nbsp;&nbsp;"&getNoCarPer
	%></b></td>
    <td align="center">
    <% if getDoctype = "C" then %>
    <input type="button"  value="พิมพ์" style="width:60; height:35; font-size:18px" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportC.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPer%>&MC=<%=getM_Code%>','_blank');}"/>
    <% else %>
    <input type="button"  value="พิมพ์"  style="width:60; height:35; font-size:18px" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportCAR.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPer%>&MC=<%=getM_Code%>','_blank');}"/>
    <% end if %>
    </td>
  </tr>
  <%
  countloop=countloop+1
  RecshowQS.MoveNext
  wend
  %>
</table><br />
<table width="70%"  align="center"><tr><td width="100%"><input type="button"  value="ย้อนกลับ" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/InternalAudit.asp','_self');}"/></td></tr></table>
<% else %>
<%
				
				Response.Clear()
				Response.Redirect("InternalAudit.asp")
				Response.End()
				
				
				if isEmpty(Request.QueryString("id")) = true then
					 if isEmpty(Request.Form("hidDid")) = false then
						getDid=Request.Form("hidDid")
					 else
						getDid = "1"
					 end if
				else
					getDid=Request.QueryString("id")
				end if
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
              <div align="center">
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:14px"    >
			  <%
			  while not recDepart.EOF
			  if recDepart("D_Id") = getDid then
			  selected = "selected=""selected"""
			  else
			  selected = ""
			  end if
			  %>
			  <option value="<%=recDepart("D_Id")%>" <%=selected%> ><%=recDepart("D_Name")%></option>
			  <%
			  recDepart.MoveNext
			  wend
			  recDepart.Close
			  Set recDepart = Nothing
			  %>
			  </select></div><br /> 
              <table width="70%" border="2" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td width="85%" align="center"><font style="font-size:24px;"><b>รายการ</b></font></td>
    <td width="15%">&nbsp;</td>
  </tr>
  <%
		countloop=1
	
  		set RecshowQS = Server.CreateObject("ADODB.RECORDSET")
		set RecshowQSR = Server.CreateObject("ADODB.RECORDSET")
  		SQL = "select * from Tb_Internalaudit where Audit_Depart='"&getDid&"' and   Audit_Year='"&(year(Dateddmmyyyy)+543)&"' order by ID Desc ,Audit_DocType ASC   "
		'response.write SQL
		RecshowQS.open SQL,ConQS,1,3
		
		while not RecshowQS.EOF
		getDoctype = RecshowQS("Audit_Doctype")
		getNoCarPer = RecshowQS("No_Car_Par")
		getM_Code = RecshowQS("M_Code")
  %>
  <tr>
    <td class="style1"><b><% 
	if getDoctype = "NC" then 
	response.write "ใบขอให้ปฎิบัติการแก้ไข"
	elseif getDoctype = "OBS" then
	response.write "ใบขอให้ปฏิบัติการป้องกัน"
	elseif getDoctype = "C" then
	response.write "รายงานการตรวจติดตามคุณภาพภายใน"
	end if
	response.write "&nbsp;&nbsp;"&getNoCarPer&"&nbsp;&nbsp;รหัสเอกสารคุณภาพ&nbsp;"&getM_Code
	%></b></td>
    <td align="center">
    <% if getDoctype = "C" then %>
    <input type="button"  value="พิมพ์" style="width:60; height:35; font-size:18px" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportC.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPer%>&MC=<%=getM_Code%>','_blank');}"/>
    <% else %>
    <input type="button"  value="พิมพ์"  style="width:60; height:35; font-size:18px" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/showReportCAR.asp?adt=<%=getDoctype%>&ncp=<%=getNoCarPer%>&MC=<%=getM_Code%>','_blank');}"/>
    <% end if %>
    </td>
  </tr>
  <%
  countloop=countloop+1
  RecshowQS.MoveNext
  wend
  %>
</table><br />
<table width="70%"  align="center"><tr><td width="100%"><input type="button"  value="ย้อนกลับ" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/InternalAudit.asp','_self');}"/></td></tr></table>
<%
end if
%>
</body>
</html>

