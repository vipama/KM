<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim chkPS,chkPC,chkQ,chkW,FlagMain_Reserve
 chkPS=""
 chkPC=""
 chkQ=""
 chkW=""
 FlagMain_Reserve=""
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = "0"
	 end if
else
	getDid=Request.QueryString("id")
end if

if isEmpty(Request.QueryString("oid")) = true then
	 if isEmpty(Request.Form("hidOid")) = false then
	 	getOid=Request.Form("hidOid")
	 else
	 	getOid = "0"
	 end if
else
	getOid=Request.QueryString("oid")
end if

if isEmpty(Request.QueryString("tid")) <> true then
	if Request.QueryString("tid") = "0" then
		getTid = "0"
		FlagMain_Reserve = "M_Main=1"
		chkPC = "checked=""checked"""
	elseif Request.QueryString("tid") = "1" then
		getTid = "1"
		FlagMain_Reserve = "M_Reserve=1"
		chkPS= "checked=""checked"""
	elseif Request.QueryString("tid") = "2" then
		getTid = "2"
		FlagMain_Reserve = "M_Reserve=1"
		chkPS= "checked=""checked"""
	end if
else
	getTid = "0"
	FlagMain_Reserve = "M_Main=1"
	chkPC = "checked=""checked"""
end if
'response.write getTid&"<br>"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>รายงานวิเคราะห์การตรวจติดตามคุณภาพภายใน</title>
<style type="text/css">
<!--
.style1 {
font-size:10px;
font-family:Arial, Helvetica, sans-serif;
}
-->
</style>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		
		
		if ((val != "" ) || (val1 != ""))
		{ 
			if ((val == "" ) && (val1 != ""))
			{
					var typeID = document.getElementById("TypeDoc").value;
					//alert("1/"+typeID);
					//var e = document.getElementById("DepartID");    
					//var strUser = e.options[e.selectedIndex].value;
					var typ = document.getElementById("TypeDoc").value;
					window.location.href="AnalaysisInternalAuditReport.asp?id=0&oid="+typeID+"&tid="+typeID;
					//window.location.href="ReviewReport.asp?id="+val+"&oid="+val1+"&tid="+typeID;
			  }else if((val != "" ) && (val1 == "")){
			  
			  		var typeID = document.getElementById("TypeDoc").value;
					//alert("2/"+typeID);
					var e = document.getElementById("DepartID");    
					var strUser = e.options[e.selectedIndex].value;
					var typ = document.getElementById("TypeDoc").value;
					
					window.location.href="AnalaysisInternalAuditReport.asp?id="+strUser+"&oid="+typeID+"&tid="+typeID;
			  }
		}else{
			
			var typeID = document.getElementById("TypeDoc").value;
			//alert("2/"+typeID);
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			var typ = document.getElementById("TypeDoc").value;
			
			window.location.href="AnalaysisInternalAuditReport.asp?id="+strUser+"&oid="+typeID+"&tid="+typeID;
		}
		
}
</script>
</head>

<body>
<div class="style1" align="center" style="font-size:18px">รายงานวิเคราะห์การตรวจติดตามคุณภาพภายใน</div><br />
<div class="style1" align="center" style="font-size:18px">
	  <select name="TypeDoc" id="TypeDoc" onChange="ChangeJobresultGroup('',this.value)" style="font-size:16px"  >
      <option value="0" <% if getOid ="0" then response.write " selected=""selected"" " end if%> >เลือกระดับหน่วยงาน</option>
	  <option value="1" <% if getOid ="1" then response.write " selected=""selected"" " end if%> >ระดับกรม</option>
      <option value="2"  <% if getOid ="2" then response.write " selected=""selected"" " end if%>>ระดับหน่วยงาน</option>
    
      </select>
</div>
<% if getOid = 2  then %>
<br />
<div class="style1" align="center" style="font-size:18px">
<%
	  Set   rec_jobresult_group = Server.CreateObject("ADODB.RECORDSET")
	  sql_jobresult_group = "select  *  from  Tb_Department order by D_Numberlist  asc"
	  rec_jobresult_group.open sql_jobresult_group,ConQS,1,3
	  %>
	  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroup(this.value,'')" style="font-size:16px"  >
      <option value="0"  <% if getDid = "0" then  response.write "seledted=""selected"" " end if %>>เลือกหน่วยงาน</option>
	  <%
	  while not rec_jobresult_group.EOF
	  if rec_jobresult_group("D_Id") = getDid then
	  selected = "selected=""selected"""
	  else
	  selected = ""
	  end if
	  %>
	  <option value="<%=rec_jobresult_group("D_Id")%>" <%=selected%> ><%=rec_jobresult_group("D_Name")%></option>
	  <%
	  rec_jobresult_group.MoveNext
	  wend
	  rec_jobresult_group.Close
	  Set rec_jobresult_group = Nothing
	  %>
      </select>
</div>
<% end if %>
<br />

<div>
<%
'if getTid ="PC" or getTid = "PS" then
'countsumcore = GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&getDid&"'")
'response.write " กระบวนการทั้งหมด  :"&countsumcore&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'"))&" กระบวนการ"
'elseif  getTid = "W" then
'countsumcore = GetCountRowQS("Tb_Workin","D_Id"," where D_Id='"&getDid&" ' ")
'response.write " กระบวนการทั้งหมด : "&countsumcore&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&" '  and Type_Sop='W' ")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'   and Type_Sop='W' "))&" กระบวนการ"
'elseif  getTid = "Q" then
'countsumcore3 = GetCountRowQS("Tb_QM","D_Id"," where D_Id='"&getDid&"' ")
'response.write " กระบวนการทั้งหมด : "&countsumcore3&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'  and Type_Sop='Q' ")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore3-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'  and Type_Sop='Q' "))&" กระบวนการ"
'end if
%>
</div>
<table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#666666">
  <tr>
    <td width="5%" rowspan="3"><div align="center">วันที่ตรวจติดตาม</div></td>
    <td width="20%" rowspan="3"><div align="center">ชื่อหน่วยงานที่<br />
      รับการตรวจ</div></td>
    <td width="5%" rowspan="3"><div align="center">รหัส</div></td>
    <td width="30%" rowspan="3"><div align="center">ชื่อเอกสาร</div></td>
    <td width="30%" colspan="3"><div align="center">ผลการตรวจติดตามคุณภาพภายใน</div></td>
  </tr>
  <tr>
    <td width="5%" rowspan="2"><div align="center">ไม่พบ<br />
      ข้อบกพร่อง</div></td>
    <td colspan="2" width="20%"><div align="center">พบข้อบกพร่อง</div></td>
  </tr>
  <tr>
    <td width="10%" align="center"><div align="center">CAR No.</div></td>
    <td width="10%" align="center"><div align="center">PAR No.</div></td>
  </tr>
 
  <%
  if getTid = "1" then
  	 ' response.write "case 1 <br>"
 	 sqlDepart1 = "select  distinct Audit_Depart  from  Tb_InternalAudit  where Audit_Level='1' order by Audit_Depart asc "
	 set recDepart = Server.CreateObject("ADODB.RECORDSET")
	 set recDoc = Server.CreateObject("ADODB.RECORDSET")
	 recDepart.Open sqlDepart1,ConQS,1,3
	 While not recDepart.EOF
	 response.write recDepart("Audit_Depart")&"<br>"
	 
	 	sqlDoc = "select * from Tb_InternalAudit where Audit_Level='1' "
		recDoc.open sqlDoc,ConQS,1,3
		While NOT recDoc.EOF
		
	%>
	<tr>
    <td align="center"><%=recDoc("Audit_Date")%></td>
    <td align="center"><%=getDepartmentname(recDoc("Audit_Depart"))%></td>
    <td align="center"><%=recDoc("M_Code")%></td>
    <td align="left"><%=recDoc("M_Name")%></td>
    <td align="center">
	<%
	if recDoc("Audit_Flag_Complete") = 0 then
		response.write "&#149;"
	else
		response.write "&nbsp;"
	end if
	%>
    </td>
    <td align="center">
	<%
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> 0 then
	 	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 else
	 	response.write "&nbsp;"
	 end if
	 %>
     </td>
    <td align="center">
	<%
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> 0 then
	 	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 else
	 	response.write "&nbsp;"
	 end if
	 %>
     </td>
  </tr>
	<%	
		
		recDoc.MoveNext
		Wend
		recDoc.Close()
	  
	 recDepart.MoveNext
	 Wend 
	 if recDepart.RecordCount = 0 then
 %>
 <tr><td colspan="7" align="center">No Data</td></tr>
 <%
 	end if
	recDepart.Close()
	 
  else
  	 'Response.write "Else case <br>"
  	 sqlDepart2 = "select  Distinct Audit_Depart  from  Tb_InternalAudit where Audit_Depart='"&getDid&"' and Audit_Level='2' order by Audit_Depart asc  "
	 set recDepart = Server.CreateObject("ADODB.RECORDSET")
	 set recDoc = Server.CreateObject("ADODB.RECORDSET")
	 recDepart.Open sqlDepart2,ConQS,1,3
	 While not recDepart.EOF
	 'response.write sqlDepart2&"<br>"
	 
	 	sqlDoc = "select * from Tb_InternalAudit where Audit_Level='2' and Audit_Depart='"&getDid&"' "
		recDoc.open sqlDoc,ConQS,1,3
		While NOT recDoc.EOF
			'response.write recDepart("Audit_Depart")&" * "&recDoc("M_Code")&"<br>"
	%>
	<tr>
    <td align="center"><%=recDoc("Audit_Date")%></td>
    <td align="center"><%=getDepartmentname(recDoc("Audit_Depart"))%></td>
    <td align="center"><%=recDoc("M_Code")%></td>
    <td align="left"><%=recDoc("M_Name")%></td>
    <td align="center">
	<%
	if recDoc("Audit_Flag_Complete") = 0 then
		response.write "&#149;"
	else
		response.write "&nbsp;"
	end if
	%>
    </td>
    <td align="center">
	<%
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> 0 then
	 	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 else
	 	response.write "&nbsp;"
	 end if
	 %>
     </td>
    <td align="center">
	<%
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> 0 then
	 	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 else
	 	response.write "&nbsp;"
	 end if
	 %>
     </td>
  </tr>
	<%
		recDoc.MoveNext
		Wend
		recDoc.Close()
	  
	 recDepart.MoveNext
	 Wend
	if recDepart.RecordCount = 0 then
%>
  <tr><td colspan="7" align="center">No Data</td></tr>
  <% 
  	 end if
	 recDepart.Close()
  end if
 %>
</table>
</body>
</html>
