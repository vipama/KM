<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
' # start code for check permission in DB 
if Session("member") <> getPermission(session("member"),"L_Email") or isnull(session("member")) = true or session("member") = "" then
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
else
	session("Depart") = getPermission(session("member"),"D_Id")
end if
' # End code for check permission in DB
dim chkPS,chkPC,chkQ,chkW,FlagMain_Reserve
 chkPS=""
 chkPC=""
 chkQ=""
 chkW=""
 FlagMain_Reserve=""
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = "0"
	 end if
else
	getDid=Request.QueryString("id")
	if getDid <> session("Depart") and getDid <> 0 and session("Depart") <> "100" then
	   getDid = session("Depart")
	   response.write "Please not change parameter"
	end if
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
<div class="style1" align="center" style="font-size:18px"><span class="style1" style="font-size:18px">ทะเบียนควบคุมสถานะการตรวจติดตามคุณภาพภายในและการปฏิบัติการแก้ไข/ป้องกัน (CAR/PAR - LOG)</span></div>
<br />
<div class="style1" align="center" style="font-size:18px"><span class="style1" style="font-size:18px">ระดับ</span>
<select name="TypeDoc" id="TypeDoc" onChange="ChangeJobresultGroup('',this.value)" style="font-size:16px"  >
      <option value="0" <% if getOid ="0" then response.write " selected=""selected"" " end if%> >เลือกระดับหน่วยงาน</option>
	  <option value="1" <% if getOid ="1" then response.write " selected=""selected"" " end if%> >ระดับกรม</option>
      <option value="2"  <% if getOid ="2" then response.write " selected=""selected"" " end if%>>ระดับหน่วยงาน</option>
    
      </select>
</div>
<% if getOid = 2  then %>
<br />
<div class="style1" align="center" style="font-size:18px"><span class="style1" style="font-size:18px">ชื่อหน่วยงาน</span>
<%
	  Set   rec_jobresult_group = Server.CreateObject("ADODB.RECORDSET")
	   if session("Depart") = "100" then
	  		sql_jobresult_group = "select  *  from  Tb_DepartmentPermission where D_Id not in('17','18') order by D_Numberlist  asc"
	  else
	  		sql_jobresult_group = "select  *  from  Tb_DepartmentPermission where D_Id='"&session("Depart")&"' order by D_Numberlist  asc"
	  end if
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
    <td width="7%" rowspan="4"><div align="center">วันที่ตรวจติดตาม</div></td>
    <td width="12%" rowspan="4"><div align="center">ชื่อหน่วยงานที่<br />
      รับการตรวจ</div></td>
    <td width="7%" rowspan="4"><div align="center">รหัสเอกสาร</div></td>
    
    <td width="25%" rowspan="4"><div align="center">ชื่อเอกสาร</div></td>
    <td width="7%" rowspan="4"><div align="center">รหัสการตรวจติดตาม</div></td>
    <td width="20%" colspan="5"><div align="center">ผลการตรวจติดตามคุณภาพภายใน</div></td>
    <td width="7%" rowspan="4"><div align="center">การตรวจติดตามซ้ำ</div></td>
  </tr>
  <tr>
    <td width="7%" rowspan="3"><div align="center">ไม่พบ<br />
      ข้อบกพร่อง</div></td>
    
  </tr>
  <tr>
    <td width="7%" colspan="2" align="center"><div align="center">CAR </div></td>
    <td width="7%" colspan="2" align="center"><div align="center">PAR</div></td>
  </tr>
  <tr>
    <td width="7%" align="center" >No.</td>
    <td width="7%" align="center">กำหนดเสร็จ</td>
    <td width="7%" align="center">No.</td>
    <td width="7%" align="center">กำหนดเสร็จ</td>
  </tr>
 
  <%
  if getTid = "1" then
  	 ' response.write "case 1 <br>"
 	 'sqlDepart1 = "select  distinct Audit_Depart as   from  Tb_InternalAudit  where Audit_Level='1' order by Audit_Depart asc "
	 sqlDepart1 = "select  Distinct Audit_Depart as AuDepart , M_Code as MCode  from  Tb_InternalAudit where  Audit_Level='1' order by Audit_Depart asc  "
	 set recDepart1 = Server.CreateObject("ADODB.RECORDSET")
	 set recDoc = Server.CreateObject("ADODB.RECORDSET")
	 recDepart1.Open sqlDepart1,ConQS,1,3
	 While not recDepart1.EOF
	 'response.write recDepart1("AuDepart")&"/"&sqlDepart1&"<br>"
	 	if session("Depart") = "100" then	 
	 		sqlDoc = "select * from Tb_InternalAudit where Audit_Level='1' and M_Code='"&recDepart1("MCode")&"' and Audit_Depart='"&recDepart1("AuDepart")&"' and Audit_DocType='C' order by No_Car_Par ASC  "
		else
			sqlDoc = "select * from Tb_InternalAudit where Audit_Level='1' and M_Code='"&recDepart1("MCode")&"' and Audit_Depart='"&session("Depart")&"' and Audit_DocType='C' order by No_Car_Par ASC  "
		end if
		recDoc.open sqlDoc,ConQS,1,3
		While NOT recDoc.EOF
		
	%>
	<tr>
    <td align="center"><%=recDoc("Audit_Date")%></td>
    <td align="center"><%=getDepartmentname(recDoc("Audit_Depart"))%></td>
    <td align="center"><%=recDoc("M_Code")%></td>
    <td align="left"><%=recDoc("M_Name")%></td>
    <td align="center"><%=recDoc("No_Car_Par")%></td>
    <td align="center">
	<%
	if recDoc("Audit_Flag_Complete") = 0 then
		response.write "&#149;"
	else
		response.write "&nbsp;"
	end if
	%>    </td>
    <td align="center">
	<%
	' if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 '	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	 '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDate) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">"&a(i)&"<br>"&"</font>" 
			end if
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	 %>     </td>
    <td align="center" valign="top"><%
	' if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 '	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	 '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDate) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">&nbsp;<br /></font>" 
			end if
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	 %></td>
    <td align="center">
	<%
	' if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 '	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	 '------------------------------------------------Start block alert PAR No-------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	  
	       getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
		b=Split(getPAR,"<br>")
		for i=0 to Ubound(b)
			getFinishDatePar = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDatePar) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">"&b(i)&"<br>"&"</font>" 
			end if
		next
	 	'response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '-------------------------------------------------End block alert PAR No--------------------------------------------------
	 %>     </td>
    <td align="center">
    <%
	' if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 '	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	 '------------------------------------------------Start block alert PAR No-------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	  
	       getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
		b=Split(getPAR,"<br>")
		for i=0 to Ubound(b)
			getFinishDatePar = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDatePar) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">&nbsp;<br /></font>" 
			end if
		next
	 	'response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '-------------------------------------------------End block alert PAR No--------------------------------------------------
	 %>
    </td>
    <td align="center">
    <%
	 '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if getFinishDate <> "" then
			    '    if cDate(Datemmddyyyy) < cDate(Month(getFinishDate)&"/"&day(getFinishDate)&"/"&Year(getFinishDate)) then
							FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
							if IsDate(getAuditDate2) <> False and IsDate(getAuditDate3) <> False then
								response.write "<font color=""#009900"">"
								response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&""">"&a(i)&"<br>"&"</span>"
								response.write "</font>"
							else
								response.write "&nbsp;"
							end if
				'	else
				'			response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&"""><font color=""#EC0076"">"&a(i)&"<br>"&"</font></span>" 
				'	end if
			else
					response.write "&nbsp;" 
			end if
			
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	%>
    <%
	 '---------------------------------------------Start block PAR No---------------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1") <> "" then
	 	getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 	
		a=Split(getPAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if getFinishDate <> "" then
			   '     if cDate(Datemmddyyyy) < cDate(Month(getFinishDate)&"/"&day(getFinishDate)&"/"&Year(getFinishDate)) then
							FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
							if IsDate(getAuditDate2) <> False and IsDate(getAuditDate3) <> False then
								response.write "<font color=""#009900"">"
								response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&""">"&a(i)&"<br>"&"</span>"
								response.write "</font>"
							else
								response.write "&nbsp;"
							end if
				'	else
				'			response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&"""><font color=""#EC0076"">"&a(i)&"<br>"&"</font></span>" 
				'	end if
			else
					response.write "&nbsp;" 
			end if
			
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write ""
	 end if
	 '----------------------------------------------end block alert PAR No---------------------------------------------------
	%>
    </td>
  </tr>
	<%	
		
		recDoc.MoveNext
		Wend
		recDoc.Close()
	  
	 recDepart1.MoveNext
	 Wend 
	 if recDepart1.RecordCount = 0 then
 %>
 <tr><td colspan="11" align="center">No Data</td></tr>
 <%
 	end if
	recDepart1.Close()
	 
  else
  	 'Response.write "Else case <br>"
  	 sqlDepart2 = "select  Distinct Audit_Depart as AuDepart , M_Code as MCode  from  Tb_InternalAudit where Audit_Depart='"&getDid&"' and Audit_Level='2' order by Audit_Depart asc   "
	 set recDepart2 = Server.CreateObject("ADODB.RECORDSET")
	 set recDoc = Server.CreateObject("ADODB.RECORDSET")
	 recDepart2.Open sqlDepart2,ConQS,1,3
	 While not recDepart2.EOF
	' response.write sqlDepart2&"<br>"
	' response.write recDepart2("MCode")&"<br>"
	 	sqlDoc = "select * from Tb_InternalAudit where Audit_Level='2' and Audit_Depart='"&recDepart2("AuDepart")&"' and M_Code='"&recDepart2("MCode")&"' and Audit_DocType='C' "
		'response.write sqlDoc&"<br>"
		recDoc.open sqlDoc,ConQS,1,3
		While NOT recDoc.EOF
			'response.write recDepart2("Audit_Depart")&" * "&recDoc("M_Code")&"<br>"
	%>
	<tr>
    <td align="center"><%=recDoc("Audit_Date")%></td>
    <td align="center"><%=getDepartmentname(recDoc("Audit_Depart"))%></td>
    <td align="center"><%=recDoc("M_Code")%></td>
    <td align="left"><%=recDoc("M_Name")%></td>
    <td align="center"><%=recDoc("No_Car_Par")%></td>
    <td align="center">
	<%
	if recDoc("Audit_Flag_Complete") = 0 then
		response.write "&#149;"
	else
		response.write "&nbsp;"
	end if
	%>    </td>
    <td align="center">
	<%
	 ' if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
	 '	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	   '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDate) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&a(i)&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">"&a(i)&"<br>"&"</font>" 
			end if
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	 %>     </td>
    <td align="center">
    <%
	 ' if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
	 '	response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 'else
	 '	response.write "&nbsp;"
	 'end if
	  '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDate) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">&nbsp;<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDate) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">&nbsp;<br />"&"</font>" 
			end if
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	 %>
    </td>
    <td align="center">
	<%
'	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
'	 	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
'	 else
'	 	response.write "&nbsp;"
'	 end if
	 '------------------------------------------------Start block alert PAR No-------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	  
	       getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
		b=Split(getPAR,"<br>")
		for i=0 to Ubound(b)
			getFinishDatePar = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDatePar) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""#000000"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&b(i)&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">"&b(i)&"<br>"&"</font>" 
			end if
		next
	 	'response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '-------------------------------------------------End block alert PAR No--------------------------------------------------
	 %>     </td>
    <td align="center"><%
'	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
'	 	response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
'	 else
'	 	response.write "&nbsp;"
'	 end if
	 '------------------------------------------------Start block alert PAR No-------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" and recDoc("Audit_Flag_Complete") <> "0" then
	  
	       getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
		b=Split(getPAR,"<br>")
		for i=0 to Ubound(b)
			getFinishDatePar = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&b(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if isDate(getFinishDatePar) = true and isDate(getAuditDate2) = False and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDate)&"/"&Day(getFinishDate)&"/"&(Year(getFinishDate)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = False then
					
					FinishDate = cDate(Day(getFinishDatePar)&"/"&Month(getFinishDatePar)&"/"&Year(getFinishDatePar))
					if DateDiff("d",Datemmddyyyy,FinishDate) > 7 and  DateDiff("d",Datemmddyyyy,FinishDate) < 16 then
						response.write "<font color=""#F0A904"">"
					elseif  DateDiff("d",Datemmddyyyy,FinishDate) >= 0 and  DateDiff("d",Datemmddyyyy,FinishDate) < 8 then
						response.write "<font color=""#FF0000"">"
					else
						response.write "<font color=""black"">"
					end if
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			elseif  isDate(getFinishDatePar) = true and isDate(getAuditDate2) = true and isDate(getAuditDate3) = true then 
					response.write "<font color=""#009900"">"
					response.write "<span  title=""วันที่ครบกำหนด : "&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&""">"&cDate(Month(getFinishDatePar)&"/"&Day(getFinishDatePar)&"/"&(Year(getFinishDatePar)+543))&"<br>"&"</span> "
					response.write "</font>"
			else
					response.write "<font color=""#000000"">&nbsp;<br></font>" 
			end if
		next
	 	'response.write getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '-------------------------------------------------End block alert PAR No--------------------------------------------------
	 %></td>
    <td align="center">
    <%
	 '---------------------------------------------Start block CAR No---------------------------------------------------------
	 if getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
	 	getCAR = getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 	
		a=Split(getCAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if getFinishDate <> "" then
			 '       if cDate(Datemmddyyyy) < cDate(Month(getFinishDate)&"/"&day(getFinishDate)&"/"&Year(getFinishDate)) then
							FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
							if IsDate(getAuditDate2) <> False and IsDate(getAuditDate3) <> False then
								response.write "<font color=""#009900"">"
								response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&""">"&a(i)&"<br>"&"</span>"
								response.write "</font>"
							else
								response.write "&nbsp;"
							end if
			'		else
			'				response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&"""><font color=""#EC0076"">"&a(i)&"<br>"&"</font></span>" 
			'		end if
			else
					response.write "&nbsp;" 
			end if
			
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write "&nbsp;"
	 end if
	 '----------------------------------------------end block alert CAR No---------------------------------------------------
	%>
    <%
	 '---------------------------------------------Start block PAR No---------------------------------------------------------
	 if getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2") <> "" then
	 	getPAR = getPARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"2")
	 	
		a=Split(getPAR,"<br>")
		for i=0 to Ubound(a)
			getFinishDate = GetSingleFieldQS("Tb_InternalAudit","Audit_Finish_Date"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate2 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date2"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			getAuditDate3 = GetSingleFieldQS("Tb_InternalAudit","Audit_Date3"," Where  No_Car_Par='"&a(i)&"' and M_Code='"&recDoc("M_Code")&"' ")
			if getFinishDate <> "" then
			    '    if cDate(Datemmddyyyy) < cDate(Month(getFinishDate)&"/"&day(getFinishDate)&"/"&Year(getFinishDate)) then
							FinishDate = cDate(Day(getFinishDate)&"/"&Month(getFinishDate)&"/"&Year(getFinishDate))
							if IsDate(getAuditDate2) <> False and IsDate(getAuditDate3) <> False then
								response.write "<font color=""#009900"">"
								response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&""">"&a(i)&"<br>"&"</span>"
								response.write "</font>"
							else
								response.write "&nbsp;"
							end if
				'	else
				'			response.write "<span  title=""วันที่ครบกำหนด : "&cDate(getFinishDate)&"""><font color=""#EC0076"">"&a(i)&"<br>"&"</font></span>" 
				'	end if
			else
					response.write "&nbsp;" 
			end if
			
		next
	 	'response.write getCARNumber(recDoc("Audit_Depart"),recDoc("M_Code"),"1")
	 else
	 	response.write ""
	 end if
	 '----------------------------------------------end block alert PAR No---------------------------------------------------
	%>
    </td>
  </tr>
	<%
		recDoc.MoveNext
		Wend
		recDoc.Close()
	  
	 recDepart2.MoveNext
	 Wend
	if recDepart2.RecordCount = 0 then
%>
  <tr><td colspan="11" align="center">No Data</td></tr>
  <% 
  	 end if
	 recDepart2.Close()
  end if
 %>
</table>
<div align="left">หมายเหตุ : ทะเบียนควบคุมสถานะการตรวจติดตามคุณภาพภายในและการปฏิบัติการแก้ไข/ป้องกัน (F-FDA-T-18) จะแสดงผลโดยอัตโนมัติ หากมีการบันทึกข้อมูลการตรวจติดตามคุณภาพ<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;ในระบบรายงานการตรวจติดตามคุณภาพภายใน ซึ่งสามารถเข้าดูได้ที่ระบบผลการดำเนินงานระบบคุณภาพ – ผลการดำเนินงานตรวจติดตามคุณภาพภายใน</div>
<div align="right">F-FDA-T-18 (1-01/06/57) หน้า.../...</div>
<br />
<table width="100%" cellpadding="0" cellspacing="4" border="0">
<tr>
  <td colspan="3"><b>หมายเหตุ :</b> ความหมายสีตัวอักษรสำหรับ CAR No. / PAR No.</td>
  </tr>
<tr>
<td width="5%"><div align="right">&nbsp;</div></td>
<td bgcolor="#F0A904" width="5%"><div align="center" style="border: thin">สีเหลือง</div></td>
<td width="90%"> หมายถึง เหลือเวลาอีก 15 วันก่อนที่จะถึงกำหนดแล้วเสร็จ</td>
</tr>
<tr>
<td width="5%"><div align="right"></div></td>
<td width="5%" bgcolor="#FF0000"><div align="center" style="border:thin">สีแดง&nbsp;</div></td>
<td width="90%"> หมายถึง เหลือเวลาอีก 7 วันก่อนที่จะถึงกำหนดแล้วเสร็จ</td>
</tr>
<tr>
<td width="5%"><div align="right"></div></td>
<td width="5%" bgcolor="#009900"><div align="center" style="border:thin">สีเขียว&nbsp;</div></td>
<td width="90%"> หมายถึง มีการตรวจติดตามซ้ำเรียบร้อยแล้ว</td>
</tr>
</table>
</body>
</html>
