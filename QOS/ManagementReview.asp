<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
' # start code for check permission in DB 
if isEmpty(session("member")) = True then
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้ กรุณา Login'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
'else
'	if Session("member") <> getPermission(session("member"),"L_Email") or isnull(session("member")) = true or session("member") = "" then
'		Response.write "<script>"
'		Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
'		Response.write " 	window.location.href=""default.asp""; "

		'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
'	else
'		session("Depart") = getPermission(session("member"),"D_Id")
'	end if
end if
' # End code for check permission in DB
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
getSave = Request.Form("hidSave")

'---------------------------------get id for change Department --------------------------------------------
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	if isnull(session("Depart")) = false and session("Depart") <> "100" then
			getDid = session("Depart")
		else
			getDid = "1"
	 	end if
	 end if
else
	get_DID = Request.QueryString("id")
	if get_DID <> "01" and get_DID <> "02" then
		if isnull(session("Depart")) = false and session("Depart") <> "100" then
			getDid = session("Depart")
		else
			getDid=Request.QueryString("id")
		end if
	else
		getDid=Request.QueryString("id")
	end if
end if
'-----------------------------------------------------------------------------------------------------------------
'----------------------------------get oid for change Level----------------------------------------------------
if isEmpty(Request.QueryString("oid")) = true then
	 if isEmpty(Request.Form("hidOid")) = false then
	 	getOid=Request.Form("hidOid")
	 else
	 	getOid = "2"
	 end if
else
	getOid=Request.QueryString("oid")
end if
'-----------------------------------------------------------------------------------------------------------------
If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "/" & Request.QueryString("day") & "/" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "/" & Request.QueryString("day") & "/" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
		' The elegant solution for those of you running IIS4
		'If Request.QueryString.Count <> 0 Then Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
	End If
End If
'--------------------------------------------------------------------------Start block save data--------------------------------------------------------------------------------
if getSave = "Save" then
getLevel = Request.Form("radio")
getDepartID = Request.Form("DepartID")
getDay = Request.Form("DayStart")
getMonth = Request.Form("MonthStart")
getYear = Request.Form("yearStart")
getC_Count = Request.Form("C_Count")
getC_Year = Request.Form("C_Year")
getC_CountFull = getC_Count&"/"&getC_Year
getTxtReview1 = Request.Form("txtReview1")
getTxtReview2 = Request.Form("txtReview2")
getTxtReview3 = Request.Form("txtReview3")
getTxtReview4 = Request.Form("txtReview4")
getTxtReview5 = Request.Form("txtReview5")
getTxtReview6 = Request.Form("txtReview6")
getTxtReview7 = Request.Form("txtReview7")
gettxtName = Request.Form("txtName")
fullDate = getMonth&"/"&getDay&"/"&getYear
sql = "Insert into Tb_ManagementReview (MR_Level,D_Id,MR_Date,MR_Review1,MR_Review2,MR_Review3,MR_Review4,MR_Review5,MR_Review6,MR_Review7,Flag_Show,MR_CountMeeting,MR_Record) values ('"&getLevel&"','"&getDepartID&"','"&fullDate&"','"&getTxtReview1&"','"&getTxtReview2&"','"&getTxtReview3&"','"&getTxtReview4&"','"&getTxtReview5&"','"&getTxtReview6&"','"&getTxtReview7&"',True,'"&getC_CountFull&"','"&gettxtName&"') "
'response.write sql&"<br />"
ConQS.execute(sql)

mrid = GetSingleFieldQS("Tb_ManagementReview","top 1 MR_ID"," order by MR_ID Desc ") 'get MR_ID before save log

sqlLog = "Insert into Tb_ManagementReviewLog (D_Id,UserName,IP,Log_Date,Log_Method,MR_ID) values ('"&getDepartID&"','"&session("member")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Datemmddyyyy&"','Add','"&mrid&"')"
ConQS.execute(sqlLog)
'response.write sqlLog&"<br />"

If Err.Number = 0 Then
	response.write "<script language=""javascript"">"
	response.write "alert(""Save Data Success"");"
	response.write "</script>"
end if
getSave=""
getDid = getDepartID 
end if
'-------------------------------------------------------------------End of block Save Data---------------------------------------------------------------------
'-------------------------------------------------------------------Start of block Cancel Data------------------------------------------------------------------
if getSave = "Cancel" then
	getMRID = Request.Form("hidMRID")
	getSave=""
	sql_cancel = "Update Tb_ManagementReview  set Flag_Show=False where MR_ID="&getMRID
	ConQS.execute(sql_cancel)
	
	sqlLog = "Insert into Tb_ManagementReviewLog (D_Id,UserName,IP,Log_Date,Log_Method,MR_ID) values ('"&getDid&"','"&session("member")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Datemmddyyyy&"','Cancel','"&getMRID&"')"
	ConQS.execute(sqlLog)
	
	If Err.Number = 0 Then
	response.write "<script language=""javascript"">"
	response.write "alert(""Cancel Data Success"");"
	response.write "window.location.href='ManagementReview.asp?Id="&getDid&"'; "
	response.write "</script>"
	end if
	
end if
'-------------------------------------------------------------------End of block Cancel Data-------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>การทบทวนโดยฝ่ายบริหาร</title>
<script language="javascript">
/*function ChangeJobresultGroup(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			
			window.location.href="ManagementReview.asp?id="+val+"&oid="+val1;
		}else{
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="ManagementReview.asp?id="+strUser+"&oid="+val1;
		}
		
}
function ManagementReview_goSave()
{
		document.frmManagementReview.action="ManagementReview.asp";
		document.frmManagementReview.method="POST";
		document.frmManagementReview.hidSave.value="Save";
		document.frmManagementReview.submit();
}
function ManagementReview_goViewDoc(ID,DID)
{
	window.location.href="View_ManagementReview.asp?id="+ID+"&DID="+DID;
}
function ManagementReview_goEditDoc(ID,DID)
{
		window.location.href="Edit_ManagementReview.asp?id="+ID+"&DID="+DID;
}
function ManagementReview_goCancelDoc(ID,DID)
{
		
		document.frmManagementReview.action="ManagementReview.asp";
		document.frmManagementReview.method="POST";
		document.frmManagementReview.hidSave.value="Save";
		document.frmManagementReview.submit();
}*/
</script>
<script  type="text/jscript" src="jScript/JS.js"></script>
<style>
.text {
					Font-size:14px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:14px}
.textsmall {
					Font-size:10px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:12px}
</style>
</head>

<body>
<div style="font-size:18px; font-weight:bold" align="center">แบบฟอร์มทบทวนโดยฝ่ายบริหาร</div><br />
<form name="frmManagementReview" id="ManageMentReview" enctype="application/x-www-form-urlencoded" >
<input type="hidden" name="hidSave" id="hidSave" value="" />
<input type="hidden" name="hidMRID" id="hidMRID" value="" />
<input type="hidden" name="hidDid" id="hidDid" value="<%=getDid%>" />
<input type="hidden" name="hidOid" id="hidOid" value="<%=getOid%>" />
<table width="100%" align="center" cellpadding="2" cellspacing="3">
  <tr>
    <td width="25%" class="text">รายการประชุมทบทวนโดยฝ่ายบริหาร</td>
    <td width="75%" class="text"><label>
        <input type="radio" name="radio" id="radioDepart" value="1" <% if getOid = "1" then response.write "checked=""checked"" " end if %>  onclick="ChangeJobresultGroupManagementReview('01',this.value)" />
        ระดับกรม
    </label>      <label>
      
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="radio" name="radio"  id="radioSubDepart" value="2" <% if getOid = "2" then response.write "checked=""checked"" " end if %>  onclick="ChangeJobresultGroupManagementReview(1,this.value)" />
        ระดับกอง
    </label></td>
  </tr>
  <% if getOid = "2" and getDid <> "01"  and getDid <> "02" then %>
  <tr>
    <td class="text">กอง / สำนัก</td>
    <td class="text">
    <%
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  '#  Original code  ### sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			   if session("Depart") = "100" then
					sqlDepart = "select  *  from  Tb_DepartmentPermission where D_Id not in('17','18') order by D_Numberlist  asc"
			  else
					sqlDepart = "select  *  from  Tb_DepartmentPermission where D_Id='"&getDid&"' order by D_Numberlist  asc"
			  end if
			  
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroupManagementReview(this.value,<%=getOid%>)" style="font-size:14px" class="text"   >
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
			  </select>    </td>
  </tr>
  <% else %>
  <tr>
    <td class="text">กอง / สำนัก</td>
    <td class="text">
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroupManagementReview(this.value,1)"  style="font-size:14px" class="text"   >
			  <option value="01"  <% if getDid = "01" then response.write " selected=""selected"" " end if %> >คณะกรรมการบริหารระบบคุณภาพ</option>
              <option value="02"  <% if getDid = "02" then response.write " selected=""selected"" " end if %>>คณะกรรมการประสานงานระบบคุณภาพ</option>
			  </select>    
    </td>
  </tr>
  <% end if %>
  <tr>
    <td class="text">วันที่ประชุม</td>
    <td class="text">
    <select name="DayStart" size="1" id="DayStart" class="text">
        	<% For i=1 to 31%>
  			<option value="<%=i%>" <% if Day(dDate) = i then response.write "selected=""selected"" " end if %>><%=i%></option>
			<% Next %>
    </select>&nbsp;&nbsp;&nbsp;
    <select name="MonthStart" id="MonthStart" class="text">
  <option value="1" <% if Month(dDate) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
  <option value="2" <% if Month(dDate) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
  <option value="3" <% if Month(dDate) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
  <option value="4" <% if Month(dDate) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
  <option value="5" <% if Month(dDate) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
  <option value="6" <% if Month(dDate) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
  <option value="7" <% if Month(dDate) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
  <option value="8" <% if Month(dDate) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
  <option value="9" <% if Month(dDate) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
  <option value="10" <% if Month(dDate) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
  <option value="11" <% if Month(dDate) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
  <option value="12" <% if Month(dDate) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
</select>&nbsp;&nbsp;&nbsp;
<select name="YearStart" id="YearStart" class="text">
<% For q=2014 to 2020 %>
  <option value="<%=q%>" <% if Year(dDate) = q then  response.write " selected=""selected"" " end if %>><%=q+543%></option>
<% Next %>
</select>    </td>
  </tr>
  <tr>
    <td class="text">การประชุมครั้งที่ </td>
    <td><label>
      <select name="C_Count" id="C_Count" class="text">
        <option value="1" selected="selected">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
      </select>
    &nbsp;&nbsp;/ &nbsp;&nbsp;
    <select name="C_Year" id="C_Year" class="text">
      <option value="2557">2557</option>
      <option value="2558">2558</option>
      <option value="2559">2559</option>
      <option value="2560">2560</option>
      <option value="2561">2561</option>
      <option value="2562">2562</option>
      <option value="2563">2563</option>
      <option value="2564">2564</option>
      <option value="2565">2565</option>
    </select>
    </label></td>
  </tr>
  <tr>
    <td class="text">1. ผลการตรวจติดตามคุณภาพ</td>
    <td><label>
      <textarea name="txtReview1" id="txtReview1" cols="150" rows="3" ></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">2. ข้อร้องเรียน ข้อคิดเห็น ของผู้รับบริการ</td>
    <td><label>
      <textarea name="txtReview2" id="txtReview2" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">3. ผลการดำเนินการตามเป้าหมายและตัวชี้วัด</td>
    <td><label>
      <textarea name="txtReview3" id="txtReview3" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">4. สถานะของการปฏิบัติการแก้ไขและป้องกัน</td>
    <td><label>
      <textarea name="txtReview4" id="txtReview4" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">5. การติดตามผลจากการประชุมที่ผ่านมา</td>
    <td><label>
      <textarea name="txtReview5" id="txtReview5" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">6. การเปลี่ยนแปลงที่อาจมีผลกระทบต่อระบบ</td>
    <td><label>
      <textarea name="txtReview6" id="txtReview6" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">7. ข้อเสนอแนะเพื่อการปรับปรุง</td>
    <td><label>
      <textarea name="txtReview7" id="txtReview7" cols="150" rows="3"></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">ผู้บันทึกการประชุม</td>
    <td><label>
      <input name="txtName" type="text" id="txtName" size="60" value="<% if getDid = "01" or getDid = "02" then response.write "พิมพ์ธิดา วงศ์สุนทร" end if %>" />
    </label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><label>
      <input type="button" name="butSave" id="butSave" value="บันทึกข้อมูล"  onclick="ManagementReview_goSave()" />
      &nbsp;&nbsp;
      <input type="button" name="butCancel" id="butCancel" value="ยกเลิก" onClick="javascript:{ window.location.href='/kmfda/_block/qos';}" />
    </label></td>
  </tr>
</table>
</form>
<div align="center"><a href="/kmfda/_block/qos" target="_self" style="text-decoration:none; color:#000000"><b>หน้าแรก</b></a></div><br />
<%
if Session("member") = getPermission(session("member"),"L_Email") and isnull(session("member")) <> true and session("member") <> "" then
		session("Depart") = getPermission(session("member"),"D_Id")
		
%>
<% if GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'") = getPermission(session("member"),"L_Name") then  response.write GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'") end if  %>
<table width="75%" cellpadding="3" cellspacing="0" border="1" align="center" bordercolor="#333333">
<tr>
  <td colspan="5">
  <% 
  if getDid = "01" then 
  	response.write "คณะกรรมการบริหารระบบคุณภาพ"
  elseif getDid = "02" then
    response.write "คณะกรรมการประสานงานระบบคุณภาพ"
  else
    response.write getDepartmentname(getDid)
  end if 
	%></td>
</tr>
<tr>
  <td  <% if GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'") = getPermission(session("member"),"L_Name") then %> width="60%" <% else response.write "width=""70%"" " end if %> align="center" class="text">รายละเอียด</td>
  <% if GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'") = getPermission(session("member"),"L_Name") then  %>
  <td width="10%" align="center" class="text">&nbsp;</td>
  <% end if %>
  <td width="10%">&nbsp;</td>
  <td width="10%">&nbsp;</td>
  <td width="10%">&nbsp;</td>
</tr>
<%
sql_get = "select * from  Tb_ManagementReview where D_Id='"&getDid&"' and Flag_Show=True order by MR_ID DESC  "
'response.write sql_get
set RecGet = Server.CreateObject("ADODB.RECORDSET")
RecGet.open sql_get,ConQS,1,3
While NOT RecGet.EOF
%>
<tr><td >รายงานการประชุม ครั้งที่  <%=RecGet("MR_Countmeeting")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;วันที่ : <%=RecGet("MR_Date")%></td>
  <% if GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'") = getPermission(session("member"),"L_Name") then  %>
  <td ><label>
  <input name="butCheck" type="button" class="textsmall" id="butCheck" value="ตรวจรายงาน" onClick="ManagementReview_goCheckDoc('<%=RecGet("MR_ID")%>','<%=getDid%>','')" />
</label></td>
  <% end if %>
  <td align="center"><label>
  <input name="butView" type="button" class="textsmall" id="butView" value="ดูรายงาน" onClick="ManagementReview_goViewDoc('<%=RecGet("MR_ID")%>','<%=getDid%>','')" />
</label></td>
<td align="center"><label>
  <input name="butEdit" type="button" class="textsmall" id="butEdit" value="แก้ไขเอกสาร" onClick="ManagementReview_goEditDoc('<%=RecGet("MR_ID")%>','<%=getDid%>')" />
</label></td>
<td align="center"><label>
  <input name="butCancel2" type="reset" class="textsmall" id="butCancel2" value="ยกเลิกเอกสาร" onClick="ManagementReview_goCancelDoc('<%=RecGet("MR_ID")%>','<%=getDid%>')" />
</label></td></tr>
<%
RecGet.MoveNext
Wend
if RecGet.RecordCount = 0 then
%>
<tr><td colspan="5" align="center"><b>No Data</b></td></tr>
<% end if %>
</table>
<% end if %>
</body>
</html>
