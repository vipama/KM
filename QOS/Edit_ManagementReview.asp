<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
' # start code for check permission in DB 
if Session("member") <> getPermission(session("member"),"L_Email") or isnull(session("member")) = true or session("member") = "" or isEmpty(session("member")) = True then
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
else
	session("Depart") = getPermission(session("member"),"D_Id")
end if
' # End code for check permission in DB
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
getID = Request.QueryString("ID")
getDID = Request.QueryString("DID")
getSave = Request.Form("hidSave")
if isEmpty(getDID) = true then
getDID = 1
end if
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if


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
getMRID = Request.Form("hidMRID")
sql = "Update Tb_ManagementReview set MR_Level='"&getLevel&"',D_Id='"&getDepartID&"',MR_Date='"&fullDate&"',MR_Review1='"&getTxtReview1&"',MR_Review2='"&getTxtReview2&"',MR_Review3='"&getTxtReview3&"',MR_Review4='"&getTxtReview4&"',MR_Review5='"&getTxtReview5&"',MR_Review6='"&getTxtReview6&"',MR_Review7='"&getTxtReview7&"',MR_Record='"&gettxtName&"',Flag_Show=True,MR_CountMeeting='"&getC_CountFull&"'  where  MR_ID="&getMRID&" and D_Id='"&getDepartID&"' "

sqlLog = "Insert into Tb_ManagementReviewLog (D_Id,UserName,IP,Log_Date,Log_Method,MR_ID) values ('"&getDepartID&"','"&session("member")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Datemmddyyyy&"','Update','"&getMRID&"')"

'response.write sql&"<br />"
ConQS.execute(sql)

ConQS.execute(sqlLog)
'response.write sqlLog&"<br />"
If Err.Number = 0 Then
	response.write "<script language=""javascript"">"
	response.write "alert(""Update Data Success"");"
	response.write "window.location.href='ManagementReview.asp?Id="&getDepartID&"'; "
	response.write "</script>"
end if
getSave=""
end if
'-------------------------------------------------------------------End of block Save Data---------------------------------------------------------------------

if isEmpty(session("member")) = False and isEmpty(getID) = False and isEmpty(getSave) = True  then
 set RecView = Server.CreateObject("ADODB.RECORDSET")
 sql = "select * from Tb_ManagementReview where MR_ID="&getID&" and  Flag_Show = True"
 RecView.open sql,conQS,1,3
 strSplit = split(RecView("MR_CountMeeting"),"/")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>แก้ไขข้อมูลรายงานการประชุมทบทวนโดยฝ่ายบริหาร</title>
<style type="text/css">
<!--
.text {					Font-size:14px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:14px}
.textsmall {
					Font-size:10px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:12px}
-->
</style>
<script  type="text/jscript" src="jScript/JS.js"></script>
</head>

<body>
<form name="frmManagementReview" id="ManageMentReview" enctype="application/x-www-form-urlencoded" >
<input type="hidden" name="hidSave" id="hidSave" value="" />
<input type="hidden" name="DepartID" id="DepartID" value="<%=RecView("D_Id")%>"  />
<input  type="hidden" name="hidMRID" id="hidMRID" value="<%=RecView("MR_ID")%>"  />
<table width="85%" border="0" align="center" cellpadding="2" cellspacing="3">
  <tr>
    <td width="45%" class="text">รายการประชุมทบทวนโดยฝ่ายบริหาร</td>
    <td width="55%" class="text"><label>
      <input type="radio" name="radio" id="radioDepart" value="1" <% if RecView("MR_Level") = "1" then response.write "checked=""checked"" " end if %> disabled="disabled" />
      ระดับกรม </label>
        <label> &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="radio" name="radio"  id="radioSubDepart" value="2" <% if RecView("MR_Level") = "2" then response.write "checked=""checked"" " end if %> disabled="disabled" />
          ระดับกอง </label></td>
  </tr>
  <tr>
    <td class="text">กอง / สำนัก</td>
    <td class="text">
		  <%
		if getDID = "01" or getDID="02" then
		  %>
			  <select name="Depart_ID" id="Depart_ID" onChange="ChangeJobresultGroup(this.value,1)"  style="font-size:14px" class="text" disabled="disabled"   >
			  <option value="01"  <% if getDID = "01" then response.write " selected=""selected"" " end if %> >คณะกรรมการบริหารระบบคุณภาพ</option>
              <option value="02"  <% if getDID = "02" then response.write " selected=""selected"" " end if %>>คณะกรรมการประสานงานระบบคุณภาพ</option>
			  </select> 
		<%
        else
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
        <select name="Depart_ID" id="Depart_ID" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:14px" class="text"  disabled="disabled"  >
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
		end if
			  %>
        </select>
    </td>
  </tr>
  <tr>
    <td class="text">วันที่ประชุม</td>
    <td class="text"><select name="DayStart" size="1" id="DayStart" class="text">
      <% For i=1 to 31%>
      <option value="<%=i%>" <% if Day(RecView("MR_Date")) = i then response.write "selected=""selected"" " end if %>><%=i%></option>
      <% Next %>
    </select>
      &nbsp;&nbsp;&nbsp;
      <select name="MonthStart" id="MonthStart" class="text">
        <option value="1" <% if Month(RecView("MR_Date")) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
        <option value="2" <% if Month(RecView("MR_Date")) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
        <option value="3" <% if Month(RecView("MR_Date")) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
        <option value="4" <% if Month(RecView("MR_Date")) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
        <option value="5" <% if Month(RecView("MR_Date")) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
        <option value="6" <% if Month(RecView("MR_Date")) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
        <option value="7" <% if Month(RecView("MR_Date")) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
        <option value="8" <% if Month(RecView("MR_Date")) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
        <option value="9" <% if Month(RecView("MR_Date")) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
        <option value="10" <% if Month(RecView("MR_Date")) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
        <option value="11" <% if Month(RecView("MR_Date")) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
        <option value="12" <% if Month(RecView("MR_Date")) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
      </select>
      &nbsp;&nbsp;&nbsp;
      <select name="YearStart" id="YearStart" class="text">
        <% For q=2014 to 2020 %>
        <option value="<%=q%>" <% if Year(RecView("MR_Date")) = q then  response.write " selected=""selected"" " end if %>><%=q+543%></option>
        <% Next %>
      </select>
    </td>
  </tr>
  <tr>
    <td class="text">การประชุมครั้งที่ </td>
    <td><label>
      <select name="C_Count" id="C_Count" class="text">
      <% for i=1 to 20 %>
        <option value="<%=i%>" <% if strSplit(0) = cStr(i) then response.write " selected=""selected"" " end if %> ><%=i%></option>
      <% Next %>
      </select>
      &nbsp;&nbsp;/ &nbsp;&nbsp;
      <select name="C_Year" id="C_Year" class="text">
      <% for j=2557 to 2565 %>
        <option value="<%=j%>" <% if strSplit(1) = cStr(j) then response.write " selected=""selected"" " end if %> ><%=j%></option>
      <% Next %>
      </select>
    </label></td>
  </tr>
  <tr>
    <td class="text">1. ผลการตรวจติดตามคุณภาพ</td>
    <td><label>
      <textarea name="txtReview1" id="txtReview1" cols="100" rows="3"><%=RecView("MR_Review1")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">2. ข้อร้องเรียน ข้อคิดเห็น ของผู้รับบริการ</td>
    <td><label>
      <textarea name="txtReview2" id="txtReview2" cols="100" rows="3"><%=RecView("MR_Review2")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">3. ผลการดำเนินการตามเป้าหมายและตัวชี้วัด</td>
    <td><label>
      <textarea name="txtReview3" id="txtReview3" cols="100" rows="3"><%=RecView("MR_Review3")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">4. สถานะของการปฏิบัติการแก้ไขและป้องกัน</td>
    <td><label>
      <textarea name="txtReview4" id="txtReview4" cols="100" rows="3"><%=RecView("MR_Review4")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">5. การติดตามผลจากการประชุมที่ผ่านมา</td>
    <td><label>
      <textarea name="txtReview5" id="txtReview5" cols="100" rows="3"><%=RecView("MR_Review5")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">6. การเปลี่ยนแปลงที่อาจมีผลกระทบต่อระบบ</td>
    <td><label>
      <textarea name="txtReview6" id="txtReview6" cols="100" rows="3"><%=RecView("MR_Review6")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">7. ข้อเสนอแนะเพื่อการปรับปรุง</td>
    <td><label>
      <textarea name="txtReview7" id="txtReview7" cols="100" rows="3"><%=RecView("MR_Review7")%></textarea>
    </label></td>
  </tr>
  <tr>
    <td class="text">ผู้บันทึกการประชุม</td>
    <td><label>
      <input name="txtName" type="text" id="txtName" size="60" value="<% if getDid = "01" or getDid = "02" then response.write "พิมพ์ธิดา วงศ์สุนทร" else response.write RecView("MR_Record") end if %>" />
    </label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><label>
      <input type="button" name="butSave" id="butSave" value="บันทึกข้อมูล"  onclick="Edit_ManagementReview_goSave()" />
      &nbsp;&nbsp;
      <input type="button" name="butCancel" id="butCancel" value="ย้อนกลับ" onClick="javascript:{ window.location.href='ManagementReview.asp?id=<%=RecView("D_Id")%>'; }" />
    </label></td>
  </tr>
</table>
</form>
<%
else
response.write "Wrong parameter  Please  back to previous page"
end if
%>
</body>
</html>
