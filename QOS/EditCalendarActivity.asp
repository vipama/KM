<!--#include file="../../Config.inc.asp"-->
<%
dim getSave
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
getSave = Request.Form("hidFlagSave")
if getSave = "Save" then

'getA_Date = Month(Request.Form("hidBDate"))&"/"&day(Request.Form("hidBDate"))&"/"&year(Request.Form("hidBDate"))
getA_Id = Request.Form("Aid")
getADate = GetSingleField("Tb_Activity","A_Date"," where A_ID="&getA_Id&" ")
getA_Date = Month(getADate)&"/"&day(getADate)&"/"&(year(getADate)+543)

getA_Topic = Request.Form("txtHead")
getA_RoomName = Request.Form("txtRoomName")
getA_StartDate = Request.Form("MonthStart")&"/"&Request.Form("DayStart")&"/"&Request.Form("YearStart")
getA_EndDate =  Request.Form("MonthEnd")&"/"&Request.Form("DayEnd")&"/"&Request.Form("YearEnd")
getA_TimeStart = Request.Form("StartHour")&":"&Request.Form("StartMinute")
getA_TimeEnd = Request.Form("EndHour")&":"&Request.Form("EndMinute")
getA_Name = Request.Form("txtSubscribers")
getA_Tel  = Request.Form("txtSubscribersTel")

getHid_AName = Request.Form("hidtxtSubscribers")
getHid_ATel = Request.Form("hidtxtSubscribersTel")

	Sql =  "update  Tb_Activity set  A_Topic='"&getA_Topic&"' , A_RoomName='"&getA_RoomName&"' , A_StartDate='"&getA_StartDate&"' , A_EndDate='"&getA_EndDate&"' , A_StartTime='"&getA_TimeStart&"' , A_EndTime='"&getA_TimeEnd&"' , A_Name='"&getA_Name&"' , A_Tel='"&getA_Tel&"' , A_IP='"&Request.ServerVariables("REMOTE_ADDR")&"'  where A_ID="&getA_Id&" "
	'response.write Sql&"<br />"
	ConActivity.execute(sql)
	
Sqllog =  "insert into Tb_ActivityLog (A_ID,L_DateAdd,L_IP,A_Name,A_Tel,A_Method) values ('"&getA_Id&"','"&Datemmddyyyy&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&getHid_AName&"','"&getHid_ATel&"','Update')"
	ConActivity.execute(sqllog)
	
		If Err.Number = 0 Then
		response.write "<script language=""javascript"">"
		response.write "alert(""Save Data Success"");"
		response.write "window.opener.location.href=""CalendarActivity.asp?date="&Day(getADate)&"/"&Month(getADate)&"/"&(year(getADate)+543)&" ""; "
		response.write "window.close();"
		'response.write "window.location.href=""CalendarBooking.asp"" "
		response.write "</script>"
		end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>ปฎิทินกิจกรรม</title>
<script language="javascript" src="jScript/JS.js"></script>
<script language="javascript">
function goSave()
{	
	
	if (alltrim(document.getElementById("txtSubscribers").value).length == 0 )
	{
		alert("กรุณาตรวจสอบ ชื่อผู้จองด้วย");
	}
	else if(isNumber(alltrim(document.getElementById("txtSubscribersTel").value)) == false)
	{
		alert("กรุณากรอกเบอร์โทรศัพท์อีกครั้ง");
	}
	else
	{
		document.frmcalendarbooking.action="EditCalendarActivity.asp";
		document.frmcalendarbooking.hidFlagSave.value="Save";
		document.frmcalendarbooking.submit();	
	}
}
</script>
</head>

<body>
<!-- Outer Table is simply to get the pretty border-->
<BR>
<%
if isEmpty(Request.Form("hidAID")) = false then
	getAID = Request.Form("hidAID")
	set Rec = Server.CreateObject("ADODB.RECORDSET")
	SQL = "Select * from Tb_Activity where A_ID="&getAID&" "
	'response.write SQL&"<br />"
	Rec.open SQL,ConActivity,1,3
	if Rec.RecordCount <= 0 then
		response.write "No Data"
	end if 
	get_txtSubscribers = Request.Form("txtSubscribers")
	get_txtSubscribersTel = Request.Form("txtSubscribersTel")
%>
<form name="frmcalendarbooking" id="frmcalendarbooking" enctype="application/x-www-form-urlencoded" method="post">
<input type="hidden"  name="hidFlagSave" id="hidFlagSave" value=""/>
<input type="hidden" name="Aid" id="Aid" value="<%=getAID%>" />
<input type="hidden"  name="hidtxtSubscribers" id="hidtxtSubscribers" value="<%=get_txtSubscribers%>"/>
<input type="hidden"  name="hidtxtSubscribersTel" id="hidtxtSubscribersTel" value="<%=get_txtSubscribersTel%>"/>
<!--<table border="1" cellpadding="3" cellspacing="0" width="30%" align="center">
<tr>
  <td colspan="2" align="center">กิจกรรม</td>
</tr>
<tr>
<td width="30%">หัวข้อ</td><td width="70%"><textarea name="txtHead" cols="40" rows="3" id="txtHead"><%'=Rec("A_Topic")%></textarea></td>
</tr>
<tr>
  <td width="30%">รายละเอียด</td>
  <td width="70%"><label>
    <textarea name="txtDes" cols="40" rows="3" id="txtDes"><%'=Rec("A_Descript")%></textarea>
  </label></td>
</tr>
    <tr>
    <td>ชื่อ</td>
    <td><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="40"  value="<%'=Rec("A_Name")%>"/></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="40" value="<%'=Rec("A_Tel")%>" /></td>
    </tr>
    <tr><td colspan="2" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onclick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle" id="butCancle" value="ยกเลิก" onclick="javascript:{ window.location.href='index.asp';}" /></td></tr>
</table>-->
<!--Start new version-->
<table border="0" cellpadding="3" cellspacing="0" width="85%" align="center">
<tr>
  <td colspan="4" align="left">ปฎิทินกิจกรรม</td>
</tr>
<!--<tr>
<td width="30%">ห้องประชุม</td><td width="70%"><input  name="txtRoom" type="text" id="txtRoom" size="40" /></td>
</tr>-->
<tr>
<td width="15%">หัวข้อการประชุม</td>
<td width="35%"><textarea name="txtHead" cols="60" rows="3" id="txtHead"><%=Rec("A_Topic")%></textarea></td>
<td width="15%">ห้องประชุม / สถานที่</td>
<td width="35%"><textarea name="txtRoomName" cols="60" rows="3" id="txtRoomName"><%=Rec("A_RoomName")%></textarea></td>
</tr>
    <tr>
      <td>วันที่ (วัน/เดือน/ปี)</td>
      <td><label>เริ่ม&nbsp;&nbsp;&nbsp;&nbsp;
          <select name="DayStart" size="1" id="DayStart">
        	<% For i=1 to 31%>
  			<option value="<%=i%>" <% if Day(Rec("A_StartDate")) = i then response.write "selected=""selected"" " end if %>><%=i%></option>
			<% Next %>
          </select>
&nbsp;&nbsp;&nbsp;
<select name="MonthStart" id="MonthStart">
  <option value="1" <% if Month(Rec("A_StartDate")) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
  <option value="2" <% if Month(Rec("A_StartDate")) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
  <option value="3" <% if Month(Rec("A_StartDate")) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
  <option value="4" <% if Month(Rec("A_StartDate")) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
  <option value="5" <% if Month(Rec("A_StartDate")) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
  <option value="6" <% if Month(Rec("A_StartDate")) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
  <option value="7" <% if Month(Rec("A_StartDate")) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
  <option value="8" <% if Month(Rec("A_StartDate")) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
  <option value="9" <% if Month(Rec("A_StartDate")) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
  <option value="10" <% if Month(Rec("A_StartDate")) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
  <option value="11" <% if Month(Rec("A_StartDate")) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
  <option value="12" <% if Month(Rec("A_StartDate")) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearStart" id="YearStart">
<% For q=2014 to 2020 %>
  <option value="<%=q%>" <% if Year(Rec("A_StartDate")) = q then  response.write " selected=""selected"" " end if %>><%=q+543%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;<br />
สิ้นสุด
<select name="DayEnd" size="1" id="DayEnd">
 <% For e=1 to 31%>
  <option value="<%=e%>" <% if Day(Rec("A_EndDate")) = e then response.write "selected=""selected"" " end if %>><%=e%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;
<select name="MonthEnd" id="MonthEnd">
  <option value="1" <% if Month(Rec("A_EndDate")) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
  <option value="2" <% if Month(Rec("A_EndDate")) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
  <option value="3" <% if Month(Rec("A_EndDate")) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
  <option value="4" <% if Month(Rec("A_EndDate")) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
  <option value="5" <% if Month(Rec("A_EndDate")) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
  <option value="6" <% if Month(Rec("A_EndDate")) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
  <option value="7" <% if Month(Rec("A_EndDate")) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
  <option value="8" <% if Month(Rec("A_EndDate")) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
  <option value="9" <% if Month(Rec("A_EndDate")) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
  <option value="10" <% if Month(Rec("A_EndDate")) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
  <option value="11" <% if Month(Rec("A_EndDate")) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
  <option value="12" <% if Month(Rec("A_EndDate")) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearEnd" id="YearEnd">
<% For qq=2014 to 2020 %>
  <option value="<%=qq%>" <% if Year(Rec("A_EndDate")) = qq then  response.write " selected=""selected"" " end if %>><%=qq+543%></option>
<% Next %>
</select>
      </label></td>
      <td>เวลา </td>
      <td><label for="StartHour">
        <select name="StartHour" id="StartHour">
        <%  For g=6 to 22 %>
        	 <% if g < 10 then %>
        	  <option value="0<%=g%>" <% if Hour(Rec("A_StartTime")) = g then  response.write " selected=""selected"" " %>>0<%=g%></option>
              <% else %>
              <option value="<%=g%>" <% if Hour(Rec("A_StartTime")) = g then  response.write " selected=""selected"" " %>><%=g%></option>
              <% end if %>
		<% Next %>
        </select>
: </label>
        <label for="StartMinute"></label>
        <select name="StartMinute" id="StartMinute">
        <% For T=0 to 60 %>
           <% if T < 10 then %>
          <option value="0<%=T%>" <% if Minute(Rec("A_StartTime")) = T then response.write " selected=""selected"" " end if%>>0<%=T%></option>
          <% else %>
          <option value="<%=T%>" <% if Minute(Rec("A_StartTime")) = T then response.write " selected=""selected"" " end if%>><%=T%></option>
          <% end if %>
         <% Next %>  
        </select>
&nbsp;
ถึง&nbsp;&nbsp;
<select name="EndHour" id="EndHour">
  <%  For gg=6 to 22 %>
        	 <% if gg < 10 then %>
        	  <option value="0<%=gg%>" <% if Hour(Rec("A_EndTime")) = gg then  response.write " selected=""selected"" " %>>0<%=gg%></option>
              <% else %>
              <option value="<%=gg%>" <% if Hour(Rec("A_EndTime")) = gg then  response.write " selected=""selected"" " %>><%=gg%></option>
              <% end if %>
		<% Next %>
</select>
:
<select name="EndMinute" id="EndMinute">
  <% For TT=0 to 60 %>
           <% if TT < 10 then %>
          <option value="0<%=TT%>" <% if Minute(Rec("A_EndTime")) = TT then response.write " selected=""selected"" " end if%>>0<%=TT%></option>
          <% else %>
          <option value="<%=TT%>" <% if Minute(Rec("A_EndTime")) = TT then response.write " selected=""selected"" " end if%>><%=TT%></option>
          <% end if %>
         <% Next %>  
</select></td>
    </tr>
    <tr>
    <td>ชื่อผู้รับผิดชอบ / หน่วยงาน</td>
    <td colspan="3"><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="80" value="<%=Rec("A_Name")%>" /></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td colspan="3"><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="80" value="<%=Rec("A_Tel")%>" /></td>
    </tr>
    <tr><td colspan="4" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onClick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle" id="butCancle" value="ยกเลิก" onClick="javascript:{ window.close();}" /></td></tr>
</table>
<!--End new version-->
</form>
<br />
<%
end if
%>
</body>
</html>

