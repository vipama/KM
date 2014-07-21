<!--#include file="admin/Config.inc.asp"-->
<%
dim getSave
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
getSave = Request.Form("hidFlagSave")
if getSave = "Save" then

'getB_Date = Month(Request.Form("hidBDate"))&"/"&day(Request.Form("hidBDate"))&"/"&year(Request.Form("hidBDate"))
getB_Id = Request.Form("Bid")
getBDate = GetSingleField("Tb_Book","B_Date"," where B_ID="&getB_Id&" ")
getB_Date = Month(getBDate)&"/"&day(getBDate)&"/"&(year(getBDate)+543)

'getB_Topic = Request.Form("txtHead")
'getB_Meeting = Month(getB_Date)&"/"&day(getB_Date)&"/"&year(getB_Date)
'getB_TimeStart = Request.Form("BookHour")&":"&Request.Form("BookMinute") 
'getB_TimeEnd = Request.Form("BookHourEnd")&":"&Request.Form("BookMinuteEnd")
'getB_Name = Request.Form("txtSubscribers")
'getB_Tel  = Request.Form("txtSubscribersTel")

getB_Topic = Request.Form("txtHead")
getB_RoomName = Request.Form("txtRoomName")
getB_StartDate = Request.Form("MonthStart")&"/"&Request.Form("DayStart")&"/"&Request.Form("YearStart")
getB_EndDate =  Request.Form("MonthEnd")&"/"&Request.Form("DayEnd")&"/"&Request.Form("YearEnd")
getB_TimeStart = Request.Form("StartHour")&":"&Request.Form("StartMinute")
getB_TimeEnd = Request.Form("EndHour")&":"&Request.Form("EndMinute")
getB_Name = Request.Form("txtSubscribers")
getB_Tel  = Request.Form("txtSubscribersTel")

getHid_BName = Request.Form("hidtxtSubscribers")
getHid_BTel = Request.Form("hidtxtSubscribersTel")

	'Sql =  "update  Tb_Book set  B_TimeStart='"&getB_TimeStart&"' , B_TimeEnd='"&getB_TimeEnd&"' , B_Topic='"&getB_Topic&"'  , B_Name='"&getB_Name&"' , B_Tel='"&getB_Tel&"' , B_IP='"&Request.ServerVariables("REMOTE_ADDR")&"'  where B_ID="&getB_Id&" "
	Sql =  "update  Tb_Book set  B_Topic='"&getB_Topic&"' , B_RoomName='"&getB_RoomName&"' , B_StartDate='"&getB_StartDate&"' , B_EndDate='"&getB_EndDate&"' , B_TimeStart='"&getB_TimeStart&"' , B_TimeEnd='"&getB_TimeEnd&"' , B_Name='"&getB_Name&"' , B_Tel='"&getB_Tel&"' , B_IP='"&Request.ServerVariables("REMOTE_ADDR")&"'  where B_ID="&getB_Id&" "
	'response.write Sql&"<br />"
	ConActivity.execute(sql)
	
Sqllog =  "insert into Tb_BookLog (B_ID,L_DateAdd,L_IP,B_Name,B_Tel,L_Method) values ('"&getB_Id&"','"&Datemmddyyyy&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&getHid_BName&"','"&getHid_BTel&"','Update')"
	ConActivity.execute(sqllog)
	
		If Err.Number = 0 Then
		response.write "<script language=""javascript"">"
		response.write "alert(""Save Data Success"");"
		response.write "window.opener.location.href=""CalendarBooking.asp?date="&Day(getBDate)&"/"&Month(getBDate)&"/"&(year(getBDate)+543)&" ""; "
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
<title>ปฎิทินการจองห้องประชุมกองแผนงานและวิชาการ</title>
<script language="javascript" src="admin/javascript/mainscript.js"></script>
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
		document.frmcalendarbooking.action="EditCalendarBooking.asp";
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
if isEmpty(Request.Form("hidBID")) = false then
	getBID = Request.Form("hidBID")
	set Rec = Server.CreateObject("ADODB.RECORDSET")
	SQL = "Select * from Tb_Book where B_ID="&getBID&" "
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
<input type="hidden" name="Bid" id="Bid" value="<%=getBID%>" />
<input type="hidden"  name="hidtxtSubscribers" id="hidtxtSubscribers" value="<%=get_txtSubscribers%>"/>
<input type="hidden"  name="hidtxtSubscribersTel" id="hidtxtSubscribersTel" value="<%=get_txtSubscribersTel%>"/>
<table border="0" cellpadding="3" cellspacing="0" width="85%" align="center">
<tr>
  <td colspan="4" align="left">จองห้องประชุมกองแผนงานและวิชาการ</td>
</tr>
<!--<tr>
<td width="30%">ห้องประชุม</td><td width="70%"><input  name="txtRoom" type="text" id="txtRoom" size="40" /></td>
</tr>-->
<tr>
<td width="15%">หัวข้อการประชุม</td>
<td width="35%"><textarea name="txtHead" cols="60" rows="3" id="txtHead"><%=Rec("B_Topic")%></textarea></td>
<td width="15%">ห้องประชุม / สถานที่</td>
<td width="35%"><input type="text"  value="<%=Rec("B_RoomName")%>"  name="txtRoomName" id="txtRoomName" readonly="readonly" size="60"/></td>
</tr>
    <tr>
      <td>วันที่ (วัน/เดือน/ปี)</td>
      <td><label>เริ่ม&nbsp;&nbsp;&nbsp;&nbsp;
          <select name="DayStart" size="1" id="DayStart">
        	<% For i=1 to 31%>
  			<option value="<%=i%>" <% if Day(Rec("B_StartDate")) = i then response.write "selected=""selected"" " end if %>><%=i%></option>
			<% Next %>
          </select>
&nbsp;&nbsp;&nbsp;
<select name="MonthStart" id="MonthStart">
  <option value="1" <% if Month(Rec("B_StartDate")) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
  <option value="2" <% if Month(Rec("B_StartDate")) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
  <option value="3" <% if Month(Rec("B_StartDate")) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
  <option value="4" <% if Month(Rec("B_StartDate")) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
  <option value="5" <% if Month(Rec("B_StartDate")) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
  <option value="6" <% if Month(Rec("B_StartDate")) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
  <option value="7" <% if Month(Rec("B_StartDate")) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
  <option value="8" <% if Month(Rec("B_StartDate")) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
  <option value="9" <% if Month(Rec("B_StartDate")) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
  <option value="10" <% if Month(Rec("B_StartDate")) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
  <option value="11" <% if Month(Rec("B_StartDate")) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
  <option value="12" <% if Month(Rec("B_StartDate")) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearStart" id="YearStart">
<% For q=2014 to 2020 %>
  <option value="<%=q%>" <% if Year(Rec("B_StartDate")) = q then  response.write " selected=""selected"" " end if %>><%=q+543%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;<br />
สิ้นสุด
<select name="DayEnd" size="1" id="DayEnd">
 <% For e=1 to 31%>
  <option value="<%=e%>" <% if Day(Rec("B_EndDate")) = e then response.write "selected=""selected"" " end if %>><%=e%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;
<select name="MonthEnd" id="MonthEnd">
  <option value="1" <% if Month(Rec("B_EndDate")) = 1 then response.write "selected=""selected"" " end if %>>มกราคม</option>
  <option value="2" <% if Month(Rec("B_EndDate")) = 2 then response.write "selected=""selected"" " end if %>>กุมภาพัน</option>
  <option value="3" <% if Month(Rec("B_EndDate")) = 3 then response.write "selected=""selected"" " end if %>>มีนาคม</option>
  <option value="4" <% if Month(Rec("B_EndDate")) = 4 then response.write "selected=""selected"" " end if %>>เมษายน</option>
  <option value="5" <% if Month(Rec("B_EndDate")) = 5 then response.write "selected=""selected"" " end if %>>พฤษภาคม</option>
  <option value="6" <% if Month(Rec("B_EndDate")) = 6 then response.write "selected=""selected"" " end if %>>มิถุนายน</option>
  <option value="7" <% if Month(Rec("B_EndDate")) = 7 then response.write "selected=""selected"" " end if %>>กรกฎาคม</option>
  <option value="8" <% if Month(Rec("B_EndDate")) = 8 then response.write "selected=""selected"" " end if %>>สิงหาคม</option>
  <option value="9" <% if Month(Rec("B_EndDate")) = 9 then response.write "selected=""selected"" " end if %>>กันยายน</option>
  <option value="10" <% if Month(Rec("B_EndDate")) = 10 then response.write "selected=""selected"" " end if %>>ตุลาคม</option>
  <option value="11" <% if Month(Rec("B_EndDate")) = 11 then response.write "selected=""selected"" " end if %>>พฤศจิกายน</option>
  <option value="12" <% if Month(Rec("B_EndDate")) = 12 then response.write "selected=""selected"" " end if %>>ธันวาคม</option>
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearEnd" id="YearEnd">
<% For qq=2014 to 2020 %>
  <option value="<%=qq%>" <% if Year(Rec("B_EndDate")) = qq then  response.write " selected=""selected"" " end if %>><%=qq+543%></option>
<% Next %>
</select>
      </label></td>
      <td>เวลา </td>
      <td><label for="StartHour">
        <select name="StartHour" id="StartHour">
        <%  For g=6 to 22 %>
        	 <% if g < 10 then %>
        	  <option value="0<%=g%>" <% if Hour(Rec("B_TimeStart")) = g then  response.write " selected=""selected"" " %>>0<%=g%></option>
              <% else %>
              <option value="<%=g%>" <% if Hour(Rec("B_TimeStart")) = g then  response.write " selected=""selected"" " %>><%=g%></option>
              <% end if %>
		<% Next %>
        </select>
: </label>
        <label for="StartMinute"></label>
        <select name="StartMinute" id="StartMinute">
        <% For T=0 to 60 %>
           <% if T < 10 then %>
          <option value="0<%=T%>" <% if Minute(Rec("B_TimeStart")) = T then response.write " selected=""selected"" " end if%>>0<%=T%></option>
          <% else %>
          <option value="<%=T%>" <% if Minute(Rec("B_TimeStart")) = T then response.write " selected=""selected"" " end if%>><%=T%></option>
          <% end if %>
         <% Next %>  
        </select>
&nbsp;
ถึง&nbsp;&nbsp;
<select name="EndHour" id="EndHour">
  <%  For gg=6 to 22 %>
        	 <% if gg < 10 then %>
        	  <option value="0<%=gg%>" <% if Hour(Rec("B_TimeEnd")) = gg then  response.write " selected=""selected"" " %>>0<%=gg%></option>
              <% else %>
              <option value="<%=gg%>" <% if Hour(Rec("B_TimeEnd")) = gg then  response.write " selected=""selected"" " %>><%=gg%></option>
              <% end if %>
		<% Next %>
</select>
:
<select name="EndMinute" id="EndMinute">
  <% For TT=0 to 60 %>
           <% if TT < 10 then %>
          <option value="0<%=TT%>" <% if Minute(Rec("B_TimeEnd")) = TT then response.write " selected=""selected"" " end if%>>0<%=TT%></option>
          <% else %>
          <option value="<%=TT%>" <% if Minute(Rec("B_TimeEnd")) = TT then response.write " selected=""selected"" " end if%>><%=TT%></option>
          <% end if %>
         <% Next %>  
</select></td>
    </tr>
    <tr>
    <td>ชื่อผู้รับผิดชอบ / หน่วยงาน</td>
    <td colspan="3"><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="80" value="<%=Rec("B_Name")%>" /></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td colspan="3"><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="80" value="<%=Rec("B_Tel")%>" /></td>
    </tr>
    <tr><td colspan="4" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onclick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle" id="butCancle" value="ยกเลิก" onclick="javascript:{window.close();}" /></td></tr>
</table>
</form>
<br />
<%
end if
%>
</body>
</html>

