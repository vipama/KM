<!--#include file="admin/Config.inc.asp"-->
<%
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
dim getSave
getSave = Request.Form("hidFlagSave")
if getSave = "Save" then

getB_Date = Month(Request.Form("hidBDate"))&"/"&day(Request.Form("hidBDate"))&"/"&year(Request.Form("hidBDate"))
getB_Topic = Request.Form("txtHead")
getB_RoomName = Request.Form("txtRoomName")

getB_StartDate = Request.Form("MonthStart")&"/"&Request.Form("DayStart")&"/"&Request.Form("YearStart")
getB_EndDate = Request.Form("MonthEnd")&"/"&Request.Form("DayEnd")&"/"&Request.Form("YearEnd")

getB_TimeStart = Request.Form("StartHour")&":"&Request.Form("StartMinute") 
getB_TimeEnd = Request.Form("EndHour")&":"&Request.Form("EndMinute")

getB_Name = Request.Form("txtSubscribers")
getB_Tel  = Request.Form("txtSubscribersTel")

'response.write Hour(GetMax("Tb_Book","B_TimeStart"," where B_Date=#"&getB_Date&"#"))&"<br />"
'response.write Hour(GetMin("Tb_Book","B_TimeStart"," where B_Date=#"&getB_Date&"#"))&"<br />"
'response.write "--------------<br>"
'response.write Hour(GetMax("Tb_Book","B_TimeEnd"," where B_Date=#"&getB_Date&"#"))&"<br />"
'response.write Hour(GetMin("Tb_Book","B_TimeEnd"," where B_Date=#"&getB_Date&"#"))&"<br />"


Sql =  "insert into Tb_Book (B_Date,B_Topic,B_RoomName,B_StartDate,B_EndDate,B_TimeStart,B_TimeEnd,B_Name,B_Tel,B_IP,B_Flag) values ('"&getB_Date&"','"&getB_Topic&"','"&getB_RoomName&"','"&getB_StartDate&"','"&getB_EndDate&"','"&getB_TimeStart&"','"&getB_TimeEnd&"','"&getB_Name&"','"&getB_Tel&"','"&Request.ServerVariables("REMOTE_ADDR")&"',True)"
'response.write Sql&"<br />"
ConActivity.execute(sql)

getB_Id = GetMax("Tb_Book","B_ID","")
Sqllog =  "insert into Tb_BookLog (B_ID,L_DateAdd,L_IP,B_Name,B_Tel,L_Method) values ('"&getB_Id&"','"&Datemmddyyyy&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&getB_Name&"','"&getB_Tel&"','Add')"
'response.write Sqllog&"<br />"
ConActivity.execute(sqllog)

	If Err.Number = 0 Then
	response.write "<script language=""javascript"">"
	response.write "alert(""Save Data Success"");"
	response.write "</script>"
	end if

end if

%>
<%
'*******************************************************
'*     ASP 101 Sample Code - http://www.asp101.com/    *
'*                                                     *
'*   This code is made available as a service to our   *
'*      visitors and is provided strictly for the      *
'*               purpose of illustration.              *
'*                                                     *
'*      http://www.asp101.com/samples/license.asp      *
'*                                                     *
'* Please direct all inquiries to webmaster@asp101.com *
'*******************************************************
%>

<%
' ***Begin Function Declaration***
' New and improved GetDaysInMonth implementation.
' Thanks to Florent Renucci for pointing out that I
' could easily use the same method I used for the
' revised GetWeekdayMonthStartsOn function.
Function GetDaysInMonth(iMonth, iYear)
	Dim dTemp
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetDaysInMonth = Day(dTemp)
End Function

' Previous implementation on GetDaysInMonth
'Function GetDaysInMonth(iMonth, iYear)
'	Select Case iMonth
'		Case 1, 3, 5, 7, 8, 10, 12
'			GetDaysInMonth = 31
'		Case 4, 6, 9, 11
'			GetDaysInMonth = 30
'		Case 2
'			If IsDate("February 29, " & iYear) Then
'				GetDaysInMonth = 29
'			Else
'				GetDaysInMonth = 28
'			End If
'	End Select
'End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
	Dim dTemp
	dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
	SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
	AddOneMonth = DateAdd("m", 1, dDate)
End Function
' ***End Function Declaration***


Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table


' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today
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
'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
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
		document.frmcalendarbooking.action="CalendarBooking.asp";
		document.frmcalendarbooking.hidFlagSave.value="Save";
		document.frmcalendarbooking.submit();	
	}
}
function goEdit(getBId)
{
	/*document.getElementById("hidBID").value = getBId;
	document.frmedit.action="EditCalendarBooking.asp"
	document.frmedit.submit();*/
	
	var obj ;
	obj = window.open("UpdateCalendarBooking.asp?BID="+getBId,"_blank","toolbar=no, scrollbars=no, resizable=no, width=400, height=150");
}
function goCancel(getBId)
{
	var obj ;
	obj = window.open("CancelCalendarBooking.asp?BID="+getBId,"_blank","toolbar=no, scrollbars=no, resizable=no, width=400, height=150");
	//obj = window.open("CancelCalendarBooking.asp?BID="+getBId,"_blank","toolbar=no, scrollbars=no, resizable=no, width=800, height=800");
			 
}
</script>
</head>

<body bgcolor="#FFFFCC">
<!-- Outer Table is simply to get the pretty border-->
<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="10%">
<TABLE BORDER=10 CELLSPACING=0 CELLPADDING=0 align="center">
<TR>
<TD>
<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=1 BGCOLOR=#99CCFF>
	<TR>
		<TD BGCOLOR=#000099 ALIGN="center" COLSPAN=7>
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 >
				<TR>
					<TD ALIGN="right"><A HREF="./calendarbooking.asp?date=<%= SubtractOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="+1">&lt;&lt;</FONT></A></TD>
					<TD ALIGN="center"><FONT COLOR=#FFFF00 size="+1"><%= MonthName(Month(dDate)) & "  " & (Year(dDate)+543) %></FONT></TD>
					<TD ALIGN="left"><A HREF="./calendarbooking.asp?date=<%= AddOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="+1">&gt;&gt;</FONT></A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">อา</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">จ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">อ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">พ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">พฤ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">ศ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="+1">ส</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=25 HEIGHT=1 BORDER=0></TD>
	</TR>
<%
' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
If iDOW <> 1 Then
	Response.Write vbTab & "<TR>" & vbCrLf
	iPosition = 1
	Do While iPosition < iDOW
		Response.Write vbTab & vbTab & "<TD>&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
End If

' Write days of month in proper day slots
iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM
	' If we're at the begginning of a row then write TR
	If iPosition = 1 Then
		Response.Write vbTab & "<TR>" & vbCrLf
	End If
	'--------------------------Code for get data from DB-------------------------------
	GDate = Month(dDate) & "/" &iCurrent  & "/" &Year(dDate)
	get_DataBook = getDataCalendarBooking(GDate)
	if get_DataBook > 0 then
		setColor="yellow"
	else
		setColor=""
	end if 
	'---------------------------------------------------------------------------------------
	' If the day we're writing is the selected day then highlight it somehow.
	If iCurrent = Day(dDate) Then
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""1"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD BGCOLOR="""&setColor&"""><A HREF=""./calendarbooking.asp?date=" &iCurrent  & "/" &  Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""+2"" color=""red"">" & iCurrent & "</FONT></A><BR></TD>" & vbCrLf
	Else
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""1"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD BGCOLOR="""&setColor&"""><A HREF=""./calendarbooking.asp?date=" &iCurrent  & "/" & Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""+1"" style=""text-decoration:none; color:#000000"">" & iCurrent & "</FONT></A><BR></TD>" & vbCrLf
	End If
	
	' If we're at the endof a row then write /TR
	If iPosition = 7 Then
		Response.Write vbTab & "</TR>" & vbCrLf
		iPosition = 0
	End If
	
	' Increment variables
	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
Loop

' Write spacer cells at end of last row if month doesn't end on a Saturday.
If iPosition <> 1 Then
	Do While iPosition <= 7
		Response.Write vbTab & vbTab & "<TD>&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
	Response.Write vbTab & "</TR>" & vbCrLf
End If
%>
</TABLE>
</TD>
</TR>
</TABLE>
</td>
<td width="90%">

<!--start form block-->
<form name="frmcalendarbooking" id="frmcalendarbooking" enctype="application/x-www-form-urlencoded" method="post">
<input type="hidden"  name="hidFlagSave" id="hidFlagSave" value=""/>
<input type="hidden" name="hidBDate" id="hidBDate" value="<%=dDate%>" />
<!--<table border="1" cellpadding="3" cellspacing="0" width="60%" align="center">
<tr>
  <td colspan="2" align="center">จองห้องประชุม กองแผนงานและวิชาการ</td>
</tr>
<tr>
<td width="40%">เรื่องที่ประชุม</td>
<td width="60%"><textarea name="txtHead" cols="50" rows="3" id="txtHead"></textarea></td>
</tr>
<tr>
  <td width="30%">เวลาประชุม</td>
  <td width="70%"><label for="BookHour">
    <select name="BookHour" id="BookHour">
      <option value="06">06</option>
      <option value="07">07</option>
      <option value="08">08</option>
      <option value="09" selected="selected">09</option>
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
      <option value="21">21</option>
      <option value="22">22</option>
    </select>
  : </label>
    <label for="BookMinute"></label>
    <select name="BookMinute" id="BookMinute">
      <option value="00">00</option>
      <option value="01">01</option>
      <option value="02">02</option>
      <option value="03">03</option>
      <option value="04">04</option>
      <option value="05">05</option>
      <option value="06">06</option>
      <option value="07">07</option>
      <option value="08">08</option>
      <option value="09">09</option>
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
      <option value="21">21</option>
      <option value="22">22</option>
      <option value="23">23</option>
      <option value="24">24</option>
      <option value="25">25</option>
      <option value="26">26</option>
      <option value="27">27</option>
      <option value="28">28</option>
      <option value="29">29</option>
      <option value="30">30</option>
      <option value="31">31</option>
      <option value="32">32</option>
      <option value="33">33</option>
      <option value="34">34</option>
      <option value="35">35</option>
      <option value="36">36</option>
      <option value="37">37</option>
      <option value="38">38</option>
      <option value="39">39</option>
      <option value="40">40</option>
      <option value="41">41</option>
      <option value="42">42</option>
      <option value="43">43</option>
      <option value="44">44</option>
      <option value="45">45</option>
      <option value="46">46</option>
      <option value="47">47</option>
      <option value="48">48</option>
      <option value="49">49</option>
      <option value="50">50</option>
      <option value="51">51</option>
      <option value="52">52</option>
      <option value="53">53</option>
      <option value="54">54</option>
      <option value="55">55</option>
      <option value="56">56</option>
      <option value="57">57</option>
      <option value="58">58</option>
      <option value="59">59</option>
      <option value="60">60</option>
    </select>
    ถึง 
    <select name="BookHourEnd" id="BookHourEnd">
      <option value="06">06</option>
      <option value="07">07</option>
      <option value="08">08</option>
      <option value="09">09</option>
      <option value="10">10</option>
      <option value="11">11</option>
      <option value="12">12</option>
      <option value="13">13</option>
      <option value="14">14</option>
      <option value="15">15</option>
      <option value="16" selected="selected">16</option>
      <option value="17">17</option>
      <option value="18">18</option>
      <option value="19">19</option>
      <option value="20">20</option>
      <option value="21">21</option>
      <option value="22">22</option>
    </select> : 
    <select name="BookMinuteEnd" id="BookMinuteEnd">
      <option value="00">00</option>
      <option value="01">01</option>
      <option value="02">02</option>
      <option value="03">03</option>
      <option value="04">04</option>
      <option value="05">05</option>
      <option value="06">06</option>
      <option value="07">07</option>
      <option value="08">08</option>
      <option value="09">09</option>
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
      <option value="21">21</option>
      <option value="22">22</option>
      <option value="23">23</option>
      <option value="24">24</option>
      <option value="25">25</option>
      <option value="26">26</option>
      <option value="27">27</option>
      <option value="28">28</option>
      <option value="29">29</option>
      <option value="30" selected="selected">30</option>
      <option value="31">31</option>
      <option value="32">32</option>
      <option value="33">33</option>
      <option value="34">34</option>
      <option value="35">35</option>
      <option value="36">36</option>
      <option value="37">37</option>
      <option value="38">38</option>
      <option value="39">39</option>
      <option value="40">40</option>
      <option value="41">41</option>
      <option value="42">42</option>
      <option value="43">43</option>
      <option value="44">44</option>
      <option value="45">45</option>
      <option value="46">46</option>
      <option value="47">47</option>
      <option value="48">48</option>
      <option value="49">49</option>
      <option value="50">50</option>
      <option value="51">51</option>
      <option value="52">52</option>
      <option value="53">53</option>
      <option value="54">54</option>
      <option value="55">55</option>
      <option value="56">56</option>
      <option value="57">57</option>
      <option value="58">58</option>
      <option value="59">59</option>
      <option value="60">60</option>
    </select>
    </td></tr>
    <tr>
    <td>ชื่อ / หน่วยงาน ที่ต้องการจอง</td>
    <td><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="50" /></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="50" /></td>
    </tr>
    <tr><td colspan="2" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onclick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="butCancle" id="butCancle" value="ล้างเนื้อหา" /></td></tr>
</table>-->
<table border="0" cellpadding="3" cellspacing="0" width="95%" align="center">
<tr>
  <td colspan="4" align="left"><b style="font-size:24px">จองห้องประชุมกองแผนงานและวิชาการ</b></td>
</tr>
<!--<tr>
<td width="30%">ห้องประชุม</td><td width="70%"><input  name="txtRoom" type="text" id="txtRoom" size="40" /></td>
</tr>-->
<tr>
<td width="15%">หัวข้อการประชุม</td>
<td width="35%"><textarea name="txtHead" cols="50" rows="3" id="txtHead"></textarea></td>
<td width="15%">ห้องประชุม / สถานที่</td>
<td width="35%"> <input type="text" name="txtRoomName" id="txtRoomName" value="ห้องประชุมกองแผนงานและวิชาการ ชั้น 4 อาคาร 5" size="40" readonly="readonly" /></td>
</tr>
    <tr>
      <td>วันที่ (วัน/เดือน/ปี)</td>
      <td><label>เริ่ม&nbsp;&nbsp;&nbsp;&nbsp;
          <select name="DayStart" size="1" id="DayStart">
        	<% For i=1 to 31%>
  			<option value="<%=i%>" <% if Day(dDate) = i then response.write "selected=""selected"" " end if %>><%=i%></option>
			<% Next %>
          </select>
&nbsp;&nbsp;&nbsp;
<select name="MonthStart" id="MonthStart">
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
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearStart" id="YearStart">
<% For q=2014 to 2020 %>
  <option value="<%=q%>" <% if Year(dDate) = q then  response.write " selected=""selected"" " end if %>><%=q+543%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;<br />
สิ้นสุด
<select name="DayEnd" size="1" id="DayEnd">
 <% For e=1 to 31%>
  <option value="<%=e%>" <% if Day(dDate) = e then response.write "selected=""selected"" " end if %>><%=e%></option>
<% Next %>
</select>
&nbsp;&nbsp;&nbsp;
<select name="MonthEnd" id="MonthEnd">
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
</select>
&nbsp;&nbsp;&nbsp;
<select name="YearEnd" id="YearEnd">
<% For qq=2014 to 2020 %>
  <option value="<%=qq%>" <% if Year(dDate) = qq then  response.write " selected=""selected"" " end if %>><%=qq+543%></option>
<% Next %>
</select>
      </label></td>
      <td>เวลา </td>
      <td><label for="StartHour">
        <select name="StartHour" id="StartHour">
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09" selected="selected">09</option>
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
          <option value="21">21</option>
          <option value="22">22</option>
        </select>
: </label>
        <label for="StartMinute"></label>
        <select name="StartMinute" id="StartMinute">
          <option value="00" selected="selected">00</option>
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
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
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
          <option value="32">32</option>
          <option value="33">33</option>
          <option value="34">34</option>
          <option value="35">35</option>
          <option value="36">36</option>
          <option value="37">37</option>
          <option value="38">38</option>
          <option value="39">39</option>
          <option value="40">40</option>
          <option value="41">41</option>
          <option value="42">42</option>
          <option value="43">43</option>
          <option value="44">44</option>
          <option value="45">45</option>
          <option value="46">46</option>
          <option value="47">47</option>
          <option value="48">48</option>
          <option value="49">49</option>
          <option value="50">50</option>
          <option value="51">51</option>
          <option value="52">52</option>
          <option value="53">53</option>
          <option value="54">54</option>
          <option value="55">55</option>
          <option value="56">56</option>
          <option value="57">57</option>
          <option value="58">58</option>
          <option value="59">59</option>
          <option value="60">60</option>
        </select>
&nbsp;
ถึง&nbsp;&nbsp;
<select name="EndHour" id="EndHour">
  <option value="06">06</option>
  <option value="07">07</option>
  <option value="08">08</option>
  <option value="09">09</option>
  <option value="10">10</option>
  <option value="11">11</option>
  <option value="12" selected="selected">12</option>
  <option value="13">13</option>
  <option value="14">14</option>
  <option value="15">15</option>
  <option value="16">16</option>
  <option value="17">17</option>
  <option value="18">18</option>
  <option value="19">19</option>
  <option value="20">20</option>
  <option value="21">21</option>
  <option value="22">22</option>
</select>
:
<select name="EndMinute" id="EndMinute">
  <option value="00" selected="selected">00</option>
  <option value="01">01</option>
  <option value="02">02</option>
  <option value="03">03</option>
  <option value="04">04</option>
  <option value="05">05</option>
  <option value="06">06</option>
  <option value="07">07</option>
  <option value="08">08</option>
  <option value="09">09</option>
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
  <option value="21">21</option>
  <option value="22">22</option>
  <option value="23">23</option>
  <option value="24">24</option>
  <option value="25">25</option>
  <option value="26">26</option>
  <option value="27">27</option>
  <option value="28">28</option>
  <option value="29">29</option>
  <option value="30">30</option>
  <option value="31">31</option>
  <option value="32">32</option>
  <option value="33">33</option>
  <option value="34">34</option>
  <option value="35">35</option>
  <option value="36">36</option>
  <option value="37">37</option>
  <option value="38">38</option>
  <option value="39">39</option>
  <option value="40">40</option>
  <option value="41">41</option>
  <option value="42">42</option>
  <option value="43">43</option>
  <option value="44">44</option>
  <option value="45">45</option>
  <option value="46">46</option>
  <option value="47">47</option>
  <option value="48">48</option>
  <option value="49">49</option>
  <option value="50">50</option>
  <option value="51">51</option>
  <option value="52">52</option>
  <option value="53">53</option>
  <option value="54">54</option>
  <option value="55">55</option>
  <option value="56">56</option>
  <option value="57">57</option>
  <option value="58">58</option>
  <option value="59">59</option>
  <option value="60">60</option>
</select></td>
    </tr>
    <tr>
    <td>ชื่อผู้รับผิดชอบ / หน่วยงาน</td>
    <td colspan="3"><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="80" /></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td colspan="3"><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="80" /></td>
    </tr>
    <tr><td colspan="4" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onclick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="butCancle" id="butCancle" value="ล้างเนื้อหา" /></td></tr>
</table>
</form>
<div align="center"><a href="index.asp">หน้าแรก</a></div>
<!--End form block-->

</td>
</tr></table>
<BR>
<%
'if isEmpty(Request.QueryString("date")) = false then
	set Rec = Server.CreateObject("ADODB.RECORDSET")
	FDate = Month(Request.QueryString("date"))&"/"&Day(Request.QueryString("date"))&"/"&Year(Request.QueryString("date"))
	'SQL = "Select * from Tb_Book where B_Date=#"&FDate&"# and B_Flag = True"
	SQL = "Select * from Tb_Book where B_Flag = True and  B_StartDate <= #"&FDate&"# and B_EndDate >= #"&FDate&"#"
	'response.write SQL&"<br />"
	Rec.open SQL,ConActivity,1,3
	if Rec.RecordCount <= 0 then
	'response.write "No Data"
	end if 
	
%>

<br />
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=4 width="100%" align="center">
<%
	while not Rec.EOF
	'response.write Rec("B_Id")
%>
<TR><TD ALIGN="center">
<table width="90%" cellpadding="3" cellspacing="0" border="0" bgcolor="#FFFF99">
<tr><td width="100%" colspan="2" align="left" bgcolor="#99FF33"><b style="font-size:24px"><%=Rec("B_Topic")%></b></td></tr>
<tr>
  <td align="left" width="70%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>ห้องประชุม : กองแผนงานและวิชาการ</strong></td>
  <td width="30%" rowspan="6" align="center">&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butEdit" id="butEdit" value="แก้ไข"  onclick="goEdit('<%=Rec("B_ID")%>')" />&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle2" id="butCancle2" value="ยกเลิก"  onclick="goCancel('<%=Rec("B_ID")%>')" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
<!--<tr><td width="80%" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;หัวข้อ :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%'=Rec("B_Topic")%></td></tr>-->
<tr>
  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;<strong>วันที่ :</strong>&nbsp;&nbsp;<span style="color:#0000FF"><%=Rec("B_StartDate")%> <% if Day(Rec("B_StartDate")) <> Day(Rec("B_EndDate"))  then %>
  <strong style="color:#000000">ถึง</strong>  <%=Rec("B_EndDate")%>
  </span>
  <% end if %>
  &nbsp;&nbsp;&nbsp;&nbsp;<strong style="color:#000000">เวลา :</strong>&nbsp;&nbsp;&nbsp;
<span style="color:#0000FF">
<%
if Minute(Rec("B_TimeStart")) < 10 then
response.write Hour(Rec("B_TimeStart"))&":0"&Minute(Rec("B_TimeStart"))
else
response.write Hour(Rec("B_TimeStart"))&":"&Minute(Rec("B_TimeStart"))
end if 
%></span> <strong style="color:#000000">&nbsp;&nbsp;ถึง&nbsp;&nbsp;</strong> 
<span style="color:#0000FF">
<%
if Minute(Rec("B_TimeEnd")) < 10 then
response.write Hour(Rec("B_TimeEnd"))&":0"&Minute(Rec("B_TimeEnd"))
else
response.write Hour(Rec("B_TimeEnd"))&":"&Minute(Rec("B_TimeEnd"))
end if
%></span></td></tr>
<tr>
  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="color:#000000">ชื่อ / หน่วยงาน ที่จอง :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF"><%=Rec("B_Name")%></span></td></tr>
<tr><td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="color:#000000">เบอร์ติดต่อ :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF"><%=Rec("B_Tel")%></span></td></tr>
<tr><td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="color:#000000">IP :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF"><%=Rec("B_IP")%></span></td></tr>
<tr><td align="left">&nbsp;</td></tr>
</table>
</TD></TR>
<%
	Rec.MoveNext
	wend
%>
</TABLE>
<%
'end if
%>

<form name="frmedit" id="frmedit" method="post" enctype="application/x-www-form-urlencoded">
<input type="hidden"  name="hidBID" id="hidBID" value=""/>
</form>
</body>
</html>

