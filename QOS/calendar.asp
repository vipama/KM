<!--#include file="admin/Config.inc.asp"-->
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
<!-- Outer Table is simply to get the pretty border-->

<TABLE BORDER=10 CELLSPACING=0 CELLPADDING=0 align="center">
<TR>
<TD>
<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=1 BGCOLOR=#99CCFF>
	<TR>
		<TD BGCOLOR=#000099 ALIGN="center" COLSPAN=7>
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 >
				<TR>
					<TD ALIGN="right"><A HREF="./calendaractivity.asp?date=<%= SubtractOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="-4">&lt;&lt;</FONT></A></TD>
					<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-4"><%= MonthName(Month(dDate)) & "  " & (Year(dDate)+543) %></FONT></TD>
					<TD ALIGN="left"><A HREF="./calendaractivity.asp?date=<%= AddOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="-4">&gt;&gt;</FONT></A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">ÍÒ</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">¨</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">Í</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">¾</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">¾Ä</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">È</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center" BGCOLOR=#0000CC><FONT COLOR=#FFFF00 size="-4">Ê</B></FONT><BR><IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
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
	
	' If the day we're writing is the selected day then highlight it somehow.
	If iCurrent = Day(dDate) Then
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD BGCOLOR=""red""><A HREF=""./calendaractivity.asp?date=" &iCurrent  & "/" &  Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
	Else
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD><A HREF=""./calendaractivity.asp?date=" &iCurrent  & "/" & Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
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

<BR>
<%
if isEmpty(Request.QueryString("date")) = false then
	set Rec = Server.CreateObject("ADODB.RECORDSET")
	FDate = Month(Request.QueryString("date"))&"/"&Day(Request.QueryString("date"))&"/"&Year(Request.QueryString("date"))
	SQL = "Select * from Tb_Book where B_Meeting=#"&FDate&"#"
	'response.write SQL&"<br />"
	Rec.open SQL,ConActivity,1,3
	if Rec.RecordCount <= 0 then
	'response.write "No Data"
	end if 
	
%>
<br />
<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=0 width="100%" align="center">
<%
	while not Rec.EOF
	'response.write Rec("B_Id")
%>
<TR><TD ALIGN="center">
<%=Rec("B_Id")%>
<!--<FORM ACTION="./calendar.asp" METHOD=GET>
<SELECT NAME="month">
	<OPTION VALUE=1>January</OPTION>
	<OPTION VALUE=2>February</OPTION>
	<OPTION VALUE=3>March</OPTION>
	<OPTION VALUE=4>April</OPTION>
	<OPTION VALUE=5>May</OPTION>
	<OPTION VALUE=6>June</OPTION>
	<OPTION VALUE=7>July</OPTION>
	<OPTION VALUE=8>August</OPTION>
	<OPTION VALUE=9>September</OPTION>
	<OPTION VALUE=10>October</OPTION>
	<OPTION VALUE=11>November</OPTION>
	<OPTION VALUE=12>December</OPTION>
</SELECT>
<SELECT NAME="day">
	<OPTION VALUE=1>1</OPTION>
	<OPTION VALUE=2>2</OPTION>
	<OPTION VALUE=3>3</OPTION>
	<OPTION VALUE=4>4</OPTION>
	<OPTION VALUE=5>5</OPTION>
	<OPTION VALUE=6>6</OPTION>
	<OPTION VALUE=7>7</OPTION>
	<OPTION VALUE=8>8</OPTION>
	<OPTION VALUE=9>9</OPTION>
	<OPTION VALUE=10>10</OPTION>
	<OPTION VALUE=11>11</OPTION>
	<OPTION VALUE=12>12</OPTION>
	<OPTION VALUE=13>13</OPTION>
	<OPTION VALUE=14>14</OPTION>
	<OPTION VALUE=15>15</OPTION>
	<OPTION VALUE=16>16</OPTION>
	<OPTION VALUE=17>17</OPTION>
	<OPTION VALUE=18>18</OPTION>
	<OPTION VALUE=19>19</OPTION>
	<OPTION VALUE=20>20</OPTION>
	<OPTION VALUE=21>21</OPTION>
	<OPTION VALUE=22>22</OPTION>
	<OPTION VALUE=23>23</OPTION>
	<OPTION VALUE=24>24</OPTION>
	<OPTION VALUE=25>25</OPTION>
	<OPTION VALUE=26>26</OPTION>
	<OPTION VALUE=27>27</OPTION>
	<OPTION VALUE=28>28</OPTION>
	<OPTION VALUE=29>29</OPTION>
	<OPTION VALUE=30>30</OPTION>
	<OPTION VALUE=31>31</OPTION>
</SELECT>
<SELECT NAME="year">
	<OPTION VALUE=1990>1990</OPTION>
	<OPTION VALUE=1991>1991</OPTION>
	<OPTION VALUE=1992>1992</OPTION>
	<OPTION VALUE=1993>1993</OPTION>
	<OPTION VALUE=1994>1994</OPTION>
	<OPTION VALUE=1995>1995</OPTION>
	<OPTION VALUE=1996>1996</OPTION>
	<OPTION VALUE=1997>1997</OPTION>
	<OPTION VALUE=1998>1998</OPTION>
	<OPTION VALUE=1999>1999</OPTION>
	<OPTION VALUE=2000 SELECTED>2000</OPTION>
	<OPTION VALUE=2001>2001</OPTION>
	<OPTION VALUE=2002>2002</OPTION>
	<OPTION VALUE=2003>2003</OPTION>
	<OPTION VALUE=2004>2004</OPTION>
	<OPTION VALUE=2005>2005</OPTION>
	<OPTION VALUE=2006>2006</OPTION>
	<OPTION VALUE=2007>2007</OPTION>
	<OPTION VALUE=2008>2008</OPTION>
	<OPTION VALUE=2009>2009</OPTION>
	<OPTION VALUE=2010>2010</OPTION>
</SELECT>
<BR>
<INPUT TYPE="submit" VALUE="Show This Date on the Calendar!">
</FORM>-->
</TD></TR>
<%
	Rec.MoveNext
	wend
%>
</TABLE>
<%
end if
%>
