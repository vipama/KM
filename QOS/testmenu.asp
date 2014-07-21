<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>

<%
'How many pixels high we want our bar graph
Const graphHeight = 300
Const graphWidth = 450
Const barImage = "images/core.jpg"


sub BarChart(data, labels, title, axislabel)
	'Print heading
	Response.Write("<TABLE CELLSPACING=0 CELLPADDING=1 BORDER=1 WIDTH=" & graphWidth & ">" & chr(13))
	Response.Write("<TR><TH COLSPAN=" & UBound(data) - LBound(data) + 2 & ">")
	Response.Write("<FONT SIZE=+2>" & title & "</FONT></TH></TR>" & chr(13))
	Response.Write("<TR><TD VALIGN=TOP ALIGN=RIGHT>" & chr(13))

	'Find the highest value
	Dim hi
	hi = data(LBound(data))

	Dim i
	for i = LBound(data) to UBound(data)
		if data(i) > hi then hi = data(i)
	next

	'Print out the highest value at the top of the chart
	Response.Write(hi & "</TD>")

	Dim widthpercent
	widthpercent = CInt((1 / (UBound(data) - LBound(data) + 1)) * 100)

	For i = LBound(data) to UBound(data)
		Response.Write(" <TD VALIGN=BOTTOM ROWSPAN=2 WIDTH=" & widthpercent & "% >" & chr(13))
		Response.Write("   <IMG SRC=""" & barImage & """ WIDTH=100% HEIGHT=" & CInt(data(i)/hi * graphHeight) & ">" & chr(13))
		Response.Write(" </TD>" & chr(13))
	Next

	Response.Write("</TR>")
	Response.Write("<TR><TD VALIGN=BOTTOM ALIGN=RIGHT>0</TD></TR>")

	'Write footer
	Response.Write("<TR><TD ALIGN=RIGHT VALIGN=BOTTOM>" & axislabel & "</TD>" & chr(13))
	for i = LBound(labels) to UBound(labels)
		Response.Write("<TD VALIGN=BOTTOM ALIGN=CENTER>" & labels(i) & "</TD>" & chr(13))
	next
	Response.Write("</TR>" & chr(13))
	Response.Write("</TABLE>")
end sub

Dim dataArray(10)
dataArray(0) = 8
dataArray(1) = 10
dataArray(2) = 8
dataArray(3) = 14
dataArray(4) = 6
dataArray(5) = 13
dataArray(6) = 7
dataArray(7) = 11
dataArray(8) = 8
dataArray(9) = 9
dataArray(10) = 11

Dim labelArray(10)
labelArray(0) = "3/2"
labelArray(1) = "3/3"
labelArray(2) = "3/4"
labelArray(3) = "3/5"
labelArray(4) = "3/6"
labelArray(5) = "3/7"
labelArray(6) = "3/8"
labelArray(7) = "3/9"
labelArray(8) = "3/10"
labelArray(9) = "3/11"
labelArray(10) = "3/12"

%>

<HTML>
<BODY>
<CENTER>
<% BarChart dataArray,labelArray,"Telephone Sales","Date" %>
</CENTER>
</BODY>
</HTML>