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

getB_Topic = Request.Form("txtHead")
getB_Meeting = Month(getB_Date)&"/"&day(getB_Date)&"/"&year(getB_Date)
getB_TimeStart = Request.Form("BookHour")&":"&Request.Form("BookMinute") 
getB_TimeEnd = Request.Form("BookHourEnd")&":"&Request.Form("BookMinuteEnd")
getB_Name = Request.Form("txtSubscribers")
getB_Tel  = Request.Form("txtSubscribersTel")

	Sql =  "update  Tb_Book set  B_Flag=False , B_CancleName='"&getB_Name&"' , B_CancleTel='"&getB_Tel&"' , B_IPCancle='"&Request.ServerVariables("REMOTE_ADDR")&"'   where B_ID="&getB_Id&" "
	'response.write Sql&"<br />"
	ConActivity.execute(sql)
	
Sqllog =  "insert into Tb_BookLog (B_ID,L_DateAdd,L_IP,B_Name,B_Tel,L_Method) values ('"&getB_Id&"','"&Datemmddyyyy&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&getB_Name&"','"&getB_Tel&"','Cancel')"
	'response.write Sqllog&"<br />"
	ConActivity.execute(sqllog)
	
		If Err.Number = 0 Then
		response.write "<script language=""javascript"">"
		response.write "alert(""Save Data Success"");"
		response.write "window.opener.location.reload();"
		response.write "window.close();"
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
		alert("กรุณาตรวจสอบ ชื่อผู้แก้ไขด้วย");
	}
	else if(isNumber(alltrim(document.getElementById("txtSubscribersTel").value)) == false)
	{
		alert("กรุณากรอกเบอร์โทรศัพท์อีกครั้ง");
	}
	else
	{
		/*document.frmcalendarbooking.action="CancelCalendarBooking.asp";
		document.frmcalendarbooking.hidFlagSave.value="Save";
		document.frmcalendarbooking.submit();*/
		document.frmcalendarbooking.action="EditCalendarBooking.asp";
		window.resizeTo(1000,800);
		document.frmcalendarbooking.submit();
			
	}
}
</script>
</head>

<body leftmargin="0" topmargin="0">
<!-- Outer Table is simply to get the pretty border-->
<BR>
<%
'if isEmpty(Request.QueryString("BID") = false then
	getBID = Request.QueryString("BID")
%>
<form name="frmcalendarbooking" id="frmcalendarbooking" enctype="application/x-www-form-urlencoded" method="post">
<input type="hidden"  name="hidFlagSave" id="hidFlagSave" value=""/>
<input type="hidden" name="hidBID" id="hidBID" value="<%=getBID%>" />
<table border="1" cellpadding="3" cellspacing="0" width="100%" align="center">
    <td width="50%">ชื่อผู้แก้ไข</td>
    <td width="50%"><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="40"  value=""/></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="40" value="" /></td>
    </tr>
    <tr><td colspan="2" align="center"><input type="button" name="butBooking" id="butBooking" value="บันทึก" onclick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle" id="butCancle" value="ยกเลิก" onclick="javascript:{ window.close(); }" /></td></tr>
</table>
</form>
<br />
<%
'end if
%>
</body>
</html>

