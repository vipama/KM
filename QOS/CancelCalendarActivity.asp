<!--#include file="../../Config.inc.asp"-->
<%
dim getSave
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
getSave = Request.Form("hidFlagSave")
if getSave = "Save" then

'getB_Date = Month(Request.Form("hidBDate"))&"/"&day(Request.Form("hidBDate"))&"/"&year(Request.Form("hidBDate"))
getA_Id = Request.Form("Aid")
getA_Name = Request.Form("txtSubscribers")
getA_Tel  = Request.Form("txtSubscribersTel")

	Sql =  "update  Tb_Activity set  A_Flag=False , A_CancleName='"&getA_Name&"' , A_CancleTel='"&getA_Tel&"' , A_IPCancle='"&Request.ServerVariables("REMOTE_ADDR")&"'   where A_ID="&getA_Id&" "
	'response.write Sql&"<br />"
	ConActivity.execute(sql)
	
Sqllog =  "insert into Tb_ActivityLog (A_ID,L_DateAdd,L_IP,A_Name,A_Tel,A_Method) values ('"&getA_Id&"','"&Datemmddyyyy&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&getA_Name&"','"&getA_Tel&"','Cancel')"
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
<title>ปฎิทินกิจกรรม</title>
<script language="javascript" src="jScript/JS.js"></script>
<script language="javascript">
function goSave()
{
	if (alltrim(document.getElementById("txtSubscribers").value).length == 0 )
	{
		alert("กรุณาตรวจสอบ ชื่อผู้ยกเลิกด้วย");
	}
	else if(isNumber(alltrim(document.getElementById("txtSubscribersTel").value)) == false)
	{
		alert("กรุณากรอกเบอร์โทรศัพท์อีกครั้ง");
	}
	else
	{
		document.frmcalendaractivity.action="CancelCalendarActivity.asp";
		document.frmcalendaractivity.hidFlagSave.value="Save";
		document.frmcalendaractivity.submit();	
	}
}
</script>
</head>

<body leftmargin="0" topmargin="0">
<!-- Outer Table is simply to get the pretty border-->
<BR>
<%
'if isEmpty(Request.QueryString("BID") = false then
	getAID = Request.QueryString("AID")
%>
<form name="frmcalendaractivity" id="frmcalendaractivity" enctype="application/x-www-form-urlencoded" method="post">
<input type="hidden"  name="hidFlagSave" id="hidFlagSave" value=""/>
<input type="hidden" name="Aid" id="Aid" value="<%=getAID%>" />
<table border="1" cellpadding="3" cellspacing="0" width="100%" align="center">
    <td width="50%">ชื่อผู้ยกเลิก</td>
    <td width="50%"><label for="txtSubscript"></label>
      <input name="txtSubscribers" type="text" id="txtSubscribers" size="40"  value=""/></td>
    </tr>
    <tr>
    <td>เบอร์โทรติดต่อ</td>
    <td><label for="txtSubscriptTel"></label>
      <input name="txtSubscribersTel" type="text" id="txtSubscribersTel" size="40" value="" /></td>
    </tr>
    <tr><td colspan="2" align="center"><input type="button" name="butActivity" id="butActivity" value="บันทึก" onClick="goSave()" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="butCancle" id="butCancle" value="ยกเลิก" onClick="javascript:{ window.close(); }" /></td></tr>
</table>
</form>
<br />
<%
'end if
%>
</body>
</html>

