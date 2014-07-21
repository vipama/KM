<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
dim chkShowSave
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
if isEmpty(Request.Form("hidDid")) = false then
		getHidDid = Request.Form("hidDid")
		ccount=0
		Set   getQS = Server.CreateObject("ADODB.RECORDSET")
		SQL_showQS = "Select * from Tb_Manual where D_Id='"&Request.Form("hidDid")&"' order by M_Id asc "
		getQS.open SQL_showQS,ConQS,1,3
		getCount = getQS.RecordCount
		dim arrvar(80)
 		while not getQS.EOF
		arrvar(ccount) = Request.Form(getQS("M_Code"))
		'response.write arrvar(ccount)&"/"&getQS("M_Code")&"/ <br>"
		if arrvar(ccount) = 1 then
			SQLUpdate="update tb_Manual set  M_Main="&arrvar(ccount)&" , M_Reserve=0 where D_Id='"&getHidDid&"' and M_Code='"&getQS("M_Code")&"'"
		else
			SQLUpdate="update tb_Manual set  M_Main="&arrvar(ccount)&" , M_Reserve=1 where D_Id='"&getHidDid&"' and M_Code='"&getQS("M_Code")&"'"
		end if
		ConQS.execute(SQLUpdate)
		'response.write SQLUpdate&"<br>"
		getQS.MoveNext
		wend
		
		'-----------------------------------------------------------------------------Start code for insert log to DB---------------------------------------------------------------------
		SQL_InsertLog = "insert into Tb_Log (User_Id,Method_Access,Date_Access,Department_Name) values ('"&session("member")&"','update','"&Datemmddyyyy&"','"&getDepartmentname(getHidDid)&"')"
		'response.write SQL_InsertLog
		ConQS.execute(SQL_InsertLog)
		response.write "<script language=""javascript"">"
		response.write "alert(""บันทึกข้อมูลเรียบร้อยค่ะ"");"
		response.write "</script>"
		
		'-----------------------------------------------------------------------------End code for insert log to DB----------------------------------------------------------------------
		chkShowSave = "User : "&session("member")&" ได้ทำการปรับปรุงข้อมูลของ หน่วยงาน : "&getDepartmentname(getHidDid)&"   เวลา : "&Datemmddyyyy&" <br> ข้อมูลนี้ได้ถูกเก็บเป็นประวัติการแก้ไขในฐานข้อมูลเรียบร้อยแล้วค่ะ"
end if

if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = "1"
	 end if
else
	getDid=Request.QueryString("id")
end if
'--------------------check session member--------------------
if Session("member") = "" Then
	'Response.Redirect("default1.asp")
end if
'--------------------end check session member--------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>QS</title>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		// alert(val+"/"+val1);
		window.location.href="showProcedure.asp?id="+val+"&oid="+val1;
}
</script>
</head>

<body bgcolor="#000000">
<form enctype="application/x-www-form-urlencoded" method="post" action="showProcedure.asp">
<input type="hidden" name="hidDid" value="<%=getDid%>" />
      <br />
      <div align="center"><font style="font-size:24px; color:#999999">การวิเคราะห์ Core Process และ Support Process</font></div>
      <br />
      <table width="85%" cellpadding="3" cellspacing="0" border="1" bgcolor="#FFFFCC" align="center">
      <tr><td colspan="5" align="center">
      <font style="font-size:18px">หน่วยงาน :</font>
	  <%
	  Set   rec_jobresult_group = Server.CreateObject("ADODB.RECORDSET")
	  sql_jobresult_group = "select  *  from  Tb_Department order by D_Numberlist  asc"
	  rec_jobresult_group.open sql_jobresult_group,ConQS,1,3
	  %>
	  <select name="JobresultGroupId" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:18px"  >
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
      </tr>
      <tr>
      <th width="5%" align="center">ลำดับ</th>
      <th width="10%" align="center">รหัสเอกสาร</th>
      <th width="75%" align="center">ชื่อกระบวนงาน</th>
      <th width="5%" align="center">Core</th>
      <th width="5%" align="center">Support</th>
      </tr>
      <%
	    dim showRow
		showRow=1 
	  	Set   RecshowQS = Server.CreateObject("ADODB.RECORDSET")
		SQL_showQS = "Select * from Tb_Manual where D_Id='"&getDid&"' order by M_Id asc "
		RecshowQS.open SQL_showQS,ConQS,1,3
 		while not RecshowQS.EOF
		chk_select_main = RecshowQS("M_Main")
		chk_select_reserve = RecshowQS("M_Reserve")
		chk_main=""
		chk_reserve=""
		if chk_select_main = 1 then
			chk_main = "checked=""checked"""
		else
			chk_reserve = "checked=""checked"""
		end if
	%>
    <tr><td align="center"><%=showRow%></td><td align="left">&nbsp;<%=RecshowQS("M_Code")%></td><td><%=RecshowQS("M_Name")%></td><td align="center"><input type="radio" name="<%=RecshowQS("M_Code")%>" value="1"  <%=chk_main%>  /></td><td align="center"><input type="radio" name="<%=RecshowQS("M_Code")%>" value="0" <%=chk_reserve%> /></td>
	<%
		showRow=showRow+1
  		RecshowQS.MoveNext
 	%>
 	</tr>
 	<%
	 	wend
	%>
    <tr><td colspan="5" align="center"><input type="submit"  value="บันทึก" width="250" height="60" style="font-size:24px; background-color:#006600; color:#FF0000; cursor:pointer; cursor:hand; width:150px; height:40px"/></td></tr>
  </table><br />
  <table align="center" width="85%" cellpadding="3" cellspacing="0" border="0" >
  <tr><td style="color:#999999; font-size:14px" align="center"><%=chkShowSave%></td></tr>
  </table>
  <table width="85%" cellpadding="3" cellspacing="0" border="0" bgcolor="#000000" align="center"><tr><td>
  <input type="button" value="กลับหน้าหลัก" onClick="javascript:{window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos';}"  style="font-size:24px; background-color:#006600; color:#FF0000; cursor:pointer; cursor:hand; width:150px; height:40px" />&nbsp;&nbsp;&nbsp;<input type="button" value="พิมพ์" onClick="javascript:{window.print();}"  style="font-size:24px; background-color:#006600; color:#FF0000; cursor:pointer; cursor:hand; width:150px; height:40px" />
  </td></tr></table>
  </form>
</body>
</html>
