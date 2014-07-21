<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
' # start code for check permission in DB 
if isEmpty(session("member")) = True then
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
else 
 	if Session("member") <> getPermission(session("member"),"L_Email") or isnull(session("member")) = true or session("member") = "" then
		Response.write "<script>"
		Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
		Response.write " 	window.location.href=""default.asp""; "
		Response.write "</script>"
		'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
	else
		session("Depart") = getPermission(session("member"),"D_Id")
	end if
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
<title>รายงานการทบทวนโดยฝ่ายบริหาร</title>
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
<div style="font-size:18px; font-weight:bold" align="center">รายงานทบทวนโดยฝ่ายบริหาร</div>
<br />
<form name="frmManagementReview" id="ManageMentReview" enctype="application/x-www-form-urlencoded" >
<input type="hidden" name="hidSave" id="hidSave" value="" />
<input type="hidden" name="hidMRID" id="hidMRID" value="" />
<input type="hidden" name="hidDid" id="hidDid" value="<%=getDid%>" />
<input type="hidden" name="hidOid" id="hidOid" value="<%=getOid%>" />
<table width="100%" align="center" cellpadding="2" cellspacing="3">
  <tr>
    <td width="25%" class="text">รายการประชุมทบทวนโดยฝ่ายบริหาร</td>
    <td width="75%" class="text"><label>
        <input type="radio" name="radio" id="radioDepart" value="1" <% if getOid = "1" then response.write "checked=""checked"" " end if %>  onclick="ChangeJobresultGroupManagementReviewReport('01',this.value)" />
        ระดับกรม
    </label>      <label>
      
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="radio" name="radio"  id="radioSubDepart" value="2" <% if getOid = "2" then response.write "checked=""checked"" " end if %>  onclick="ChangeJobresultGroupManagementReviewReport(1,this.value)" />
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
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroupManagementReviewReport(this.value,<%=getOid%>)" style="font-size:14px" class="text"   >
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
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroupManagementReviewReport(this.value,1)"  style="font-size:14px" class="text"   >
			  <option value="01"  <% if getDid = "01" then response.write " selected=""selected"" " end if %> >คณะกรรมการบริหารระบบคุณภาพ</option>
              <option value="02"  <% if getDid = "02" then response.write " selected=""selected"" " end if %>>คณะกรรมการประสานงานระบบคุณภาพ</option>
			  </select>    </td>
  </tr>
  <% end if %>
</table>
</form>
<div align="center"><a href="/kmfda/_block/qos" target="_self" style="text-decoration:none; color:#000000"><b>หน้าแรก</b></a></div><br />
<table width="75%" cellpadding="3" cellspacing="0" border="1" align="center" bordercolor="#333333">
<tr>
  <td colspan="2">
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
  <td width="90%" align="center" class="text">รายละเอียด</td>
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
<td align="center"><label>
  <input name="butView" type="button" class="textsmall" id="butView" value="ดูรายงาน" onClick="ManagementReview_goViewDoc('<%=RecGet("MR_ID")%>','<%=getDid%>','report')" />
</label>  <label></label><label></label></td>
</tr>
<%
RecGet.MoveNext
Wend
if RecGet.RecordCount = 0 then
%>
<tr><td colspan="2" align="center"><b>No Data</b></td></tr>
<% end if %>
</table>
</body>
</html>
