<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->

<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
dim setComport
dim setUncomport
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
Datemmddyyyy1=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
dim getRID
dim chkPS,chkPC,chkQ,chkW,FlagMain_Reserve,getMCode,getTid,getDid,getMCode1,getTid1,getDid1
'--------------------------------------------------------start block for sava data---------------------------------------------------
getSave = Request.Form("hidS")
if getSave <> "" then
get_txtReviewNumber = Request.Form("txtReviewNumber")
get_radioReviewType = Request.Form("radioReviewType")
get_DepartID = Request.Form("DepartID")
get_Manual = Request.Form("Manual")
getRID = get_txtReviewNumber
get_txtName = Request.Form("txtName")
get_txtPosition = Request.Form("txtPosition")
get_radioPerfect = Request.Form("radioPerfect")
get_chkCurrent = Request.Form("chkCurrent")
get_chkSupportWork = Request.Form("chkSupportWork")
get_chkBelongManual = Request.Form("chkBelongManual")
get_chkElse = Request.Form("chkElse")
get_txtElse = Request.Form("txtElse")
get_radioRemake = Request.Form("radioRemake")
get_RemakefinishDay = Request.Form("RemakefinishDay")
get_RemakefinishMonth = Request.Form("Remakefinishmonth")
get_RemakefinishYear = Request.Form("RemakefinishYear")

get_Editfinishday = Request.Form("EditfinishDay")
get_EditfinishMonth = Request.Form("Editfinishmonth")
get_EditfinishYear = Request.Form("EditfinishYear")
get_chkNotNow = Request.Form("chkNotNow")
get_chkNotSupportWork = Request.Form("chkNotSupportWork")
get_chkNewWayWork = Request.Form("chkNewWayWork")
get_chkElse2 = Request.Form("chkElse2")
get_txtElse2 = Request.Form("txtElse2")

RemakeFinishDate = get_RemakefinishMonth&"/"&get_Remakefinishday&"/"&get_RemakefinishYear
EditFinishDate = get_EditfinishMonth&"/"&get_Editfinishday&"/"&get_EditfinishYear

if get_radioReviewType = "PS" or get_radioReviewType = "PC" then
get_Manual_Name =GetSingleFieldQS("Tb_Manual","M_Name"," where M_Code='"&get_Manual&"'")
get_Manual_Code =GetSingleFieldQS("Tb_Manual","M_Code"," where M_Code='"&get_Manual&"'")
elseif get_radioReviewType = "Q" then
get_Manual_Name =GetSingleFieldQS("Tb_QM","QM_Name"," where QM_Code='"&get_Manual&"'")
get_Manual_Code =GetSingleFieldQS("Tb_QM","QM_Code"," where QM_Code='"&get_Manual&"'")
elseif get_radioReviewType = "W" then
get_Manual_Name =GetSingleFieldQS("Tb_Workin","W_Name"," where W_Code='"&get_Manual&"'")
get_Manual_Code =GetSingleFieldQS("Tb_Workin","W_Code"," where W_Code='"&get_Manual&"'")
end if
'response.write "<br>"&get_Manual_Code&"<br>"
if get_radioPerfect = "1"  then
	setComport = True
	setUncomport = False
	RemakeFinishDate = Datemmddyyyy
	EditFinishDate = Datemmddyyyy
	get_radioRemake = 0
	get_chkElse2 = False
	if isEmpty(get_chkElse) = true then
		get_chkElse = False
	end if
else
	setComport = False
	setUncomport = True
	get_chkElse = False
	if isEmpty(get_chkElse2) = true then
		get_chkElse2 = False
	end if	
end if

	dim chkRepeat
	dim SQL_LOG 
	chkRepeat = GetCountRowQS("Tb_Review","M_Code"," where M_Code='"&get_Manual_Code&"' and D_Id='"&get_DepartID&"'") 'ส่วนสำหรับตรวจสอบว่าในฐานมีข้อมูลอันนี้อยู่หรือไม่
	if chkRepeat <> 0 then
	SQL_ADD = " Update Tb_Review set Type_Sop='"&get_radioReviewType&"' ,CurrentReviewDate='"&Datemmddyyyy&"' , D_Id='"&get_DepartID&"' , M_Code='"&get_Manual_Code&"' , M_Name='"&get_Manual_Name&"' , Comport="&setComport&" , Logic_Comport1='"&get_chkCurrent&"' , Logic_Comport2='"&get_chkSupportWork&"' , Logic_Comport3='"&get_chkBelongManual&"' , Logic_Comport4="&get_chkElse&" , Logic_Comport5='"&get_txtElse&"' , Uncomport="&setUncomport&" , MethodType="&get_radioRemake&" , Remake_Date='"&RemakeFinishDate&"' , Edit_Date='"&EditFinishDate&"' , Logic_Uncomport1='"&get_chkNotNow&"' , Logic_Uncomport2='"&get_chkNotSupportWork&"' , Logic_Uncomport3='"&get_chkNewWayWork&"' , Logic_Uncomport4="&get_chkElse2&" , Logic_Uncomport5='"&get_txtElse2&"' , Name_Review='"&get_txtName&"' , Level_Review='"&get_txtPosition&"'  where D_Id='"&get_DepartID&"'  and M_Code='"&get_Manual_Code&"' "
	'response.write SQL_ADD
	getRID = GetSingleFieldQS("Tb_Review","No_Review","where M_Code='"&get_Manual_Code&"' and D_Id='"&get_DepartID&"'")
	ConQS.execute(SQL_ADD)
	
	'---------------------------------------------------------------------Start Block for Add to Log table------------------------------------------------------------------------------
	SQL_LOG = "insert into Tb_LogReview (User_Id,Method_Access,Date_Access,Department_Name,M_Code) values ('"&session("member")&"','Update','"&Datemmddyyyy1&"','"&getDepartmentname(get_DepartID)&"','"&get_Manual_Code&"')"
	ConQS.execute(SQL_LOG)
	'---------------------------------------------------------------------End Block for Add to Log table---------------------------------------------------------------------------------
	
	response.write "<script language=""javascript"">"
	response.write "alert(""ปรับปรุงข้อมูลเรียบร้อยค่ะ"");"
	response.write "</script>"
	getSave=""
	else
	SQL_ADD = "insert into Tb_Review (No_Review,Type_Sop,CurrentReviewDate,D_Id,M_Code,M_Name,Comport,Logic_Comport1,Logic_Comport2,Logic_Comport3,Logic_Comport4,Logic_Comport5,Uncomport,MethodType,Remake_Date,Edit_Date,Logic_Uncomport1,Logic_Uncomport2,Logic_Uncomport3,Logic_Uncomport4,Logic_Uncomport5,Name_Review,Level_Review) values ('"&get_txtReviewNumber&"','"&get_radioReviewType&"','"&Datemmddyyyy&"','"&get_DepartID&"','"&get_Manual_Code&"','"&get_Manual_Name&"',"&setComport&",'"&get_chkCurrent&"','"&get_chkSupportWork&"','"&get_chkBelongManual&"',"&get_chkElse&",'"&get_txtElse&"',"&setUncomport&","&get_radioRemake&",'"&RemakeFinishDate&"','"&EditFinishDate&"','"&get_chkNotNow&"','"&get_chkNotSupportWork&"','"&get_chkNewWayWork&"',"&get_chkElse2&",'"&get_txtElse2&"','"&get_txtName&"','"&get_txtPosition&"')"
	'response.write SQL_ADD
	ConQS.execute(SQL_ADD)
	
	'---------------------------------------------------------------------Start Block for Add to Log table------------------------------------------------------------------------------
	SQL_LOG = "insert into Tb_LogReview (User_Id,Method_Access,Date_Access,Department_Name,M_Code) values ('"&session("member")&"','Insert','"&Datemmddyyyy1&"','"&getDepartmentname(get_DepartID)&"','"&get_Manual_Code&"')"
	ConQS.execute(SQL_LOG)
	'---------------------------------------------------------------------End Block for Add to Log table---------------------------------------------------------------------------------
	
	response.write "<script language=""javascript"">"
	response.write "alert(""บันทึกข้อมูลเรียบร้อยค่ะ"");"
	response.write "</script>"
	getSave=""
	end if
	getDid1=get_DepartID
	getMCode1=get_Manual_Code
	getTid1=get_radioReviewType
	chkShowSave = "User : "&session("member")&" ได้ทำการปรับปรุงข้อมูลของ <BR /> กระบวนงาน รหัส : "&getMCode1&"  หน่วยงาน : "&getDepartmentname(getDid1)&"   เวลา : "&Datemmddyyyy&" <br> ข้อมูลนี้ได้ถูกเก็บเป็นประวัติการแก้ไขในฐานข้อมูลเรียบร้อยแล้วค่ะ"
end if
'------------------------------------------------------------------------------------------------end block for sava data------------------------------------------------------------------------------------------------------------


 chkPS=""
 chkPC=""
 chkQ=""
 chkW=""
 FlagMain_Reserve=""
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = getDid1
	 end if
else
	getDid=Request.QueryString("id")
end if

if isEmpty(Request.QueryString("MC")) = false then
getMCode = Request.QueryString("MC")
else
getMcode = getMcode1
end if


if isEmpty(Request.QueryString("tid")) <> true then
	if Request.QueryString("tid") = "PC" then
		getTid = "PC"
		FlagMain_Reserve = "M_Main=1"
		chkPC = "checked=""checked"""
	elseif Request.QueryString("tid") = "PS" then
		getTid = "PS"
		FlagMain_Reserve = "M_Reserve=1"
		chkPS= "checked=""checked"""
	end if
	if Request.QueryString("tid") = "W" then
		getTid = "W"
		chkW= "checked=""checked"""
	elseif  Request.QueryString("tid") = "Q" then
		getTid = "Q"
		chkQ= "checked=""checked"""
	end if
else
	getTid = getTid1
	if getTid = "PC" then
		FlagMain_Reserve = "M_Main=1"
		chkPC = "checked=""checked"""
	elseif getTid = "PS" then
		FlagMain_Reserve = "M_Reserve=1"
		chkPS= "checked=""checked"""
	elseif getTid = "Q" then
		chkQ= "checked=""checked"""
	elseif getTid = "W" then
		chkW= "checked=""checked"""
	end if
end if
'response.write FlagMain_Reserve

'---------------------------------------------------------------check parameter for query data----------------------------------------------------------------------------------------
if isEmpty(getDid) = true and isEmpty(getMCode) = true  and  isEmpty(getTid)  = true then
	Response.Redirect("FReview.asp")
end if
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------Start code get data from DB-----------------------------------------------------------------
'response.write "<br>==="&getMCode
set recRev = Server.CreateObject("ADODB.RECORDSET")
'sql_rev = "select * from Tb_Review where D_id='"&getDid&"' and  M_Code='"&getMCode&"' and Type_Sop='"&getTid&"' "
sql_rev = "select * from Tb_Review where M_Code='"&getMCode&"'"
'response.write "<br>"&sql_rev
recRev.open sql_rev,ConQS,1,3
if recRev.RecordCount <= 0 then
	Response.write "<script language='javascript'>"
	response.Write "alert('No data ! \r\n Please try again!');"
	Response.write "window.location.href='FReview.asp' "
	Response.write "</script>"
	'Response.Redirect("FReview.asp")
end if
while not recRev.EOF
getR_Id = recRev("R_Id")
getNo_Review = recRev("No_Review")
getType_Sop = recRev("Type_Sop")
getCurrentReviewDate = recRev("CurrentReviewDate")
getD_Id = recRev("D_Id")
getM_Code = recRev("M_Code")
getM_Name = recRev("M_Name")
getComport = recRev("Comport")
getLogic_comport1 = recRev("Logic_Comport1")
getLogic_comport2 = recRev("Logic_Comport2")
getLogic_comport3 = recRev("Logic_Comport3")
getLogic_comport4 = recRev("Logic_Comport4")
getLogic_comport5 = recRev("Logic_Comport5")
getUncomport = recRev("Uncomport")
getMethod_Type = recRev("MethodType")
getRemake_Date = recRev("Remake_Date")
getEdit_Date = recRev("Edit_Date")
getLogic_Uncomport1 = recRev("Logic_Uncomport1")
getLogic_Uncomport2 = recRev("Logic_Uncomport2")
getLogic_Uncomport3 = recRev("Logic_Uncomport3")
getLogic_Uncomport4 = recRev("Logic_Uncomport4")
getLogic_Uncomport5 = recRev("Logic_Uncomport5")
getName_Review = recRev("Name_Review")
getLevel_Review = recRev("Level_Review")
recRev.MoveNext
Wend
if getType_Sop = "PC" then
		getTid = "PC"
		FlagMain_Reserve = "M_Main=1"
		chkPC = "checked=""checked"""
elseif getType_Sop = "PS" then
		getTid = "PS"
		FlagMain_Reserve = "M_Reserve=1"
		chkPS= "checked=""checked"""
elseif getType_Sop = "W" then
		getTid = "W"
		chkW= "checked=""checked"""
elseif getType_Sop = "Q" then
		getTid = "Q"
		chkQ= "checked=""checked"""
end if
'---------------------------------------------------------------End code get data from DB------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>ทบทวนกระบวนงาน</title>
<style type="text/css">
<!--
.style1 {
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			var typeID = getRadioValue("radioReviewType");
			window.location.href="FReview.asp?id="+val+"&oid="+val1+"&tid="+typeID;
		}else{
			var typeID = getRadioValue("radioReviewType");
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="FReview.asp?id="+strUser+"&oid="+val1+"&tid="+typeID;
		}
		
}
function goSave()
{
		if ((document.frmFReview.txtName.value != "")&&(document.frmFReview.txtPosition.value != "") )
		{
			document.frmFReview.hidS.value="S"
			document.frmFReview.submit();	
		}else{
			alert("กรุณากรอกชื่อและตำแหน่งด้วยค่ะ");
			document.frmFReview.txtName.focus();
		}
}
</script>
<script type="text/javascript" src="JScript/JS.js"></script>
</head>
<%
	dim runNum
	Set   rec_get_Id = Server.CreateObject("ADODB.RECORDSET")
	  sql_get_Id = "select  top 1 *  from  Tb_Review order by R_Id  desc"
	  rec_get_Id.open sql_get_Id,ConQS,1,3
	   while not rec_get_Id.EOF
	   	runNum = rec_get_Id("R_Id")
	   rec_get_Id.MoveNext
	   wend
	   
	  if runNum = "" or runNum =0 then
	  	runNum = 1
	  else
	  	runNum = runNum+1
	  end if
	  if isEmpty(getRID) = false then
	  response.write "<div align=""center"">"&chkShowSave&"</div><br />"
	  end if
%>
<body onLoad="<% if getComport = true and getUncomport = false then response.write " setPageEnable('2');" else response.write " setPageEnable('1');" end if %>">
<form name="frmFReview" method="post" enctype="application/x-www-form-urlencoded" action="FReview_update.asp">
<input type="hidden"  name="hidS" id="hidS" value=""/>
<input type="hidden" name="DepartID" id="DepartID"  value="<%=getD_Id%>"/>
<input type="hidden" name="Manual" id="Manual" value="<%=getM_Code%>" />
<input type="hidden" name="radioReviewType" value="<%=getType_Sop%>" />
<%
if isEmpty(getRID) = False then
response.write "<input type=""hidden"" value="""&getRID&""" name=""hidRID"" />"
end if
%>
<div align="center"  style="font-size:24px;"><strong>แบบทบทวนกระบวนงาน</strong></div><br />
<table width="85%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr><th align="right">No Review : <input type="text"  name="txtReviewNumber" id="txtReviewNumber" readonly value="<%=getNo_Review%>" /></th></tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>ดำเนินการทบทวน</td>
        <td><input type="radio" name="radioReviewTypeShow" id="radioReviewType1" value="Q" <%=chkQ%>onclick="ChangeJobresultGroup('','')"  disabled="disabled"  />
          <label>คู่มือคุณภาพ</label>
&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewTypeShow" id="radioReviewType2" value="PC"  <%=chkPC%> onClick="ChangeJobresultGroup('','')" disabled="disabled"   />
          <label >คู่มือขั้นตอนการปฏิบัติงาน (P) (Core Process)</label></td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewTypeShow" id="radioReviewType3" value="W" <%=chkW%> onClick="ChangeJobresultGroup('','')" disabled="disabled" />
          <label >คู่มือขั้นวิธีการปฏิบัติงาน (W)</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewTypeShow" id="radioReviewType4" value="PS" <%=chkPS%>  onclick="ChangeJobresultGroup('','')" disabled="disabled" />
          <label >คู่มือขั้นตอนการปฏิบัติงาน (P) (Support Process)</label></td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="15%">หน่วยงาน :</td>
        <td width="85%">
             <%
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
			  <select name="DepartIDShow" id="DepartIDShow" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:14px"  disabled="disabled"  >
			  <%
			  while not recDepart.EOF
			  if recDepart("D_Id") = getD_Id then
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
			  </select> 
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="15%">กระบวนงาน :</td>
        <td width="85%">
        <%
	  Set   recSOP = Server.CreateObject("ADODB.RECORDSET")
	  if getType_Sop = "PC" or getType_Sop = "PS" then
	  	sqlSOP = "select  *  from  Tb_Manual where  D_Id='"&getD_Id&"' and "&FlagMain_Reserve&" order by M_Id  asc"
	  elseif getType_Sop = "W" then
	   sqlSOP = "select  *  from  Tb_Workin where  D_Id='"&getD_Id&"' order by W_Id  asc"
	  elseif getType_Sop = "Q" then
	  	sqlSOP = "select  *  from  Tb_QM where  D_Id='"&getD_Id&"' order by QM_Id  asc"
	  else
	  response.write "Error ! Type not match"
	  end if
	  'response.write sqlSOP 
	  recSOP.open sqlSOP,ConQS,1,3
	  %>
	  <select name="ManualShow" id="ManualShow" style="font-size:14px" disabled="disabled"  >
	  <%
	  while not recSOP.EOF
	 if  getType_Sop = "PS" or getType_Sop = "PC" then 
		 if recSOP("M_Id") = GetSingleFieldQS("Tb_Manual","M_Id","where M_Code='"&getM_Code&"' ") then
				selected = "selected=""selected"""
		  else
				selected = ""
		  end if
	 elseif  getType_Sop = "Q" then
	 	 if recSOP("QM_Id") = GetSingleFieldQS("Tb_QM","QM_Id","where QM_Code='"&getM_Code&"' ") then
				selected = "selected=""selected"""
		  else
				selected = ""
		  end if
	 elseif  getType_Sop = "W" then
	 		 if recSOP("W_Id") = GetSingleFieldQS("Tb_Workin","W_Id","where W_Code='"&getM_Code&"' ") then
				selected = "selected=""selected"""
		  else
				selected = ""
		  end if
	 end if
	 
	 if  getType_Sop = "PS" or getType_Sop = "PC" then 
	  %>
	  <option value="<%=recSOP("M_Id")%>" <%=selected%> ><%response.write recSOP("M_Code")&" "&recSOP("M_Name")%></option>
	  <%
	  elseif  getType_Sop = "Q" then
	  %>
	  <option value="<%=recSOP("QM_Id")%>" <%=selected%> ><%response.write recSOP("QM_Code")&" "&recSOP("QM_Name")%></option>
	  <% 
	   elseif  getType_Sop = "W" then
	   %>
	   <option value="<%=recSOP("W_Id")%>" <%=selected%> ><%response.write recSOP("W_Code")&" "&recSOP("W_Name")%></option>
	   <%
	 end if
	  recSOP.MoveNext
	  wend
	  recSOP.Close
	  Set recSOP = Nothing
	  %>
      </select>    
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>ชื่อ : 
      <input name="txtName" type="text" id="txtName" size="60" value="<%=getName_Review%>" /></td>
        <td>ตำแหน่ง : 
          <input type="text" name="txtPosition" id="txtPosition" size="60"  value="<%=getLevel_Review%>" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">ผลการทบทวน :</td>
        <td width="90%"><input name="radioPerfect" type="radio" id="radioPerfect" value="1" <% if getComport = true and getUncomport = false then response.write "checked=""checked""  " end if %> onClick="setPageEnable('2')"   />
          <label for="radioPerfect">มีความเหมาะสม ไม่ต้องดำเนินการใดๆ</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">&nbsp;</td>
        <td>เหตุผล :</td>
        <td><input type="checkbox"  id="chkCurrent" name="chkCurrent" value="เป็นปัจจุบัน"  <% if getLogic_comport1 <> "" then response.write " checked=""checked"" " end if%>  />
          <label for="chkCurrent">เป็นปัจจุบัน</label></td>
        <td><input type="checkbox" name="chkSupportWork" id="chkSupportWork"  value="สอดคล้องกับการปฏิบัติงาน" <% if getLogic_comport2 <> "" then response.write " checked=""checked"" " end if%> />
          <label for="chkSupportWork">สอดคล้องกับการปฏิบัติงาน</label></td>
        <td><input type="checkbox" name="chkBelongManual" id="chkBelongManual" value="มีการดำเนินการตามคู่มือ" <% if getLogic_comport3 <> "" then response.write " checked=""checked"" " end if%> />
          <label for="chkBelongManual">มีการดำเนินการตามคู่มือ</label></td>
      </tr>
      <tr>
        <td width="10%">&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse" id="chkElse" value="1" <% if getLogic_comport4 <> false then response.write " checked=""checked"" " end if%> />
          <label for="chkElse">อื่นๆ&nbsp;&nbsp;&nbsp;&nbsp;</label>
          <label for="txtElse"></label>
          <input name="txtElse" type="text" id="txtElse" size="70" value="<%=getLogic_comport5%>" /></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">ผลการทบทวน :</td>
        <td width="90%"><input type="radio" name="radioPerfect" id="radioPerfect2" value="0"  onclick="setPageEnable('1')" <% if getComport = false and getUncomport = true then response.write "checked=""checked"" " end if %>  />
          <label for="radioPerfect2">ไม่มีความเหมาะสม ต้องดำเนินการ</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%"><input type="radio" name="radioRemake" id="radioRemake1" value="1"  disabled="disabled"  <% if getMethod_Type = 1 then response.write " checked=""checked"" " end if %>    />
          <label for="radioRemake">จัดทำใหม่</label></td>
        <td width="15%">คาดว่าจะแล้วเสร็จวันที่</td>
        <td>
        <label>
              <select name="RemakefinishDay" id="RemakefinishDay" disabled="disabled">
                <%
				for i=1 to 31
				%>
                <option value="<%=i%>" <% if day(getRemake_Date) = i then response.write " selected=""selected"" " end if %>><%=i%></option>
                <%
				Next
				%>
              </select>
              เดือน
              <select name="RemakefinishMonth" id="RemakefinishMonth" disabled="disabled">
                <option value="1" <% if Month(getRemake_Date) = 1 then response.write " selected=""selected"" " end if %> >มกราคม</option>
                <option value="2" <% if Month(getRemake_Date) = 2 then response.write " selected=""selected"" " end if %>>กุมภาพันธ์</option>
                <option value="3" <% if Month(getRemake_Date) = 3 then response.write " selected=""selected"" " end if %>>มีนาคม</option>
                <option value="4" <% if Month(getRemake_Date) = 4 then response.write " selected=""selected"" " end if %>>เมษายน</option>
                <option value="5" <% if Month(getRemake_Date) = 5 then response.write " selected=""selected"" " end if %>>พฤษภาคม</option>
                <option value="6" <% if Month(getRemake_Date) = 6 then response.write " selected=""selected"" " end if %>>มิถุนายน</option>
                <option value="7" <% if Month(getRemake_Date) = 7 then response.write " selected=""selected"" " end if %>>กรกฎาคม</option>
                <option value="8" <% if Month(getRemake_Date) = 8 then response.write " selected=""selected"" " end if %>>สิงหาคม</option>
                <option value="9" <% if Month(getRemake_Date) = 9 then response.write " selected=""selected"" " end if %>>กันยายน</option>
                <option value="10" <% if Month(getRemake_Date) = 10 then response.write " selected=""selected"" " end if %>>ตุลาคม</option>
                <option value="11" <% if Month(getRemake_Date) = 11 then response.write " selected=""selected"" " end if %>>พฤศจิกายน</option>
                <option value="12"<% if Month(getRemake_Date) = 12 then response.write " selected=""selected"" " end if %>>ธันวาคม</option>
              </select>
              ปี
              <select name="RemakefinishYear" id="RemakefinishYear" disabled="disabled">
                <option value="2017" <% if Year(getRemake_Date) = 2017 then response.write " selected=""selected"" " end if %>>2560</option>
                <option value="2016" <% if Year(getRemake_Date) = 2016 then response.write " selected=""selected"" " end if %>>2559</option>
                <option value="2015" <% if Year(getRemake_Date) = 2015 then response.write " selected=""selected"" " end if %>>2558</option>
                <option value="2014" <% if Year(getRemake_Date) = 2014 then response.write " selected=""selected"" " end if %>>2557</option>
                                                        </select>
              </label>
        </td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake2" value="2" disabled="disabled" <% if getMethod_Type = 2 then response.write " checked=""checked"" " end if %> />
          <label for="radioRemake">แก้ไข</label></td>
        <td width="15%">คาดว่าจะแล้วเสร็จวันที่</td>
        <td><label>
              <select name="EditfinishDay" id="EditfinishDay" disabled="disabled">
                <%
				for i=1 to 31 
				%>
                <option value="<%=i%>" <% if day(getEdit_Date) = i then response.write " selected=""selected"" " end if %> ><%=i%></option>
                <%
				Next
				%>
              </select>
              เดือน
              <select name="EditfinishMonth" id="EditfinishMonth" disabled="disabled">
                <option value="1" <% if Month(getEdit_Date) = 1 then  response.write " selected=""selected"" " end if %>>มกราคม</option>
                <option value="2" <% if Month(getEdit_Date) = 2 then   response.write " selected=""selected"" " end if%>>กุมภาพันธ์</option>
                <option value="3" <% if Month(getEdit_Date) = 3 then  response.write " selected=""selected"" " end if %>>มีนาคม</option>
                <option value="4" <% if Month(getEdit_Date) = 4 then  response.write " selected=""selected"" " end if %>>เมษายน</option>
                <option value="5" <% if Month(getEdit_Date) = 5 then   response.write " selected=""selected"" " end if %>>พฤษภาคม</option>
                <option value="6" <% if Month(getEdit_Date) = 6 then   response.write " selected=""selected"" " end if%>>มิถุนายน</option>
                <option value="7" <% if Month(getEdit_Date) = 7 then   response.write " selected=""selected"" " end if %>>กรกฎาคม</option>
                <option value="8" <% if Month(getEdit_Date) = 8 then  response.write " selected=""selected"" " end if%>>สิงหาคม</option>
                <option value="9" <% if Month(getEdit_Date) = 9 then   response.write " selected=""selected"" " end if%>>กันยายน</option>
                <option value="10" <% if Month(getEdit_Date) = 10 then  response.write " selected=""selected"" " end if%>>ตุลาคม</option>
                <option value="11" <% if Month(getEdit_Date) = 11 then  response.write " selected=""selected"" " end if%>>พฤศจิกายน</option>
                <option value="12" <% if Month(getEdit_Date) = 12 then   response.write " selected=""selected"" " end if%>>ธันวาคม</option>
              </select>
              ปี
              <select name="EditfinishYear" id="EditfinishYear" disabled="disabled">
               <option value="2017" <% if Year(getEdit_Date) = 2017 then response.write " selected=""selected"" " end if %>>2560</option>
                <option value="2016" <% if Year(getEdit_Date) = 2016 then response.write " selected=""selected"" " end if %>>2559</option>
                <option value="2015" <% if Year(getEdit_Date) = 2015 then response.write " selected=""selected"" " end if %>>2558</option>
                <option value="2014" <% if Year(getEdit_Date) = 2014 then response.write " selected=""selected"" " end if %>>2557</option>
                                                        </select>
              </label></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake3" value="3" disabled="disabled" <% if getMethod_Type = 3 then response.write " checked=""checked"" " end if %> />
          <label for="radioRemake">ยกเลิก</label></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%">เหตุผล :</td>
        <td><input type="checkbox" name="chkNotNow" id="chkNotNow" value="ไม่เป็นปัจจุบัน"  disabled="disabled" <% if getLogic_Uncomport1 <> "" then response.write " checked=""checked"" " end if %> />
          <label for="chkNotNow">ไม่เป็นปัจจุบัน</label></td>
        <td><input type="checkbox" name="chkNotSupportWork" id="chkNotSupportWork" value="ไม่สอดคล้องกับการปฏิบัติงาน"ddisabled="disabled" <% if getLogic_Uncomport2 <> "" then response.write " checked=""checked"" " end if %>  />
          <label for="chkNotSupportWork">ไม่สอดคล้องกับการปฏิบัติงาน</label></td>
        <td><input type="checkbox" name="chkNewWayWork" id="chkNewWayWork" value="มีแนวทางการปฏิบัติงานใหม่" disabled="disabled" <% if getLogic_Uncomport3 <> "" then response.write " checked=""checked"" " end if %> />
          <label for="chkNewWayWork">มีแนวทางการปฏิบัติงานใหม่</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse2" id="chkElse2" value="1"  disabled="disabled" <% if getLogic_Uncomport4 <> false then response.write " checked=""checked"" " end if %> />
          <label for="chkElse2">อื่นๆ 
            <input name="txtElse2" type="text" id="txtElse2" size="70" disabled="disabled" value="<%=getLogic_Uncomport5%>" />
          </label></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>
          <input type="button" name="butSave" id="butSave" value="บันทึกข้อมูล" onClick="goSave()" />&nbsp;&nbsp;&nbsp; <input type="button"  value="กลับหน้ากรอกข้อมูล"  onclick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos/FReview.asp','_self');}"/> <%  if isEmpty(getRID) = False then response.write "<input type=""button""  value=""ดูหน้ารายงาน"" onclick=""openReportReview()""  />" end if  %></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
