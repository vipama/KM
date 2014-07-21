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
get_Manual_Name =GetSingleFieldQS("Tb_Manual","M_Name"," where M_Id="&get_Manual)
get_Manual_Code =GetSingleFieldQS("Tb_Manual","M_Code"," where M_Id="&get_Manual)
elseif get_radioReviewType = "Q" then
get_Manual_Name =GetSingleFieldQS("Tb_QM","QM_Name"," where QM_ID="&get_Manual)
get_Manual_Code =GetSingleFieldQS("Tb_QM","QM_Code"," where QM_Id="&get_Manual)
elseif get_radioReviewType = "W" then
get_Manual_Name =GetSingleFieldQS("Tb_Workin","W_Name"," where W_ID="&get_Manual)
get_Manual_Code =GetSingleFieldQS("Tb_Workin","W_Code"," where W_Id="&get_Manual)
end if

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
	dim chkRepeat,chkReviewNum
	dim SQL_LOG 
	chkRepeat = GetCountRowQS("Tb_Review","M_Code"," where M_Code='"&get_Manual_Code&"' and D_Id='"&get_DepartID&"'") 'ส่วนสำหรับตรวจสอบว่าในฐานมีข้อมูลอันนี้อยู่หรือไม่
	chkReviewNum = GetCountRowQS("Tb_Review","No_Review"," where No_Review='"&get_txtReviewNumber&"' ")
	if chkRepeat <> 0  then
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
	elseif chkRepeat = 0 and  chkReviewNum = 0 then
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
	else
	response.write "<script language=""javascript"">"
	response.write "alert(""รหัสเอกสารมีการใช้แล้วกรุณาอัพเดตหน้าเว็บและกรอกข้อมูลอีกครั้ง"");"
	response.write "window.location.href='FReview.asp';"
	response.write "</script>"
	end if
end if
'--------------------------------------------------------end block for sava data----------------------------------------------------

dim chkPS,chkPC,chkQ,chkW,FlagMain_Reserve
 chkPS=""
 chkPC=""
 chkQ=""
 chkW=""
 FlagMain_Reserve=""
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = "1"
	 end if
else
	getDid=Request.QueryString("id")
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
	getTid = "PC"
	FlagMain_Reserve = "M_Main=1"
	chkPC = "checked=""checked"""
end if
'response.write FlagMain_Reserve
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
%>
<body>
<form name="frmFReview" method="post" enctype="application/x-www-form-urlencoded" action="FReview.asp">
<input type="hidden"  name="hidS" id="hidS" value=""/>
<%
if isEmpty(getRID) = False then
response.write "<input type=""hidden"" value="""&getRID&""" name=""hidRID"" />"
end if
%>
<div align="center"  style="font-size:24px;"><strong>แบบทบทวนกระบวนงาน</strong></div><br />
<table width="85%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr><th align="right">No Review : <input type="text"  name="txtReviewNumber" id="txtReviewNumber" readonly value="<%=(year(Now)+543)%>-7-<%=runNum%>" /></th></tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>ดำเนินการทบทวน</td>
        <td><input type="radio" name="radioReviewType" id="radioReviewType1" value="Q" <%=chkQ%>onclick="ChangeJobresultGroup('','')"  />
          <label>คู่มือคุณภาพ</label>
&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType2" value="PC"  <%=chkPC%> onClick="ChangeJobresultGroup('','')"   />
          <label >คู่มือขั้นตอนการปฏิบัติงาน (P) (Core Process)</label></td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType3" value="W" <%=chkW%> onClick="ChangeJobresultGroup('','')" />
          <label >คู่มือขั้นวิธีการปฏิบัติงาน (W)</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType4" value="PS" <%=chkPS%>  onclick="ChangeJobresultGroup('','')" />
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
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:14px"   >
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
	  if getTid = "PC" or getTid = "PS" then
	  	sqlSOP = "select  *  from  Tb_Manual where  D_Id='"&getDid&"' and "&FlagMain_Reserve&" order by M_Id  asc"
	  elseif getTid = "W" then
	   sqlSOP = "select  *  from  Tb_Workin where  D_Id='"&getDid&"' order by W_Id  asc"
	  elseif getTid = "Q" then
	  	sqlSOP = "select  *  from  Tb_QM where  D_Id='"&getDid&"' order by QM_Id  asc"
	  else
	  response.write "ssssss"
	  end if
	  'response.write sqlSOP 
	  recSOP.open sqlSOP,ConQS,1,3
	  %>
	  <select name="Manual" style="font-size:14px"  >
	  <%
	  while not recSOP.EOF
	'  if recSOP("M_Id") = getDid then
	'  selected = "selected=""selected"""
	'  else
	'  selected = ""
	'  end if
	 if  getTid = "PS" or getTid = "PC" then 
	  %>
	  <option value="<%=recSOP("M_Id")%>" <%=selected%> ><%response.write recSOP("M_Code")&" "&recSOP("M_Name")%></option>
	  <%
	  elseif  getTid = "Q" then
	  %>
	  <option value="<%=recSOP("QM_Id")%>" <%=selected%> ><%response.write recSOP("QM_Code")&" "&recSOP("QM_Name")%></option>
	  <% 
	   elseif  getTid = "W" then
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
      <input name="txtName" type="text" id="txtName" size="60" /></td>
        <td>ตำแหน่ง : 
          <input type="text" name="txtPosition" id="txtPosition" size="60" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">ผลการทบทวน :</td>
        <td width="90%"><input name="radioPerfect" type="radio" id="radioPerfect" value="1" checked="checked"  onclick="setPageEnable('2')" />
          <label for="radioPerfect">มีความเหมาะสม ไม่ต้องดำเนินการใดๆ</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">&nbsp;</td>
        <td>เหตุผล :</td>
        <td><input type="checkbox"  id="chkCurrent" name="chkCurrent" value="เป็นปัจจุบัน"  />
          <label for="chkCurrent">เป็นปัจจุบัน</label></td>
        <td><input type="checkbox" name="chkSupportWork" id="chkSupportWork"  value="สอดคล้องกับการปฏิบัติงาน" />
          <label for="chkSupportWork">สอดคล้องกับการปฏิบัติงาน</label></td>
        <td><input type="checkbox" name="chkBelongManual" id="chkBelongManual" value="มีการดำเนินการตามคู่มือ" />
          <label for="chkBelongManual">มีการดำเนินการตามคู่มือ</label></td>
      </tr>
      <tr>
        <td width="10%">&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse" id="chkElse" value="1" />
          <label for="chkElse">อื่นๆ&nbsp;&nbsp;&nbsp;&nbsp;</label>
          <label for="txtElse"></label>
          <input name="txtElse" type="text" id="txtElse" size="70" /></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">ผลการทบทวน :</td>
        <td width="90%"><input type="radio" name="radioPerfect" id="radioPerfect2" value="0"  onclick="setPageEnable('1')"  />
          <label for="radioPerfect2">ไม่มีความเหมาะสม ต้องดำเนินการ</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%"><input type="radio" name="radioRemake" id="radioRemake1" value="1" checked="checked" disabled="disabled" />
          <label for="radioRemake">จัดทำใหม่</label></td>
        <td width="15%">คาดว่าจะแล้วเสร็จวันที่</td>
        <td>
        <label>
              <select name="RemakefinishDay" id="RemakefinishDay" disabled="disabled">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
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
              </select>
              เดือน
              <select name="RemakefinishMonth" id="RemakefinishMonth" disabled="disabled">
                <option value="1" selected="selected">มกราคม</option>
                <option value="2">กุมภาพันธ์</option>
                <option value="3">มีนาคม</option>
                <option value="4">เมษายน</option>
                <option value="5">พฤษภาคม</option>
                <option value="6">มิถุนายน</option>
                <option value="7">กรกฎาคม</option>
                <option value="8">สิงหาคม</option>
                <option value="9">กันยายน</option>
                <option value="10">ตุลาคม</option>
                <option value="11">พฤศจิกายน</option>
                <option value="12">ธันวาคม</option>
              </select>
              ปี
              <select name="RemakefinishYear" id="RemakefinishYear" disabled="disabled">
                <option value="2017">2560</option>
                <option value="2016">2559</option>
                <option value="2015">2558</option>
                <option value="2014" selected="selected">2557</option>
                                                        </select>
              </label>
        </td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake2" value="2" disabled="disabled" />
          <label for="radioRemake">แก้ไข</label></td>
        <td width="15%">คาดว่าจะแล้วเสร็จวันที่</td>
        <td><label>
              <select name="EditfinishDay" id="EditfinishDay" disabled="disabled">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
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
              </select>
              เดือน
              <select name="EditfinishMonth" id="EditfinishMonth" disabled="disabled">
                <option value="1" selected="selected">มกราคม</option>
                <option value="2">กุมภาพันธ์</option>
                <option value="3">มีนาคม</option>
                <option value="4">เมษายน</option>
                <option value="5">พฤษภาคม</option>
                <option value="6">มิถุนายน</option>
                <option value="7">กรกฎาคม</option>
                <option value="8">สิงหาคม</option>
                <option value="9">กันยายน</option>
                <option value="10">ตุลาคม</option>
                <option value="11">พฤศจิกายน</option>
                <option value="12">ธันวาคม</option>
              </select>
              ปี
              <select name="EditfinishYear" id="EditfinishYear" disabled="disabled">
                <option value="2017">2560</option>
                <option value="2016">2559</option>
                <option value="2015">2558</option>
                <option value="2014" selected="selected">2557</option>
                                                        </select>
              </label></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake3" value="3" disabled="disabled" />
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
        <td><input type="checkbox" name="chkNotNow" id="chkNotNow" value="ไม่เป็นปัจจุบัน"  disabled="disabled" />
          <label for="chkNotNow">ไม่เป็นปัจจุบัน</label></td>
        <td><input type="checkbox" name="chkNotSupportWork" id="chkNotSupportWork" value="ไม่สอดคล้องกับการปฏิบัติงาน"ddisabled="disabled" />
          <label for="chkNotSupportWork">ไม่สอดคล้องกับการปฏิบัติงาน</label></td>
        <td><input type="checkbox" name="chkNewWayWork" id="chkNewWayWork" value="มีแนวทางการปฏิบัติงานใหม่" disabled="disabled" />
          <label for="chkNewWayWork">มีแนวทางการปฏิบัติงานใหม่</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse2" id="chkElse2" value="1"  disabled="disabled" />
          <label for="chkElse2">อื่นๆ 
            <input name="txtElse2" type="text" id="txtElse2" size="70" disabled="disabled" />
          </label></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>
          <input type="button" name="butSave" id="butSave" value="บันทึกข้อมูล" onClick="goSave()" />&nbsp;&nbsp;&nbsp; <input type="button"  value="กลับหน้าแรก"  onclick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos','_self');}"/> <%  if isEmpty(getRID) = False then response.write "<input type=""button""  value=""ดูหน้ารายงาน"" onclick=""openReportReview()""  />" end if  %>  &nbsp;&nbsp;&nbsp;<input type="button" value="แก้ไข" onClick="goReviewEdit()"  />&nbsp;&nbsp;&nbsp;<input type="button" value="พิมพ์รายงาน" onClick="goReviewReport()"  />&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="txtREditSOP" id="txtREditSOP" />&nbsp;&nbsp;&nbsp;หมายเหตุ กรุณาใส่รหัสเอกสารคุณภาพที่ต้องการแก้ไข / พิมพ์รายงาน</td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
