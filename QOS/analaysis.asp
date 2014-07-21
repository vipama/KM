<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim chkShowSave
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
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
dim getMID
'-----------------------------------------------------------block for sava data-----------------------------------------------------
if isEmpty(Request.Form("hidSave")) = False then
	dim Depart_Id
	dim Manual_Id
	dim Manual_Code,Manual_Name,MAccess,flagstrategy1,flagstrategy2,flagstrategy3
	flagstrategy1=0
	flagstrategy2=0
	flagstrategy3=0
	Depart_Id = Request.Form("DepartID")
	Manual_Id = Request.Form("Manual")
	getMID=Manual_Id
	Manual_Code = GetSingleFieldQS("Tb_Manual","M_Code"," where M_Id="&Manual_Id)
	Manual_Name =  GetSingleFieldQS("Tb_Manual","M_Name"," where M_Id="&Manual_Id)
	Strategic1 = Request.Form("chkStrategic1")
	Strategic2 = Request.Form("chkStrategic2")
	Strategic3 = Request.Form("chkStrategic3")
	
	
	Strategy11 = Request.Form("chkStrategy11")
	Strategy12 = Request.Form("chkStrategy12")
	Strategy13 = Request.Form("chkStrategy13")
	Strategy14 = Request.Form("chkStrategy14")
	Strategy15 = Request.Form("chkStrategy15")
	Strategy16 = Request.Form("chkStrategy16")
	Strategy17 = Request.Form("chkStrategy17")
	Strategy21 = Request.Form("chkStrategy21")
	Strategy22 = Request.Form("chkStrategy22")
	Strategy23 = Request.Form("chkStrategy23")
	Strategy24 = Request.Form("chkStrategy24")
	Strategy31 = Request.Form("chkStrategy31")
	Strategy32 = Request.Form("chkStrategy32")
	Strategy33 = Request.Form("chkStrategy33")
	Strategy34 = Request.Form("chkStrategy34")
	Strategy35 = Request.Form("chkStrategy35")
	
	if isEmpty(Strategy11) = true then
	Strategy11="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy12) = true then
	Strategy12="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy13) = true then
	Strategy13="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy14) = true then
	Strategy14="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy15) = true then
	Strategy15="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy16) = true then
	Strategy16="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy17) = true then
	Strategy17="00"
	flagstrategy1 = flagstrategy1+1
	end if
	if isEmpty(Strategy21) = true then
	Strategy21="00"
	flagstrategy2 = flagstrategy2+1
	end if
	if isEmpty(Strategy22) = true then
	Strategy22="00"
	flagstrategy2 = flagstrategy2+1
	end if
	if isEmpty(Strategy23) = true then
	Strategy23="00"
	flagstrategy2 = flagstrategy2+1
	end if
	if isEmpty(Strategy24) = true then
	Strategy24="00"
	flagstrategy2 = flagstrategy2+1
	end if
	if isEmpty(Strategy31) = true then
	Strategy31="00"
	flagstrategy3 = flagstrategy3+1
	end if
	if isEmpty(Strategy32) = true then
	Strategy32="00"
	flagstrategy3 = flagstrategy3+1
	end if
	if isEmpty(Strategy33) = true then
	Strategy33="00"
	flagstrategy3 = flagstrategy3+1
	end if
	if isEmpty(Strategy34) = true then
	Strategy34="00"
	flagstrategy3 = flagstrategy3+1
	end if
	if isEmpty(Strategy35) = true then
	Strategy35="00"
	flagstrategy3 = flagstrategy3+1
	end if
	
	if (isEmpty(Strategic1) = true) and (flagstrategy1 = 7)  then
		Strategic1="0"
	elseif (isEmpty(Strategic1) = true) and (flagstrategy1 < 7) then  
		Strategic1="1"
	end if
	
	if (isEmpty(Strategic2) = true) and (flagstrategy2 = 4) then
		Strategic2="0"
	elseif (isEmpty(Strategic2) = true) and (flagstrategy2 < 4) then 
	 	Strategic2="2"
	end if
	
	if (isEmpty(Strategic3) = true)  and (flagstrategy3 = 5) then
		Strategic3="0"
	elseif (isEmpty(Strategic3) = true) and (flagstrategy3 < 5) then
		Strategic3="3" 
	end if
	
	Support = Request.Form("radioSupport")
	Period = Request.Form("radioPeriod")
	Quality = Request.Form("radioQuality")
	Charge = Request.Form("radioCharge")
	TotalSum = Request.Form("txtSumAll")
	
	chkValue = Request.Form("chkValue")
	chkFairness = Request.Form("chkFairness")
	chkAccuracy = Request.Form("chkAccuracy")
	chkTransparency = Request.Form("chkTransparency")
	chkParticipation = Request.Form("chkParticipation")
	chkResponse = Request.Form("chkResponse")
	chkEase = Request.Form("chkEase")
	chkWorksupport = Request.Form("chkWorksupport")
	chkElse = Request.Form("chkElse")
	
	if isEmpty(chkValue) = true then
	chkValue="0"
	end if
	if isEmpty(chkFairness) = true then
	chkFairness="0"
	end if
	if isEmpty(chkAccuracy) = true then
	chkAccuracy="0"
	end if
	if isEmpty(chkTransparency) = true then
	chkTransparency="0"
	end if
	if isEmpty(chkParticipation) = true then
	chkParticipation="0"
	end if
	if isEmpty(chkResponse) = true then
	chkResponse="0"
	end if
	if isEmpty(chkEase) = true then
	chkEase="0"
	end if
	if isEmpty(chkWorksupport) = true then
	chkWorksupport="0"
	end if
	if isEmpty(chkElse) = true then
	chkElse="0"
	end if
	
	
	txtElse = Request.Form("txtElse")
	
	

	fulStrategic = Strategic1&","&Strategic2&","&Strategic3
	
	if GetSingleFieldQS("Tb_AnalisProcedure","M_Id"," where M_Id="&Manual_Id) <> 0 then

'SQL = "update Tb_AnalisProcedure set Analis_Strategic='"&Strategic1&","&Strategic2&","&Strategic3&"' , Analis_Strategy='"&Strategy11&","&Strategy12&","&Strategy13&","&Strategy14&","&Strategy15&","&Strategy16&","&Strategy17&","&Strategy21&","&Strategy22&","&Strategy23&","&Strategy24&","&Strategy31&","&Strategy32&","&Strategy33&","&Strategy34&","&Strategy35&"' , Analis_Support="&Support&" , Analis_period="&Period&" , Analis_Quality="&Quality&" , Analis_Charge="&Charge&" , Analis_Sum="&TotalSum&" , Analis_value="&chkValue&" , Analis_Fairness="&chkFairness&" , Analis_Accuracy="&chkAccuracy&" , Analis_Transparency="&chkTransparency&" , Analis_Participation="&chkParticipation&" , Analis_Response="&chkResponse&" , Analis_Ease="&chkEase&" , Analis_Worksupport="&chkWorksupport&" , Analis_else="&chkElse" , Analis_DesElse='"&txtElse&"' where M_Id="&Manual_Id&" and M_Code='"&Manual_Code&"'"

		SQL = "update Tb_AnalisProcedure set Analis_Strategic='"&fulStrategic&"' , Analis_Strategy='"&Strategy11&","&Strategy12&","&Strategy13&","&Strategy14&","&Strategy15&","&Strategy16&","&Strategy17&","&Strategy21&","&Strategy22&","&Strategy23&","&Strategy24&","&Strategy31&","&Strategy32&","&Strategy33&","&Strategy34&","&Strategy35&"' , Analis_Support="&Support&" , Analis_period="&Period&" , Analis_Quality="&Quality&" , Analis_Charge="&Charge&" , Analis_Sum="&TotalSum&" , Analis_value="&chkValue&" , Analis_Fairness="&chkFairness&" , Analis_Accuracy="&chkAccuracy&" , Analis_Transparency="&chkTransparency&" , Analis_Participation="&chkParticipation&"  , Analis_Response="&chkResponse&"  , Analis_Ease="&chkEase&"  , Analis_Worksupport="&chkWorksupport&"  , Analis_else="&chkElse&" , Analis_DesElse='"&txtElse&"' , Analis_Date='"&Datemmddyyyy&"'  where M_Id="&Manual_Id&" and M_Code='"&Manual_Code&"'"
		MAccess="Update"
		'response.write SQL
	else
		SQL = "insert into Tb_AnalisProcedure (Analis_Strategic,M_Id,M_Code,M_Name,Analis_Strategy,Analis_Support,Analis_period,Analis_Quality,Analis_Charge,Analis_Sum,Analis_value,Analis_Fairness,Analis_Accuracy,Analis_Transparency,Analis_Participation,Analis_Response,Analis_Ease,Analis_Worksupport,Analis_else,Analis_DesElse,Analis_Date) values ('"&fulStrategic&"',"&Manual_Id&",'"&Manual_Code&"','"&Manual_Name&"','"&Strategy11&","&Strategy12&","&Strategy13&","&Strategy14&","&Strategy15&","&Strategy16&","&Strategy17&","&Strategy21&","&Strategy22&","&Strategy23&","&Strategy24&","&Strategy31&","&Strategy32&","&Strategy33&","&Strategy34&","&Strategy35&"',"&Support&","&Period&","&Quality&","&Charge&","&TotalSum&","&chkValue&","&chkFairness&","&chkAccuracy&","&chkTransparency&","&chkParticipation&","&chkResponse&","&chkEase&","&chkWorksupport&","&chkElse&",'"&txtElse&"','"&Datemmddyyyy&"')"
		MAccess="Insert"
		'response.write SQL
	end if
	ConQS.execute(SQL)
	
	'---------------------------------------------------------------------Start Block for Add to Log table------------------------------------------------------------------------------
	SQL_LOG = "insert into Tb_LogAnalisProcedure (User_Id,Method_Access,Date_Access,Department_Name,M_Code) values ('"&session("member")&"','"&MAccess&"','"&Datemmddyyyy&"','"&getDepartmentname(Depart_Id)&"','"&Manual_Code&"')"
	ConQS.execute(SQL_LOG)
	'---------------------------------------------------------------------End Block for Add to Log table---------------------------------------------------------------------------------
	
		response.write "<script language=""javascript"">"
		response.write "alert(""บันทึกข้อมูลเรียบร้อยค่ะ"");"
		response.write "</script>"
	chkShowSave = "User : "&session("member")&" ได้ทำการปรับปรุงข้อมูลของ <BR /> กระบวนงาน รหัส : "&Manual_Code&"  หน่วยงาน : "&getDepartmentname(Depart_Id)&"   เวลา : "&Datemmddyyyy&" <br> ข้อมูลนี้ได้ถูกเก็บเป็นประวัติการแก้ไขในฐานข้อมูลเรียบร้อยแล้วค่ะ"
	
end if
'----------------------------------------------------------------------------------------------------------------------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>การวิเคาระห์ระบบงาน</title>
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
		// alert(val+"/"+val1);
		window.location.href="analaysis.asp?id="+val+"&oid="+val1;
}
</script>
<script type="text/javascript" src="JScript/JS.js"></script>
</head>

<body>
<form name="frmAnalaysis" enctype="application/x-www-form-urlencoded" method="post" action="analaysis.asp">
<input type="hidden" value="3" name="hidSum1" />
<input type="hidden" value="3" name="hidSum2" />
<input type="hidden" value="3" name="hidSum3" />
<input type="hidden" value="3" name="hidSum4" />
<input type="hidden" value="S" name="hidSave" />
<%
if isEmpty(getMID) = False then
response.write "<input type=""hidden"" value="""&getMID&""" name=""hidMID"" />"
end if
%>
<%
if isEmpty(Request.Form("hidSave")) = False then
%>
<table  width="85%" cellpadding="2" cellspacing="0" border="0" align="center">
<tr><td align="center" style=" font-size:18px; color:#FF0000; font-weight:200"><%=chkShowSave%></td></tr>
</table>
<%
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="5">
  <tr><th align="center">แบบฟอร์มการวิเคราะห์กระบวนงานของสำนักงานอาหารและยา ประจำปีงบประมาณ พ.ศ.2557</th></tr>
  <tr>
    <td>หน่วยงาน : 
     <%
	  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
	  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
	  recDepart.open sqlDepart,ConQS,1,3
	  %>
	  <select name="DepartID" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:18px"  >
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
      </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%  if isEmpty(getMID) = False then response.write "<input type=""button""  value=""ดูหน้ารายงาน"" onclick=""openReport()""  />" end if  %>    </td>
  </tr>
  <tr>
    <td>
    กระบวนงาน : 
     <%
	  Set   recSOP = Server.CreateObject("ADODB.RECORDSET")
	  sqlSOP = "select  *  from  Tb_Manual where  D_Id='"&getDid&"' and M_Main=1 order by M_Id  asc"
	  recSOP.open sqlSOP,ConQS,1,3
	  %>
	  <select name="Manual" id="Manual" style="font-size:18px"  >
	  <%
	  while not recSOP.EOF
	'  if recSOP("M_Id") = getDid then
	'  selected = "selected=""selected"""
	'  else
	'  selected = ""
	'  end if
	  %>
	  <option value="<%=recSOP("M_Id")%>" <%=selected%> ><%response.write recSOP("M_Code")&" "&recSOP("M_Name")%></option>
	  <%
	  recSOP.MoveNext
	  wend
	  recSOP.Close
	  Set recSOP = Nothing
	  %>
      </select>    </td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%" class="style1"><strong>ความสอดคล้อง :</strong> </td>
        <td width="20%" class="style1">ประเด็นยุทธศาสตร์</td>
        <td width="70%" class="style1">กลยุทธ์</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td class="style1"><label>
          <input type="checkbox" name="chkStrategic1" id="chkStrategic1" value="1" />
          1. พัฒนาระบบการควบคุมกำกับดูแลผลิตภัณฑ์สุขภาพให้มีประสิทธิภาพทัดเทียมระดับสากล</label></td>
        <td><span class="style1">
          <label>
          <input type="checkbox" name="chkStrategy11" id="chkStrategy11" value="11" />
1. พัฒนาและปรับปรุงกฏหมายด้านการคุ้มครองผู้บริโภคด้านผลิตภัณฑ์สุขภาพให้ทันต่อสถานการณ์และสอดคล้องกับสากล </label>
        </span></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy12" id="chkStrategy12" value="12" />
2. พัฒนาระบบการควบคุม กำกับดูแลผลิตภัณฑ์สุขภาพให้เป็นมาตรฐานเดียวกันทั่วประเทศและสามารถเทียบเคียงได้ในระดับสากล </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy13" id="chkStrategy13" value="13" />
3. เพิ่มประสิทธิภาพการดำเนินงานควบคุม กำกับดูแลผลิตภัณฑ์สุขภาพ โดยการถ่ายโอนภารกิจให้ภาคเอกชนหรือหน่วยงานอื่นที่มีศักยภาพและส่งเสริมให้องค์กรปกครองส่วนท้องถิ่นมีการดำเนินงานคุ้มครองผู้บริโภค </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy14" id="chkStrategy14" value="14" />
4. พัฒนาระบบการจัดการด้านการเงินเพื่อเพิ่มศักยภาพในการคุ้มครองผู้บริโภค </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy15" id="chkStrategy15" value="15" />
5.สร้างความเข้มแข็งของเครือข่ายและเปิดโอกาสให้ผู้มีส่วนได้ส่วนเสียเข้ามามีส่วนร่วม ในการดำเนินงานคุ้มครองผู้บริโภคด้านผลิตภัณฑ์สุขภาพ </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy16" id="chkStrategy16" value="16" />
6.พัฒนาประสิทธิภาพและคุณภาพงานบริการ พิจารณาอนุญาตให้มีความโปร่งใสและเป็นธรรม </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td class="style1"><label>
        <input type="checkbox" name="chkStrategy17" id="chkStrategy17" value="17" />
7.พัฒนาการบริหารจัดการองค์กรสู่ความเป็นเลิศ </label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td class="style1"><label>
          <input type="checkbox" name="chkStrategic2" id="chkStrategic2"  value="2" />
          2. พัฒนาผู้บริโภคให้มีศักยภาพเพื่อการพึ่งพิงตนเองได้ในการบริโภคผลิตภัณฑ์สุขภาพ</label></td>
        <td><span class="style1">
          <label>
          <input type="checkbox" name="chkStrategy21" id="chkStrategy21" value="21" />
1. เสริมสร้างความรู้ของประชาชนในการเลือกซื้อเลือกบริโภคผลิตภัณฑ์สุขภาพ </label>
        </span></td>
      </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy22" id="chkStrategy22" value="22" />
2. สร้างความตระหนักเพื่อการปรับเปลี่ยนพฤติกรรมการบริโภคผลิตภัณฑ์สุขภาพที่ถูกต้อง</label></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy23" id="chkStrategy23" value="23" />
3. สร้างและพัฒนาภาคีเครือข่ายด้านผลิตภัณฑ์สุขภาพ  โดยถ่ายทอดเชื่อมโยงความรู้สู่บุคคล ครอบครัว ชุมชน สังคม</label></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy24" id="chkStrategy24" value="24" />
4. พัฒนาการบริหารจัดการองค์กรสู่ความเป็นเลิศ</label></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
       </tr>
       <tr>
        <td>&nbsp;</td>
        <td class="style1"><label>
          <input type="checkbox" name="chkStrategic3" id="chkStrategic3" value="3" />
          3. การควบคุมตัวยาและสารตั้งต้นที่เป็นวัตถุเสพติด</label></td>
        <td><span class="style1">
          <label>
          <input type="checkbox" name="chkStrategy31" id="chkStrategy31" value="31" />
1. พัฒนาและเพิ่มประสิทธิภาพระบบการเฝ้าระวังและติดตามการเคลื่อนไหวของตัวยา  เคมีภัณฑ์จำเป็น และสารตั้งต้นด้านวัตถุเสพติด ที่อยู่ในความรับผิดชอบของ อย.ให้มีมาตรฐานเดียวกันทั่วประเทศ</label>
        </span></td>
      </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy32" id="chkStrategy32" value="32" />
2. พัฒนาและเพิ่มประสิทธิภาพการกำกับดูแลผลิตภัณฑ์สุขภาพที่เกี่ยวกับวัตถุเสพติดที่ใช้ในทางการแพทย์ด้านคุณภาพและความปลอดภัยที่เป็นไปตามมาตรฐานการบริการ</label></td>
       </tr>
       <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy33" id="chkStrategy33" value="33" />
3. พัฒนาระบบเครือข่ายสารสนเทศเกี่ยวกับตัวยา  เคมีภัณฑ์จำเป็น และสารตั้งต้นด้านวัตถุเสพติด ให้สามารถสื่อสารกันได้ระหว่างหน่วยงานภาครัฐและภาคเอกชนที่ได้รับอนุญาตตามกฎหมายว่าด้วยวัตถุเสพติด</label></td>
       </tr>
        <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy34" id="chkStrategy34" value="34" />
4. พัฒนาและปรับปรุงกฎหมายในการกำกับดูแลตัวยา  เคมีภัณฑ์จำเป็น และสารตั้งต้น ให้ทันต่อสถานการณ์และสอดคล้องกับระบบสากล</label></td>
       </tr>
        <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1"><label>
         <input type="checkbox" name="chkStrategy35" id="chkStrategy35" value="35" />
5. พัฒนาการบริหารจัดการองค์กรสู่ความเป็นเลิศ</label></td>
       </tr>
        <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td class="style1">&nbsp;</td>
       </tr>
    </table></td>
  </tr>
  <tr>
    <td class="style1"><strong>เกณฑ์การประเมิน  :</strong></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="40%" class="style1">ให้ระบุคะแนนเพื่อประเมินประโยชน์ที่จะได้รับจากแต่ละกระบวนงาน  โดยสามารถให้คะแนนได้ 3 ระดับ คือ</td>
        <td width="60%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="20%" align="center" class="style1">สนับสนุนพันธกิจขององค์กร</td>
            <td width="20%" align="center" class="style1">ลดระยะเวลาในการทำงาน</td>
            <td width="20%" align="center" class="style1">เพิ่มคุณภาพการให้บริการ</td>
            <td width="20%" align="center" class="style1">ลดค่าใช้จ่ายในการทำงาน</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td class="style1">ระดับ  3 มีความสำคัญหรือประโยชน์ที่คาดว่าจะได้รับ<u>สูง</u></td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%" align="center"><input type="radio" name="radioSupport" id="radioSupport3" value="3" checked="checked" onClick="showSum('1',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioPeriod" id="radioPeriod3" value="3" checked="checked" onClick="showSum('2',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioQuality" id="radioQuality3" value="3" checked="checked" onClick="showSum('3',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioCharge" id="radioCharge3" value="3" checked="checked" onClick="showSum('4',this.value)" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td class="style1">ระดับ  2 มีความสำคัญหรือประโยชน์ที่คาดว่าจะได้รับ<u>ปานกลาง</u></td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%" align="center"><input type="radio" name="radioSupport" id="radioSupport2" value="2" onClick="showSum('1',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioPeriod" id="radioPeriod2" value="2"  onClick="showSum('2',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioQuality" id="radioQuality2" value="2" onClick="showSum('3',this.value)"  /></td>
            <td width="20%" align="center"><input type="radio" name="radioCharge" id="radioCharge2" value="2" onClick="showSum('4',this.value)" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td class="style1">ระดับ  1 มีความสำคัญหรือประโยชน์ที่คาดว่าจะได้รับ<u>ต่ำ</u></td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%" align="center"><input type="radio" name="radioSupport" id="radioSupport1" value="1" onClick="showSum('1',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioPeriod" id="radioPeriod1" value="1" onClick="showSum('2',this.value)" /></td>
            <td width="20%" align="center"><input type="radio" name="radioQuality" id="radioQuality1" value="1" onClick="showSum('3',this.value)"  /></td>
            <td width="20%" align="center"><input type="radio" name="radioCharge" id="radioCharge1" value="1" onClick="showSum('4',this.value)" /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
      <td>&nbsp;</td>
      <td>คะแนนรวม : <input type="text"  name="txtSumAll" id="txtSumAll" readonly width="20" value="12"/></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><strong>ความต้องการของผู้รับบริการ :  (ใส่เครื่องหมาย / โดยตอบได้มากกว่า 1 ข้อ)</strong></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkValue" id="chkValue" value="1" />
    1. ความคุ้มค่า		หมายถึง		สามารถควบคุมความเสี่ยงที่จะเกิดขึ้น โดยไม่สิ้นเปลืองค่าใช้จ่ายมากเกินไป</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkFairness" id="chkFairness" value="1" />
    2. ความเป็นธรรม	หมายถึง		ให้บริการอย่างเสมอภาค มีมาตรฐาน ไม่เลือกปฏิบัติ</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkAccuracy" id="chkAccuracy" value="1" />
    3. ความถูกต้อง		หมายถึง		ไม่มีการเลี่ยงหรือใช้ช่องโหว่ทางกฎหมายในทางที่ผิด</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkTransparency" id="chkTransparency" value="1" />
    4. ความโปร่งใส		หมายถึง		มีหลักเกณฑ์ในการกำกับ ตรวจสอบ ที่ชัดเจน</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkParticipation" id="chkParticipation" value="1" />
    5. การมีส่วนร่วม		หมายถึง		มีการรับฟังความคิดเห็น/ช่องทางในการรับฟังความคิดเห็นจากผู้รับบริการ</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkResponse" id="chkResponse" value="1" />
    6. มีผู้รับผิดชอบชัดเจน	หมายถึง		มีผู้รับผิดชอบในแต่ละขั้นตอนที่ชัดเจน สามารถระบุตัวเจ้าภาพได้</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkEase" id="chkEase" value="1" />
    7. ความสะดวกรวดเร็ว	หมายถึง		ความรวดเร็วในขั้นตอนบริการ การมีคำชี้แจง/บอกวิธีการรับบริการ รวมทั้งสามารถรับบริการได้โดยไม่ยุ่งยาก</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkWorksupport" id="chkWorksupport" value="1" />
    8. สนับสนุนการปฏิบัติงาน	หมายถึง		เพิ่มประสิทธิภาพการดำเนินงานให้แก่องค์กร บุคลากร และการปฏิบัติงานประจำวัน</label></td>
  </tr>
  <tr>
    <td><label>
      <input type="checkbox" name="chkElse" id="chkElse" value="1" />
    9. อื่นๆ (โปรดระบุ) 
    <input name="txtElse" type="text" id="txtElse" size="60" />
    </label></td>
  </tr>
  <tr>
  <td><input type="button"  value="บันทึกข้อมูล"  onclick="AnalisCheckSave()" />&nbsp;&nbsp;&nbsp;<input type="button" value="หน้าแรก" onClick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos','_self')}" /> <%  if isEmpty(getMID) = False then response.write "<input type=""button""  value=""ดูหน้ารายงาน"" onclick=""openReport()""  />" end if  %>  &nbsp;&nbsp;&nbsp;<input type="button" value="แก้ไข" onclick="goAnalisEdit()"  />&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="txtEditSOP" id="txtEditSOP" />&nbsp;&nbsp;&nbsp;หมายเหตุ กรุณาใส่รหัสเอกสารคุณภาพที่ต้องการแก้ไข</td>
  </tr>
</table>
</form>
</body>
</html>