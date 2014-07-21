<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->

<%
	if isEmpty(session("member")) = True then
		'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
	end if
	if isEmpty(Request.QueryString("id")) = true then
		 if isEmpty(Request.Form("hidDid")) = false then
			getDid=Request.Form("hidDid")
		 else
			getDid = "0"
		 end if
	else
		getDid=Request.QueryString("id")
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>รายงานการวิเคราะห์ความสอดคล้องและความต้องการของผู้รับบริการ</title>
<style type="text/css">
<!--
.style1 {
font-size:10px;
font-family:Arial, Helvetica, sans-serif;
}
-->
</style>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		// alert(val+"/"+val1);
		window.location.href="ReportAnalaysis.asp?id="+val+"&oid="+val1;
}
</script>
</head>

<body>
<!--start show core and support-->
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#999999">
    <td width="40%" align="center" class="style3">สำนัก / กอง / กลุ่ม</td>
    <td width="20%" align="center" class="style3">กระบวนการหลัก</td>
    <td width="20%" align="center" class="style3">ดำเนินการวิเคราะห์</td>
    <td width="20%" align="center" class="style3">คงเหลือ</td>
  </tr>
  <tr bgcolor="#FFFF99">
<td  align="left" class="style3" >กองผลิตภัณฑ์</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
</tr>
  <%
  dim  countsumcore,countsumAnalis,countsumAllAnalis,colorSet
  countsumcore=0
  countsumAnalis=0
  countsumAllAnalis=0
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='0' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
    <td align="center" class="style3">	
			<% 
			response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") 
			countsumcore = countsumcore+ GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			%>
    </td>
    <%
	if (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") = 0 and getCountRowAnalis(RecDepart("D_Id"))=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (getCountRowAnalis(RecDepart("D_Id"))<>0 and (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id")) <> 0) )  then
		colorSet=" bgcolor=""#FFFF00"""
	else
		 colorSet = "bgcolor=""#99FF33"""
	end if  
	%>
    <td align="center" class="style3" <%=colorSet%>>
	<% 
	response.write getCountRowAnalis(RecDepart("D_Id"))
	countsumAnalis = countsumAnalis+getCountRowAnalis(RecDepart("D_Id"))
	%></td>
    <td align="center" class="style3">
	<% 
	response.write (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id"))) 
	countsumAllAnalis = countsumAllAnalis+(GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id")))
	%>
    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3" >รวม</td>
<td align="center" class="style3"><%=countsumcore%></td>
<td align="center" class="style3"><%=countsumAnalis%></td>
<td align="center" class="style3"><%=countsumAllAnalis%></td>
</tr>
 <tr  bgcolor="#FFFF99">
<td  align="left" class="style3">กองสนับสนุน</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
<td align="center">&nbsp;</td>
</tr>
 <%
  countsumcore=0
  countsumAnalis=0
  countsumAllAnalis=0
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='1' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
    <td align="center" class="style3">
	<% 

			response.write GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
			countsumcore = countsumcore+ GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")
	%>
    </td>
    <%
	if (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'") = 0 and getCountRowAnalis(RecDepart("D_Id"))=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (getCountRowAnalis(RecDepart("D_Id"))<>0 and (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id")) <> 0) )  then
		colorSet=" bgcolor=""#FFFF00"""
	else
		 colorSet = "bgcolor=""#99FF33"""
	end if  
	%>
    <td align="center" class="style3" <%=colorSet%>>
	<% 
			response.write getCountRowAnalis(RecDepart("D_Id"))
			countsumAnalis = countsumAnalis+getCountRowAnalis(RecDepart("D_Id"))
	%>
    </td>
    <td align="center">
    <% 
	response.write (GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id"))) 
	countsumAllAnalis = countsumAllAnalis+(GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&RecDepart("D_Id")&"'")-getCountRowAnalis(RecDepart("D_Id")))
	%>
    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3">รวม</td>
<td align="center" class="style3"><%=countsumcore%></td>
<td align="center" class="style3"><%=countsumAnalis%></td>
<td align="center" class="style3"><%=countsumAllAnalis%></td>
</tr>
</table>
<!--End show core and support-->
<br />
<div class="style1" align="center" style="font-size:18px">รายงานการวิเคราะห์ความสอดคล้องและความต้องการของผู้รับบริการ</div><br />
<div class="style1" align="center" style="font-size:18px">
<%
	  Set   rec_jobresult_group = Server.CreateObject("ADODB.RECORDSET")
	  sql_jobresult_group = "select  *  from  Tb_Department order by D_Numberlist  asc"
	  rec_jobresult_group.open sql_jobresult_group,ConQS,1,3
	  %>
	  <select name="JobresultGroupId" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:18px"  >
      <option value="0"  <% if getDid = "0" then  response.write "seledted=""selected"" " end if %>>เลือกหน่วยงาน</option>
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
</div><br />
<div>
<% 
countsumcore = GetCountRowQS("Tb_Manual","M_Id"," where M_Main=1 and D_Id='"&getDid&"'")
response.write getDepartmentname(getDid)&" กระบวนการหลักทั้งหมด : "&countsumcore&" กระบวนการ" %>
</div><br />
<table width="100%"  cellspacing="0" cellpadding="0" border="1">
  <tr>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">ลำดับที่</td>
   <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">วันที่</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">ชื่อกระบวนงาน<br />
    (Procedure)</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">ประเด็นยุทธ์</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">กลยุทธ์</td>
    <td colspan="5" align="center" bgcolor="#CCCCCC" class="style1">เกณฑ์การประเมิน (ระบุคะแนน)</td>
    <td colspan="9" align="center" bgcolor="#CCCCCC" class="style1">ความต้องการของผู้รับบริการ<br />
    (ใส่เครื่องหมาย / โดยตอบได้มากกว่า 1 ข้อ)</td>
  </tr>
  <tr>
    <td  align="center" bgcolor="#CCCCCC" class="style1">สนับสนุน<br />
    พันธกิจ<br />
    ขององค์กร</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ลดระยะเวลา<br />
    ในการทำงาน</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">เพิ่มคุณภาพ<br />
    การให้บริการ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ลดค่าใช้จ่าย<br />
    ในการทำงาน</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">คะแนนรวม</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ความคุ้มค่า</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ความเป็นธรรม</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ความถูกต้อง</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ความโปรงใส</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">การมีส่วนร่วม</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">มีผู้รับผิดชอบ<br />
    ชัดเจน</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ความสะดวก<br />
    รวดเร็ว</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">สนับสนุน<br />
    การปฏิบัติงาน</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">อื่นๆ<br />
    (โปรดระบุ)</td>
  </tr>
  <%
  		'dim getM_Id
		'getM_Id = Request.Form("hidMID")
		'SQL = "select * from Tb_AnalisProcedure where M_Id="&getM_Id
		dim showArrayStategic 
		dim getDDArray(3)
		dim getDArray(16)
		dim showDArray(16)
		dim CountRowSOP
		CountRowSOP=1
		SQL = "select Tb_AnalisProcedure.Analis_ID as Analis_ID,Tb_AnalisProcedure.M_Id as M_Id ,Tb_AnalisProcedure.M_Code as M_code,Tb_AnalisProcedure.M_Name as M_Name,Tb_AnalisProcedure.Analis_Strategic as Analis_Strategic,Tb_AnalisProcedure.Analis_Strategy as Analis_Strategy,Tb_AnalisProcedure.Analis_Support as Analis_Support,Tb_AnalisProcedure.Analis_period as Analis_period,Tb_AnalisProcedure.Analis_Quality as Analis_Quality,Tb_AnalisProcedure.Analis_Charge as Analis_Charge,Tb_AnalisProcedure.Analis_Sum as Analis_Sum,Tb_AnalisProcedure.Analis_value as Analis_value ,Tb_AnalisProcedure.Analis_Fairness as Analis_Fairness,Tb_AnalisProcedure.Analis_Accuracy as Analis_Accuracy,Tb_AnalisProcedure.Analis_Transparency as Analis_Transparency,Tb_AnalisProcedure.Analis_Participation as Analis_Participation,Tb_AnalisProcedure.Analis_Response as Analis_Response,Tb_AnalisProcedure.Analis_Ease as Analis_Ease,Tb_AnalisProcedure.Analis_Worksupport as Analis_Worksupport,Tb_AnalisProcedure.Analis_Else as Analis_Else,Tb_AnalisProcedure.Analis_DesElse as Analis_DesElse,Tb_AnalisProcedure.Analis_Date as Analis_Date from (Tb_AnalisProcedure inner join Tb_Manual on Tb_AnalisProcedure.M_Id=Tb_Manual.M_Id) inner join Tb_Department on Tb_Manual.D_Id = Tb_Department.D_Id where Tb_Department.D_Id='"&getDid&"' and Tb_Manual.M_Main=1 and  Tb_Manual.M_Reserve=0"
		'response.write SQL
		set RecAnalis = Server.CreateObject("ADODB.RECORDSET")
		RecAnalis.open SQL,ConQS,1,3
		 while not RecAnalis.EOF
		 getAnalis_ID = RecAnalis("Analis_ID")
		 getAnallis_M_Id = RecAnalis("M_Id")
		 getAnalis_M_Code = RecAnalis("M_Code")
		 getAnalis_M_Name = RecAnalis("M_Name")
		 getAnalis_Strategic = RecAnalis("Analis_Strategic")
		 getAnalis_Strategy = RecAnalis("Analis_Strategy")
		 getAnalis_Support = RecAnalis("Analis_Support")
		 getAnalis_period = RecAnalis("Analis_period")
		 getAnalis_Quality = RecAnalis("Analis_Quality")
		 getAnalis_Charge = RecAnalis("Analis_Charge")
		 getAnalis_Sum = RecAnalis("Analis_Sum")
		 getAnalis_value = RecAnalis("Analis_value")
		 getAnalis_Fairness = RecAnalis("Analis_Fairness")
		 getAnalis_Accuracy = RecAnalis("Analis_Accuracy")
		 getAnalis_Transparency = RecAnalis("Analis_Transparency")
		 getAnalis_Participation = RecAnalis("Analis_Participation")
		 getAnalis_Response = RecAnalis("Analis_Response")
		 getAnalis_Ease = RecAnalis("Analis_Ease")
		 getAnalis_Worksupport = RecAnalis("Analis_Worksupport")
		 getAnalis_Else = RecAnalis("Analis_Else")
		 getAnalis_DesElse = RecAnalis("Analis_DesElse")
		 getAnalis_Date = RecAnalis("Analis_Date")
		 
		 
		  '----------------------------------start block separate strategic value--------------------------
		  		'dim showArrayStategic 
				'dim getDDArray(3)
				PPStart = 1
				Pcount=0
				PcountCom=0
				showArrayStrategic=""
				for i=0 to 2
					getDDArray(i) = Mid(getAnalis_Strategic,PPStart,1)
					PPStart = PPStart+2
					if getDDArray(i) <> "0" then
					Pcount=Pcount+1
					end if
				next
				
				for i=0 to 2
					if getDDArray(i)  <> "0" then
					
						PcountCom = PcountCom+1
						if  PcountCom < Pcount   then
							showArrayStrategic  = showArrayStrategic&getDDArray(i)&","
							'showArrayStrategy = Mid(showArrayStrategy,1,len(showArrayStrategy)-1)
							'response.write showArrayStrategic&"/<br>"
						else
							showArrayStrategic  = showArrayStrategic&getDDArray(i)
							'response.write showArrayStrategic&"\<br>"
						end if
						
'							if i = 0 then
'								showArrayStrategic  = showArrayStrategic&getDDArray(i)
'							else
'									if i  < 2 and Pcount <> 1 then
'										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
'										'response.write "1"
'									elseif i = 1 and getDDArray(i) <> "0" and Pcount <> 1 then
'										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
'										'response.write "2"
'									elseif i = 2 and getDDArray(i) <> "0" and Pcount <> 1 then
'										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
'										'response.write "3"
'									else
'										showArrayStrategic  = showArrayStrategic&getDDArray(i)
'										'response.write "4"
'									end if
'							end if
					end if
			 	next
				'response.write showArrayStrategic
		  '----------------------------------start block separate strategic value--------------------------
		  '----------------------------------start block separate strategy value--------------------------
		 	showArrayStrategy=""
			PStart = 1
		 	'dim getDArray(16)
			'dim showDArray(16)
			showDArray(0)="1.1"
			showDArray(1)="1.2"
			showDArray(2)="1.3"
			showDArray(3)="1.4"
			showDArray(4)="1.5"
			showDArray(5)="1.6"
			showDArray(6)="1.7"
			showDArray(7)="2.1"
			showDArray(8)="2.2"
			showDArray(9)="2.3"
			showDArray(10)="2.4"
			showDArray(11)="3.1"
			showDArray(12)="3.2"
			showDArray(13)="3.3"
			showDArray(14)="3.4"
			showDArray(15)="3.5"
		 	dim showArrayStrategy
			countVal=0
			countValCom=0
			 for i=0 to 15
				getDArray(i) = Mid(getAnalis_Strategy,PStart,2)
				if getDArray(i) <> "00" then
				countVal = countVal+1
				end if
				PStart = PStart+3
				'response.write getDArray(i)&"<br>"
			 next
			 'response.write countVal&"<br>"
			for i=0 to 15
				'response.write getDArray(i)&"<br>"
				if getDArray(i)  <> "00" then
						countValCom = countValCom+1
						if  countValCom < countVal   then
							showArrayStrategy  = showArrayStrategy&showDArray(i)&","
							'showArrayStrategy = Mid(showArrayStrategy,1,len(showArrayStrategy)-1)
							'response.write showArrayStrategy&"/<br>"
						else
							showArrayStrategy  = showArrayStrategy&showDArray(i)
							'response.write showArrayStrategy&"\<br>"
						end if 
				end if
			 next
		 '----------------------------------end block separate strategy value--------------------------
		 
  %>
  <tr>
    <td align="center" class="style1" height="25"><%=CountRowSOP%></td>
    <td align="center" class="style1"><%=getAnalis_Date%></td>
    <td align="left" class="style1"><%=getAnalis_M_Code&" "&getAnalis_M_Name%></td>
    <td align="center" class="style1"><%=showArrayStrategic%></td>
    <td align="center" class="style1"><%=showArrayStrategy%></td>
    <td align="center" class="style1"><%=getAnalis_Support%></td>
    <td align="center" class="style1"><%=getAnalis_Period%></td>
    <td align="center" class="style1"><%=getAnalis_Quality%></td>
    <td align="center" class="style1"><%=getAnalis_Charge%></td>
    <td align="center" class="style1"><%=getAnalis_Sum%></td>
    <td align="center" class="style1"><% if getAnalis_value = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Fairness = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Accuracy = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Transparency = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Participation = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Response = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Ease = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Worksupport = True then %><strong>&#47;</strong><% else %>&nbsp;<% end if%></td>
    <td align="center" class="style1"><% if getAnalis_Else = True then %><strong><%=getAnalis_DesElse%></strong><% else  response.write "ไม่มีความเห็น" end if%></td>
  </tr>
  <%
  CountRowSOP = CountRowSOP+1
  RecAnalis.MoveNext()
  wend
  %>
  <% if RecAnalis.RecordCount = 0 then %>
  <tr><td colspan="19" align="center">No Data</td></tr>
  <% end if %>
</table>
<br />
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/>&nbsp;&nbsp;</div>
</body>
</html>
