<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->

<%
	if isEmpty(session("member")) = True then
		Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
	end if
		dim getM_Id
		getM_Id = Request.Form("hidMID")
		getM_Id=331
'		SQL = "select * from Tb_AnalisProcedure where M_Id="&getM_Id
'		set RecAnalis = Server.CreateObject("ADODB.RECORDSET")
'		RecAnalis.open SQL,ConQS,1,3
'		 while not RecAnalis.EOF
'		 getAnalis_ID = RecAnalis("Analis_ID")
'		 getAnallis_M_Id = RecAnalis("M_Id")
'		 getAnalis_M_Code = RecAnalis("M_Code")
'		 getAnalis_M_Name = RecAnalis("M_Name")
'		 getAnalis_Strategic = RecAnalis("Analis_Strategic")
'		 getAnalis_Strategy = RecAnalis("Analis_Strategy")
'		 getAnalis_Support = RecAnalis("Analis_Support")
'		 getAnalis_period = RecAnalis("Analis_period")
'		 getAnalis_Quality = RecAnalis("Analis_Quality")
'		 getAnalis_Charge = RecAnalis("Analis_Charge")
'		 getAnalis_Sum = RecAnalis("Analis_Sum")
'		 getAnalis_value = RecAnalis("Analis_value")
'		 getAnalis_Fairness = RecAnalis("Analis_Fairness")
'		 getAnalis_Accuracy = RecAnalis("Analis_Accuracy")
'		 getAnalis_Transparency = RecAnalis("Analis_Transparency")
'		 getAnalis_Participation = RecAnalis("Analis_Participation")
'		 getAnalis_Response = RecAnalis("Analis_Response")
'		 getAnalis_Ease = RecAnalis("Analis_Ease")
'		 getAnalis_Worksupport = RecAnalis("Analis_Worksupport")
'		 getAnalis_Else = RecAnalis("Analis_Else")
'		 getAnalis_DesElse = RecAnalis("Analis_DesElse")
'		 getAnalis_Date = RecAnalis("Analis_Date")
'		 RecAnalis.MoveNext()
'		 wend
'		  '----------------------------------start block separate strategic value--------------------------
'		  		dim showArrayStategic 
'				dim getDDArray(3)
'				PPStart = 1
'				Pcount=0
'				PcountCom=0
'				
'				for i=0 to 2
'					getDDArray(i) = Mid(getAnalis_Strategic,PPStart,1)
'					PPStart = PPStart+2
'					if getDDArray(i) <> "0" then
'					Pcount=Pcount+1
'					end if
'				next
'				
'				for i=0 to 2
'					if getDDArray(i)  <> "0" then
'					
'						PcountCom = PcountCom+1
'						if  PcountCom < Pcount   then
'							showArrayStrategic  = showArrayStrategic&getDDArray(i)&","
'							'showArrayStrategy = Mid(showArrayStrategy,1,len(showArrayStrategy)-1)
'							'response.write showArrayStrategy&"/<br>"
'						else
'							showArrayStrategic  = showArrayStrategic&getDDArray(i)
'							'response.write showArrayStrategy&"\<br>"
'						end if
'						
''							if i = 0 then
''								showArrayStrategic  = showArrayStrategic&getDDArray(i)
''							else
''									if i  < 2 and Pcount <> 1 then
''										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
''										'response.write "1"
''									elseif i = 1 and getDDArray(i) <> "0" and Pcount <> 1 then
''										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
''										'response.write "2"
''									elseif i = 2 and getDDArray(i) <> "0" and Pcount <> 1 then
''										showArrayStrategic  = showArrayStrategic&","&getDDArray(i)
''										'response.write "3"
''									else
''										showArrayStrategic  = showArrayStrategic&getDDArray(i)
''										'response.write "4"
''									end if
''							end if
'					end if
'			 	next
		  '----------------------------------start block separate strategic value--------------------------
		  '----------------------------------start block separate strategy value--------------------------
'		 	PStart = 1
'		 	dim getDArray(16)
'			dim showDArray(16)
'			showDArray(0)="1.1"
'			showDArray(1)="1.2"
'			showDArray(2)="1.3"
'			showDArray(3)="1.4"
'			showDArray(4)="1.5"
'			showDArray(5)="1.6"
'			showDArray(6)="1.7"
'			showDArray(7)="2.1"
'			showDArray(8)="2.2"
'			showDArray(9)="2.3"
'			showDArray(10)="2.4"
'			showDArray(11)="3.1"
'			showDArray(12)="3.2"
'			showDArray(13)="3.3"
'			showDArray(14)="3.4"
'			showDArray(15)="3.5"
'		 	dim showArrayStrategy
'			countVal=0
'			countValCom=0
'			 for i=0 to 15
'				getDArray(i) = Mid(getAnalis_Strategy,PStart,2)
'				if getDArray(i) <> "00" then
'				countVal = countVal+1
'				end if
'				PStart = PStart+3
'				'response.write getDArray(i)&"<br>"
'			 next
'			 'response.write countVal&"<br>"
'			for i=0 to 15
'				'response.write getDArray(i)&"<br>"
'				if getDArray(i)  <> "00" then
'						countValCom = countValCom+1
'						if  countValCom < countVal   then
'							showArrayStrategy  = showArrayStrategy&showDArray(i)&","
'							'showArrayStrategy = Mid(showArrayStrategy,1,len(showArrayStrategy)-1)
'							'response.write showArrayStrategy&"/<br>"
'						else
'							showArrayStrategy  = showArrayStrategy&showDArray(i)
'							'response.write showArrayStrategy&"\<br>"
'						end if 
'				end if
'			 next
		 '----------------------------------end block separate strategy value--------------------------
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>��§ҹ����������</title>
<style type="text/css">
<!--
.style1 {
font-size:10px;
font-family:Arial, Helvetica, sans-serif;
}
-->
</style>
</head>

<body>
<div class="style1" align="center" style="font-size:18px">Ẻ��������������������ʹ���ͧ�ͧ��кǹ�ҹ��ͤ�����ͧ��âͧ����Ѻ��ԡ��</div><br />
<div class="style1" align="center" style="font-size:18px">˹��§ҹ : <% response.write getJoinDepartmentname(getM_Id) %></div><br />
<table width="100%"  cellspacing="0" cellpadding="0" border="1">
  <tr>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">�ӴѺ���</td>
   <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">�ѹ���</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">���͡�кǹ�ҹ<br />
    (Procedure)</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">������ط��</td>
    <td rowspan="2" align="center" bgcolor="#CCCCCC" class="style1">���ط��</td>
    <td colspan="5" align="center" bgcolor="#CCCCCC" class="style1">ࡳ���û����Թ (�кؤ�ṹ)</td>
    <td colspan="9" align="center" bgcolor="#CCCCCC" class="style1">������ͧ��âͧ����Ѻ��ԡ��<br />
    (�������ͧ���� / �µͺ���ҡ���� 1 ���)</td>
  </tr>
  <tr>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ʹѺʹع<br />
    �ѹ��Ԩ<br />
    �ͧͧ���</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">Ŵ��������<br />
    㹡�÷ӧҹ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����س�Ҿ<br />
    �������ԡ��</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">Ŵ��������<br />
    㹡�÷ӧҹ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">��ṹ���</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����������</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����繸���</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����١��ͧ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����ç��</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�������ǹ����</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�ռ���Ѻ�Դ�ͺ<br />
    �Ѵਹ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">�����дǡ<br />
    �Ǵ����</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">ʹѺʹع<br />
    ��û�Ժѵԧҹ</td>
    <td  align="center" bgcolor="#CCCCCC" class="style1">����<br />
    (�ô�к�)</td>
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
		SQL = "select Tb_AnalisProcedure.Analis_ID as Analis_ID,Tb_AnalisProcedure.M_Id as M_Id ,Tb_AnalisProcedure.M_Code as M_code,Tb_AnalisProcedure.M_Name as M_Name,Tb_AnalisProcedure.Analis_Strategic as Analis_Strategic,Tb_AnalisProcedure.Analis_Strategy as Analis_Strategy,Tb_AnalisProcedure.Analis_Support as Analis_Support,Tb_AnalisProcedure.Analis_period as Analis_period,Tb_AnalisProcedure.Analis_Quality as Analis_Quality,Tb_AnalisProcedure.Analis_Charge as Analis_Charge,Tb_AnalisProcedure.Analis_Sum as Analis_Sum,Tb_AnalisProcedure.Analis_value as Analis_value ,Tb_AnalisProcedure.Analis_Fairness as Analis_Fairness,Tb_AnalisProcedure.Analis_Accuracy as Analis_Accuracy,Tb_AnalisProcedure.Analis_Transparency as Analis_Transparency,Tb_AnalisProcedure.Analis_Participation as Analis_Participation,Tb_AnalisProcedure.Analis_Response as Analis_Response,Tb_AnalisProcedure.Analis_Ease as Analis_Ease,Tb_AnalisProcedure.Analis_Worksupport as Analis_Worksupport,Tb_AnalisProcedure.Analis_Else as Analis_Else,Tb_AnalisProcedure.Analis_DesElse as Analis_DesElse,Tb_AnalisProcedure.Analis_Date as Analis_Date from (Tb_AnalisProcedure inner join Tb_Manual on Tb_AnalisProcedure.M_Id=Tb_Manual.M_Id) inner join Tb_Department on Tb_Manual.D_Id = Tb_Department.D_Id where Tb_Department.D_Id='"&getJoinDepartmentId(getM_Id)&"'"
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
						'response.write getDDArray(i)&"/"
						PcountCom = PcountCom+1
						if  PcountCom < Pcount   then
							showArrayStrategic  = showArrayStrategic&getDDArray(i)&","
							'showArrayStrategy = Mid(showArrayStrategy,1,len(showArrayStrategy)-1)
							response.write showArrayStrategic&"<br>"
						else
							showArrayStrategic  = showArrayStrategic&getDDArray(i)
							'response.write showArrayStrategy&"\<br>"
							response.write showArrayStrategic&"<br>"
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
    <td align="center" class="style1"><%=CountRowSOP%></td>
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
    <td align="center" class="style1"><% if getAnalis_Else = True then %><strong><%=getAnalis_DesElse%></strong><% else  response.write "����դ������" end if%></td>
  </tr>
  <%
  CountRowSOP = CountRowSOP+1
  RecAnalis.MoveNext()
  wend
  %>
</table>
<br />
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/>&nbsp;&nbsp;<input type="button"  value="��Ѻ˹�ҡ�͡������"  onclick="javascript:{ window.location.href='analaysis.asp';}"/></div>
</body>
</html>
