<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<%
ID_L2=request("ID_L2")
ID_L3=request("ID_L3")
ID_L4=request("ID_L4")

if ID_L2<>"" Then call OpenRecord(rs,"Select * From TabData_L2 Where Id_L2="&ID_L2,con,1,1)
if ID_L3<>"" Then call OpenRecord(rs,"Select * From TabData_L3 Where Id_L3="&ID_L3,con,1,1)
if ID_L4<>"" Then call OpenRecord(rs,"Select * From TabData_L4 Where Id_L4="&ID_L4,con,1,1)%>
<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">



<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><b><%=rs("Topic")%></b></td>
  </tr>
  <tr> 
    <td class="FontEditor" style="font-size:14px" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�к��س�Ҿ (Quality System) ���¶֧ �к����������ͧ���㹡�äǺ�����л�Сѹ�س�Ҿ�ͧ˹��§ҹ ��觻�Сͺ仴����ç���ҧ�ͧͧ��� ˹�ҷ������Ѻ�Դ�ͺ �Ըմ��Թ��� ��кǹ��ô��Թ��� ��Ѿ�ҡ� ���͹ӹ�º�¡�ú����çҹ��ҹ�س�Ҿ任�Ժѵ� ��ô��Թ��ôѧ����Ǩ��繵�ͧ�Ѵ�����͡��� ��������ö���Թ����ѡ���к��س�Ҿ�����ҧ������� �������ö����������ҧ�ջ���Է���Ҿ</td>
  </tr>
  <tr> 
    <td class="FontEditor" ><%=rs("Desc")%></td>
  </tr>
  <tr>
    <td >&nbsp;</td>
  </tr>
  <tr> 
    <td ><a href="javascript:history.back()"><img src="<%=path_link%>_images/i.p.prevpg.gif" border=0 align="absmiddle">&nbsp;��Ѻ</a></td>
  </tr>
</table>
