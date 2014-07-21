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
    <td class="FontEditor" style="font-size:14px" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ระบบคุณภาพ (Quality System) หมายถึง ระบบที่เป็นเครื่องมือในการควบคุมและประกันคุณภาพของหน่วยงาน ซึ่งประกอบไปด้วยโครงสร้างขององค์กร หน้าที่ความรับผิดชอบ วิธีดำเนินการ กระบวนการดำเนินการ ทรัพยากร เพื่อนำนโยบายการบริหารงานด้านคุณภาพไปปฏิบัติ การดำเนินการดังกล่าวจำเป็นต้องจัดทำเป็นเอกสาร เพื่อสามารถดำเนินการรักษาระบบคุณภาพได้อย่างเหมาะสม และสามารถนำไปใช้ได้อย่างมีประสิทธิภาพ</td>
  </tr>
  <tr> 
    <td class="FontEditor" ><%=rs("Desc")%></td>
  </tr>
  <tr>
    <td >&nbsp;</td>
  </tr>
  <tr> 
    <td ><a href="javascript:history.back()"><img src="<%=path_link%>_images/i.p.prevpg.gif" border=0 align="absmiddle">&nbsp;กลับ</a></td>
  </tr>
</table>
