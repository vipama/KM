<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">
<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<div align="center"><span class="textdefault"> ***จำนวน Uni IP คือจำนวนเครื่องคอมพิวเตอร์ที่เข้ามาในหน้าเว็บ 
  โดยไม่ซ้ำกัน***<br>
  <br>
  <%call openrecord(rs,"Select Distinct  year([date]) as logyear from TabUniIP Order by year([date]) Desc",con,1,1)
  for i=1 to rs.recordcount%>
  </span> </div>
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC"  bgcolor="#FFFFCC" class="textdefault" style="empty-cells:show;border:1px">
  <tr> 
    <td colspan="15" bgcolor="#CCCCCC"><div align="center">ปี <%=CheckYear(rs("logyear"),"TH")%></div></td>
  </tr>
  <tr> 
    <td colspan="13" bgcolor="#00FFFF"><div align="center">จำนวน Uni IP</div></td>
  </tr>
  <tr bgcolor="#CCFF00"> 
    <td><div align="center">ม.ค.</div></td>
    <td><div align="center">ก.พ.</div></td>
    <td><div align="center">มี.ค.</div></td>
    <td><div align="center">เม.ย.</div></td>
    <td><div align="center">พ.ค.</div></td>
    <td><div align="center">มิ.ย.</div></td>
    <td><div align="center">ก.ค.</div></td>
    <td><div align="center">ส.ค.</div></td>
    <td><div align="center">ก.ย.</div></td>
    <td><div align="center">ต.ค.</div></td>
    <td><div align="center">พ.ย.</div></td>
    <td><div align="center">ธ.ค.</div></td>
    <td bgcolor="#B9C1FF"><div align="center"><b>รวม</b></div></td>
  </tr>
  
  <tr onMouseOver="this.style.backgroundColor='F3EE8D'" onMouseOut="this.style.backgroundColor=''"> 
    <%
	for numbermonth=1 to 12
	call openrecord(rs2,"Select Count(IP) as CountIP From TabUniIP  Where month(date)="&numbermonth,con,1,1)%>
    <td align="right" ><%=rs2("CountIP")%></td>
    <%totalCountIP=totalCountIP+rs2("CountIP")
	rs2.movenext
	next
	closerecord(rs2)%>
    <td align="right" ><b><%=totalCountIP%></b></td>
  </tr>
  <tr bgcolor="#99CCCC" > 

  </tr>
</table>
  <%rs.movenext
  next
  %>
  
<p>