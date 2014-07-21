<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim getDid , getTid,RunRow
  getDid = Request.QueryString("Did")
  getTid = Request.QueryString("Tid")
  if  isEmpty(getDid) = true or isEmpty(getTid) = true then
  		response.write "<script  language=""javascript"">"
		response.write "window.close();"
		response.write "</script>"
  end if 
  RunRow=1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>Core Process</title>
<style type="text/css">
<!--
.style1 {
font-size:14px;
font-family:Arial, Helvetica, sans-serif;
color:#FFFFFF;
}
.style2 {
	font-size:22px;
	font-family:Arial, Helvetica, sans-serif;
	font-weight:600;
	color:#FFFFFF;
}
.style3 {
font-size:16px;
font-family:Arial, Helvetica, sans-serif;
font-weight:100;
color:#FFFFFF;
}
-->
</style>
</head>
<body bgcolor="#000000">
<table width="100%" border="2" cellpadding="2" cellspacing="0" bordercolor="#999999" bgcolor="#000000" align="center">
  <tr>
    <td colspan="3" class="style2"><%=getDepartmentname(getDid)%></td>
  </tr>
  <tr>
  <td width="2%"><div align="center" class="style1">ลำดับ</div></td>
    <td width="28%"><div align="center" class="style1">รหัสเอกสาร</div></td>
    <td width="70%"><div align="center" class="style1">ชื่อเอกสาร</div></td>
  </tr>
  <%
  set rec = Server.CreateObject("ADODB.RECORDSET")
  if getTid = "PC" then
  sql = "select * from Tb_Manual where D_Id='"&getDid&"'  and  M_Main=1 and M_Reserve=0 order by  M_Id asc "
  else
  sql = "select * from Tb_Manual where D_Id='"&getDid&"'  and  M_Main=0  and M_Reserve=1 order by  M_Id asc "
  end if
  rec.open sql,ConQS,1,3
  while not rec.eof
  %>
  <tr>
  	<td class="style3" align="center"><%=RunRow%></td>
    <td class="style3"><%=rec("M_Code")%></td>
    <td class="style3"><%=rec("M_Name")%></td>
  </tr>
  <%
  RunRow = RunRow+1
  rec.MoveNext
  Wend
  rec.Close()
  %>
</table>
</body>
</html>
