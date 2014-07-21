<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
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
	 	getDid = "0"
	 end if
else
	getDid=Request.QueryString("id")
end if

if isEmpty(Request.QueryString("oid")) = true then
	 if isEmpty(Request.Form("hidOid")) = false then
	 	getOid=Request.Form("hidOid")
	 else
	 	getOid = "0"
	 end if
else
	getOid=Request.QueryString("oid")
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
	getTid = "0"
	FlagMain_Reserve = "M_Main=1"
	chkPC = "checked=""checked"""
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>รายงานวิเคราะห์การทบทวนเอกสารคุณภาพ</title>
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
		
		
		if ((val != "" ) || (val1 != ""))
		{ 
			
			var typeID = document.getElementById("TypeDoc").value;
			//alert("1/"+typeID);
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			var typ = document.getElementById("TypeDoc").value;
			
			window.location.href="ReviewReport.asp?id="+strUser+"&oid="+typeID+"&tid="+typeID;
			//window.location.href="ReviewReport.asp?id="+val+"&oid="+val1+"&tid="+typeID;
		}else{
			
			var typeID = document.getElementById("TypeDoc").value;
			//alert("2/"+typeID);
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			var typ = document.getElementById("TypeDoc").value;
			
			window.location.href="ReviewReport.asp?id="+strUser+"&oid="+typeID+"&tid="+typeID;
		}
		
}
</script>
</head>

<body>
<!------------------------------------------------------------Start code table show details------------------------------------------------------------------>
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#999999">
    <td width="25%" align="center" class="style3">สำนัก / กอง / กลุ่ม</td>
    <td width="20%" colspan="3" align="center" class="style3">คู่มือคุณภาพ (Q)</td>
    <td width="20%" colspan="3" align="center" class="style3">คู่มือขั้นตอนการปฏิบัติงาน (P)</td>
    <td width="20%" colspan="3" align="center" class="style3">คู่มือขั้นตอนวิธีปฏิบัติงาน (W)</td>
  </tr>
  <tr bgcolor="#FFFF99">
<td  align="left" class="style3" >กองผลิตภัณฑ์</td>
<td width="8%" align="center">เอกสารทั้งหมด</td>
<td width="8%" align="center">ทบทวนไปแล้ว</td>
<td width="9%" align="center">คงเหลือ</td>
<td width="8%" align="center">เอกสารทั้งหมด</td>
<td width="8%" align="center">ทบทวนไปแล้ว</td>
<td width="9%" align="center">คงเหลือ</td>
<td width="8%" align="center">เอกสารทั้งหมด</td>
<td width="8%" align="center">ทบทวนไปแล้ว</td>
<td width="9%" align="center">คงเหลือ</td>
  </tr>
  <%
  dim  countsumPAll,countsumPReview,countsumPRemaining
  dim countsumWAll,countsumWReview,countsumWRemaining
  dim countsumQAll,countsumQReview,countsumQRemaining
  dim colorSet
  
  countsumPAll=0
  countsumPReview=0
  countsumPRemaining=0
  
  countsumWAll=0
  countsumWReview=0
  countsumWRemaining=0
  
  countsumQAll=0
  countsumQReview=0
  countsumQRemaining=0
  
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='0' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
   <%
	if (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ") < GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") = GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ") and GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if    
	%>
    <td align="center" class="style3">
	<% 
	response.write (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")) 
	countsumQAll = countsumQAll+(GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'"))
	%>    </td>
    <td align="center" class="style3" <%=colorSet%>>
    <%
	response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")
	countsumQReview = countsumQReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")
	%>    </td>
    <td align="center" class="style3">
    <%
	response.write (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' "))
	countsumQRemaining = countsumQRemaining+ (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' "))
	%>    </td>
        <%
	if (GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ") < GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'")=GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ") and GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if  
	%>
    <td align="center" class="style3" >	
			<% 
			response.write GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'") 
			countsumPAll = countsumPAll+ GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'")
			%>    </td>
    <td align="center" class="style3" <%=colorSet%>>
	<% 
	response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")
	countsumPReview = countsumPReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")
	%>    </td>
    <td align="center" class="style3">
	<% 
	response.write (GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")) 
	countsumPRemaining = countsumPRemaining+(GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' "))
	%>    </td>
<%
	if (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ") < GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif  GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") = GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ") and GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if   
	%>
    <td align="center" class="style3" >
	<% 
	response.write GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")
	countsumWAll = countsumWAll+GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")
	%></td>
    <td align="center" class="style3" <%=colorSet%>>
    <%
	response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")
	countsumWReview = countsumWReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")
	%>    </td>
    <td align="center" class="style3" >
    <%
	response.write (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' "))
	countsumWRemaining = countsumWRemaining+ (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' "))
	%>    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3" >รวม</td>
<td align="center" class="style3"><%=countsumQAll%></td>
<td align="center" class="style3"><%=countsumQReview%></td>
<td align="center" class="style3"><%=countsumQRemaining%></td>
<td align="center" class="style3"><%=countsumPAll%></td>
<td align="center" class="style3"><%=countsumPReview%></td>
<td align="center" class="style3"><%=countsumPRemaining%></td>
<td align="center" class="style3"><%=countsumWAll%></td>
<td align="center" class="style3"><%=countsumWReview%></td>
<td align="center" class="style3"><%=countsumWRemaining%></td>

</tr>
 <tr  bgcolor="#FFFF99">
<td  align="left" class="style3">กองสนับสนุน</td>
<td colspan="3" align="center">&nbsp;</td>
<td colspan="3" align="center">&nbsp;</td>
<td colspan="3" align="center">&nbsp;</td>
</tr>
 <%
 countsumPAll=0
  countsumPReview=0
  countsumPRemaining=0
  
  countsumWAll=0
  countsumWReview=0
  countsumWRemaining=0
  
  countsumQAll=0
  countsumQReview=0
  countsumQRemaining=0
  
  set RecDepart = Server.CreateObject("ADODB.RECORDSET")
  sqlDepart = "select * from Tb_Department where D_Type='1' order by D_Numberlist ASC "
  RecDepart.open sqlDepart,ConQS,1,3
  while not RecDepart.EOF 
%>
  <tr>
    <td class="style3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= getDepartmentname(RecDepart("D_Id"))%></td>
     <%
	if (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ") < GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") = GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ") and GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if    
	%>
    <td align="center" class="style3">
			<% 
            response.write (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")) 
            countsumQAll = countsumQAll+(GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'"))
            %>    
    </td>
    <td align="center" class="style3" <%=colorSet%>>
			<%
            response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")
            countsumQReview = countsumQReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' ")
            %>    
    </td>
    <td align="center" class="style3">
			<%
            response.write (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' "))
            countsumQRemaining = countsumQRemaining+ (GetCountRowQS("Tb_QM","QM_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'Q' "))
            %>    
    </td>
    <%
	if (GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ") < GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'")=GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ") and GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if  
	%>
    <td align="center" class="style3" >	
			<% 
			response.write GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'") 
			countsumPAll = countsumPAll+ GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&RecDepart("D_Id")&"'")
			%>    
    </td>
    <td align="center" class="style3" <%=colorSet%>>
			<% 
            response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")
            countsumPReview = countsumPReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")
            %>    
    </td>
    <td align="center" class="style3">
			<% 
            response.write (GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")) 
            countsumPRemaining = countsumPRemaining+(GetCountRowQS("Tb_Manual","M_Id"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' "))
            %>    
    </td>
   <%
	if (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 and GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")=0)  then
		colorSet=" bgcolor=""#FE4541"""
	elseif (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")>0 and (GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ") < GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") ) )  then
		colorSet=" bgcolor=""#FFFF00"""
	elseif  GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") = GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ") and GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'") > 0 then
		 colorSet = "bgcolor=""#99FF33"""
	else
		colorSet = ""
	end if   
	%>
    <td align="center" class="style3" >
			<% 
            response.write GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")
            countsumWAll = countsumWAll+GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")
            %>
    </td>
    <td align="center" class="style3" <%=colorSet%>>
			<%
            response.write GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")
            countsumWReview = countsumWReview+GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' ")
            %>    
    </td>
    <td align="center" class="style3" >
			<%
            response.write (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' "))
            countsumWRemaining = countsumWRemaining+ (GetCountRowQS("Tb_Workin","W_ID"," where  D_Id='"&RecDepart("D_Id")&"'")-GetCountRowQS("Tb_Review","R_Id"," where  D_Id='"&RecDepart("D_Id")&"' and Type_Sop = 'W' "))
            %>    
    </td>
  </tr>
<%
RecDepart.MoveNext
Wend
%>
<tr bgcolor="#CCCCCC">
<td  align="center" class="style3" >รวม</td>
<td align="center" class="style3"><%=countsumQAll%></td>
<td align="center" class="style3"><%=countsumQReview%></td>
<td align="center" class="style3"><%=countsumQRemaining%></td>
<td align="center" class="style3"><%=countsumPAll%></td>
<td align="center" class="style3"><%=countsumPReview%></td>
<td align="center" class="style3"><%=countsumPRemaining%></td>
<td align="center" class="style3"><%=countsumWAll%></td>
<td align="center" class="style3"><%=countsumWReview%></td>
<td align="center" class="style3"><%=countsumWRemaining%></td>

</tr>
</table>
<!------------------------------------------------------------End code table show details------------------------------------------------------------------->
<div class="style1" align="center" style="font-size:18px">รายงานวิเคราะห์การทบทวนเอกสารคุณภาพ</div><br />
<div class="style1" align="center" style="font-size:18px">
<%
	  Set   rec_jobresult_group = Server.CreateObject("ADODB.RECORDSET")
	  sql_jobresult_group = "select  *  from  Tb_Department order by D_Numberlist  asc"
	  rec_jobresult_group.open sql_jobresult_group,ConQS,1,3
	  %>
	  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroup(this.value,'')" style="font-size:16px"  >
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
<div class="style1" align="center" style="font-size:18px">
	  <select name="TypeDoc" id="TypeDoc" onChange="ChangeJobresultGroup('',this.value)" style="font-size:16px"  >
      <option value="0" <% if getOid ="0" then response.write " selected=""selected"" " end if%> >เลือกประเภทเอกสาร</option>
	  <option value="Q" <% if getOid ="Q" then response.write " selected=""selected"" " end if%> >คู่มือคุณภาพ (Q)</option>
      <option value="PC"  <% if getOid ="PC" then response.write " selected=""selected"" " end if%>>คู่มือขั้นตอนการปฏิบัติงาน (P)</option>
      <option value="W" <% if getOid ="W" then response.write " selected=""selected"" " end if%> >คู่มือขั้นตอนวิธีปฏิบัติงาน (W)</option>
      </select>
</div>
<br />
<div>
<%
if getTid ="PC" or getTid = "PS" then
countsumcore=0
countsumcore = GetCountRowQS("Tb_Manual","M_Id"," where D_Id='"&getDid&"'")
response.write " กระบวนการทั้งหมด  : "&countsumcore&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' ")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"' and Type_Sop <> 'Q' and Type_Sop <> 'W' "))&" กระบวนการ"
elseif  getTid = "W" then
countsumcore=0
countsumcore = GetCountRowQS("Tb_Workin","D_Id"," where D_Id='"&getDid&" ' ")
response.write " กระบวนการทั้งหมด : "&countsumcore&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&" '  and Type_Sop='W' ")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'   and Type_Sop='W' "))&" กระบวนการ"
elseif  getTid = "Q" then
countsumcore3=0
countsumcore3 = GetCountRowQS("Tb_QM","D_Id"," where D_Id='"&getDid&"' ")
response.write " กระบวนการทั้งหมด : "&countsumcore3&" กระบวนการ  มีการทบทวน : "&GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'  and Type_Sop='Q' ")&" กระบวนการ  คงเหลือต้องทบทวน : "&(countsumcore3-GetCountRowQS("Tb_Review","D_Id"," where D_Id='"&getDid&"'  and Type_Sop='Q' "))&" กระบวนการ"
end if
%>
</div>
<table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#666666">
  <tr>
    <td width="20%" rowspan="2"><div align="center">ชื่อสำนัก/กอง</div></td>
    <td width="10%" rowspan="2"><div align="center">รหัส</div></td>
    <td width="40%" rowspan="2"><div align="center">ชื่อเอกสาร</div></td>
    <td width="30%" colspan="5"><div align="center">ผลการทบทวนเอกสาร</div></td>
  </tr>
  <tr>
    <td><div align="center">เหมาะสม</div></td>
    <td><div align="center">ยกเลิก</div></td>
    <td><div align="center">แก้ไข</div></td>
    <td><div align="center">ทำใหม่</div></td>
    <td><div align="center">วันที่คาดว่าจะแล้วเสร็จ</div></td>
  </tr>
  <%
   Set   recSOP = Server.CreateObject("ADODB.RECORDSET")
	  if getTid = "PC" or getTid = "PS" then
	  	sqlSOP = "select  *  from  Tb_Review where  D_Id='"&getDid&"' and  Type_Sop in ('PC','PS') order by  R_Id DESC"
	  elseif getTid = "W" then
	   sqlSOP = "select  *  from  Tb_Review where  D_Id='"&getDid&"' and Type_Sop='W'  order by  R_Id DESC"
	  elseif getTid = "Q" then
	  	sqlSOP = "select  *  from  Tb_Review where  D_Id='"&getDid&"' and Type_Sop='Q'  order by  R_Id DESC"
	  else
	  sqlSOP = "select  *  from  Tb_Review where  D_Id='"&getDid&"' and  Type_Sop='"&getTid&"' order by  R_Id DESC"
	  end if
	  'response.write sqlSOP 
	  recSOP.open sqlSOP,ConQS,1,3
	  While not recSOP.EOF
  %>
  <tr>
    <td><%=getDepartmentname(recSOP("D_Id"))%></td>
    <td><%=recSOP("M_Code")%></td>
    <td><%=recSOP("M_Name")%></td>
    <td align="center">
	<%
	if recSOP("Comport") = true then
	 response.write "&#149;"
	else
	response.write "&nbsp;"
	end if
	%></td>
    <td align="center">
    <%
	if recSOP("Comport") = false then 
		if recSOP("MethodType") = 3 then
			response.write "&#149;"
		else
			response.write "&nbsp;"
		end if
	else
		response.write "&nbsp;"		
	end if
	%>
    </td>
    <td align="center">
     <%
	if recSOP("Comport") = false then 
		if recSOP("MethodType") = 2 then
			response.write "&#149;"
		else
			response.write "&nbsp;"
		end if
	else
		response.write "&nbsp;"		
	end if
	%>
    </td>
    <td align="center">
    <%
	if recSOP("Comport") = false then 
		if recSOP("MethodType") = 1 then
			response.write "&#149;"
		else
			response.write "&nbsp;"
		end if
	else
		response.write "&nbsp;"		
	end if
	%>
    </td>
    <td align="center">
    <%
	 if recSOP("Comport") = false then
		 if recSOP("MethodType") = 1 then
		 	response.write  recSOP("Remake_Date")
		 elseif recSOP("MethodType") = 2 then
		 	 response.write  recSOP("Edit_Date")
		 else
		 	response.write "&nbsp;"		
		 end if
	 
	 else
	 	response.write "&nbsp;"
	 end if
	%>
    </td>
  </tr>
  <%
  recSOP.MoveNext
  Wend
  %>
  <% if recSOP.RecordCount = 0 then %>
  <tr><td colspan="8" align="center">No Data</td></tr>
  <% end if %>
</table>
</body>
</html>
