<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
getID = Request.QueryString("ID")
getDID = Request.QueryString("DID")
if isEmpty(getDID) = true then
getDID = 1
end if
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
if isEmpty(session("member")) = False and isEmpty(getID) = False  then
 set RecView = Server.CreateObject("ADODB.RECORDSET")
 sql = "select * from Tb_ManagementReview where MR_ID="&getID&" and  Flag_Show = True"
 RecView.open sql,conQS,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>����§ҹ</title>
<style>
.text {
					Font-size:14px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:14px}
.textsmall {
					Font-size:10px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:12px}
.textbig {
					Font-size:20px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:20px}
</style>
</head>

<body bgcolor="#ffffff">
<table width="75%" border="2" align="center" cellpadding="10" cellspacing="0" bordercolor="#333333">
  <tr>
    <td  colspan="2" class="text" align="center">
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td align="center" class="textbig">��¡�û�Ъ�����ǹ�½��º�����
            <% if RecView("MR_Level") = 1 then %>
�дѺ���
<% elseif RecView("MR_Level") = 2 then %>
�дѺ�ͧ
<% end if %></td>
        </tr>
        <tr>
          <td align="center" class="textbig"><%=getDepartmentname(RecView("D_Id"))%></td>
        </tr>
        <tr>
          <td align="center" class="textbig">���駷�� <%=RecView("MR_CountMeeting")%>&nbsp;&nbsp;&nbsp;�ѹ��� <%=DAy(RecView("MR_Date"))%>&nbsp;&nbsp;<%=thmonthFull(Month(RecView("MR_DAte")))%>&nbsp;&nbsp;<%=Year(RecView("MR_Date"))+543%></td>
        </tr>
        
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>      </td>
  </tr>
  
  <tr>
    <td width="35%" height="50" class="textbig" align="center">��Ǣ��</td>
    <td width="65%" class="textbig" align="center">��</td>
  </tr>
  <tr>
    <td  class="text">1.&nbsp;�š�õ�Ǩ�Դ����س�Ҿ</td>
    <td class="text"><% if RecView("MR_Review1") <> "" then response.write RecView("MR_Review1") else response.write "&nbsp;" end if %></td>
  </tr>
  
  <tr>
    <td  class="text">2. �����Դ��繢ͧ����Ѻ��ԡ��</td>
    <td  class="text"><% if RecView("MR_Review2") <> "" then response.write RecView("MR_Review2") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td  class="text">3. �š�ô��Թ��õ����͡�˹��к��س�Ҿ</td>
    <td  class="text"><% if RecView("MR_Review3") <> "" then response.write RecView("MR_Review3") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td  class="text">4. ʶҹТͧ��û�Ժѵԡ�����/��ͧ�ѹ</td>
    <td  class="text"><% if RecView("MR_Review4") <> "" then response.write RecView("MR_Review4") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td  class="text"><p>5. ��õԴ����Ũҡ��û�Ъ�����ǹ</p>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;�½��º����ä��駡�͹</p></td>
    <td  class="text"><% if RecView("MR_Review5") <> "" then response.write RecView("MR_Review5") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td  class="text">6. �������¹�ŧ����Ҩ�ռš�з�����к��س�Ҿ</td>
    <td  class="text"><% if RecView("MR_Review6") <> "" then response.write RecView("MR_Review6") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td  class="text">7. ����ʹ�������Ѻ��û�Ѻ��ا</td>
    <td class="text"><% if RecView("MR_Review7") <> "" then response.write RecView("MR_Review7") else response.write "&nbsp;" end if %></td>
  </tr>
  <tr>
    <td colspan="2" align="center">&nbsp;</td>
  </tr>
</table>
<br />
<div align="center"><label>
      <input type="button" name="butSave" id="butSave" value="������͡���"  onclick="javascript:{ window.print();}" />
      &nbsp;&nbsp;
      <input type="button" name="butCancel" id="butCancel" value="��Ѻ" onClick="javascript:{ window.location.href='ManagementReview.asp?ID=<%=getDID%>';}" />
    </label></div>
<%
end if 
%>
</body>
</html>
