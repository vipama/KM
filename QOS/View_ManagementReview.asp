<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
' # start code for check permission in DB 
if Session("member") <> getPermission(session("member"),"L_Email") or isnull(session("member")) = true or session("member") = "" or isEmpty(session("member")) = True then
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
else
	session("Depart") = getPermission(session("member"),"D_Id")
end if
' # End code for check permission in DB
getID = Request.QueryString("ID")
getDID = Request.QueryString("DID")
getSource = Request.QueryString("Source")
if isEmpty(getDID) = true then
getDID = 1
end if
if isEmpty(session("member")) = True then
	'Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
	Response.write "<script>"
	Response.write "	alert('ท่านไม่ได้รับสิทธิ์ในการเข้าดูระบบนี้'); "
	Response.write " 	window.location.href=""default.asp""; "
	Response.write "</script>"
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
<title>ดูรายงาน</title>
<style>
.text {
					Font-size:14px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:14px}
.textsmall {
					Font-size:10px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:12px}
.textbig {
					Font-size:16px; Color:#000000;
					Font-family:MS Sans Serif ; line-height:16px}
.style1 {Font-size: 14px; Color: #000000; Font-family: MS Sans Serif; line-height: 14px; font-weight: bold; }
</style>
</head>

<body bgcolor="#ffffff">
<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center" bordercolor="#333333">
  <tr><td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" bordercolor="#333333">
  <tr>
    <td colspan="2" align="center" class="text">
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td align="center" class="textbig">รายงานการประชุมทบทวนโดยฝ่ายบริหาร</td>
        </tr>
        <tr>
          <td align="center" class="textbig">
		  <%
		  if RecView("D_Id") <> "01" and  RecView("D_Id") <> "02" then
		  response.write getDepartmentname(RecView("D_Id"))
		  else
		  	if RecView("D_Id") = "01" then
			response.write "คณะกรรมการบริหารระบบคุณภาพ"
			elseif RecView("D_Id") = "02" then
			response.write "คณะกรรมการประสานงานระบบคุณภาพ"
			end if
		  end if
		  %>
          </td>
        </tr>
        <tr>
          <td align="center" class="textbig">ครั้งที่ <%=RecView("MR_CountMeeting")%>&nbsp;&nbsp;&nbsp;วันที่ <%=DAy(RecView("MR_Date"))%>&nbsp;&nbsp;<%=thmonthFull(Month(RecView("MR_DAte")))%>&nbsp;&nbsp;<%=Year(RecView("MR_Date"))+543%></td>
        </tr>
      </table>      </td>
  </tr>
  
  <tr>
    <td height="50" colspan="2" align="left" class="text">สรุปเนื้อหาการประชุมตามประเด็นสำคัญดังนี้</td>
  </tr>
  <tr>
    <td colspan="2"  class="style1">1.&nbsp;ผลการตรวจติดตามคุณภาพ :</td>
  </tr>
  <% if RecView("MR_Review1") <> "" then %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <% if RecView("MR_Review1") <> "" then response.write RecView("MR_Review1") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="style1"><br />
      2. <span class="text">ข้อร้องเรียน ข้อคิดเห็น ของผู้รับบริการ</span> :</td>
  </tr>
  <% if RecView("MR_Review2") <> "" then  %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <% if RecView("MR_Review2") <> "" then response.write RecView("MR_Review2") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="style1"><br />
      3. <span class="text">ผลการดำเนินการตามเป้าหมายและตัวชี้วัด</span> :</td>
  </tr>
  <% if RecView("MR_Review3") <> "" then %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <% if RecView("MR_Review3") <> "" then response.write RecView("MR_Review3") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="style1"><br />
      4. <span class="text">สถานะของการปฏิบัติการแก้ไขและป้องกัน</span> :</td>
  </tr>
  <% if RecView("MR_Review4") <> "" then %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <% if RecView("MR_Review4") <> "" then response.write RecView("MR_Review4") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="text"><p><strong><br />
      5. การติดตามผลจากการประชุมที่ผ่านมา :</strong></p>    </td>
  </tr>
  <% if RecView("MR_Review5") <> "" then  %>
  <tr>
    <td colspan="2" valign="top"  class="text">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <% if RecView("MR_Review5") <> "" then response.write RecView("MR_Review5") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="style1"><br />
      6. <span class="text">การเปลี่ยนแปลงที่อาจมีผลกระทบต่อระบบ</span> :</td>
  </tr>
  <% if RecView("MR_Review6") <> "" then  %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <% if RecView("MR_Review6") <> "" then response.write RecView("MR_Review6") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <tr>
    <td colspan="2"  class="style1"><br />
      7. <span class="text">ข้อเสนอแนะเพื่อการปรับปรุง</span> :</td>
  </tr>
  <% if RecView("MR_Review7") <> "" then %>
  <tr>
    <td colspan="2" valign="top"  class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <% if RecView("MR_Review7") <> "" then response.write RecView("MR_Review7") else response.write "&nbsp;" end if %></td>
  </tr>
  <% end if %>
  <% if getDID = "01" or getDID = "02" then %>
  <tr>
    <td align="center" class="text" valign="top"><br />
      พิมพ์ธิดา วงศ์สุนทร<br />
      ผู้บันทึกรายงานการประชุม</td>
    <td align="center" class="text"><p><br />
      นายเจษฎาพร เจียรตระกูล<br />
      ผู้จัดการระบบคุณภาพ กองแผนงานและวิชาการ<br />
      ผู้ตรวจรายงานการประชุม</p>
      <% if getDID = "01" then %>
      <p>นายชาพล รัตนพันธุ์<br />
      ผู้อำนวยการกองแผนงานและวิชาการ<br />
      ผู้ตรวจรายงานการประชุม</p>
      <% end if %>
      </td>
  </tr>
  <% else %>
  <tr>
    <td width="40%" align="center" class="text"><p>&nbsp;</p>
      <p><%=RecView("MR_Record")%><br />
        ผู้บันทึกรายงานการประชุม</p></td>
    <td width="60%" align="center" class="text"><p>&nbsp;</p>
    <% if RecView("Flag_Check") = True then %>
      <p><%=GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'")%><br />
      ผู้จัดการระบบคุณภาพ<span class="text">&nbsp;<%=getDepartmentname(RecView("D_Id"))%><br />
      </span>ผู้ตรวจรายงานการประชุม</p>
    <% else %>
    &nbsp;
	<% end if%>
    </td>
  </tr>
  <% end if %>
</table>
</td></tr></table>
<br />
<div align="center"><label>
      <input type="button" name="butSave" id="butSave" value="พิมพ์เอกสาร"  onclick="javascript:{ window.print();}" />
      &nbsp;&nbsp;
      <% if getSource = "report" then %>
      <input type="button" name="butCancel" id="butCancel" value="กลับ" onClick="javascript:{ window.location.href='ManagementReviewReport.asp';}" />
      <% else %>
      <input type="button" name="butCancel" id="butCancel" value="กลับ" onClick="javascript:{ window.location.href='ManagementReview.asp?ID=<%=getDID%>';}" />
      <% end if %>
    </label></div>
<%
end if 
%>
</body>
</html>
