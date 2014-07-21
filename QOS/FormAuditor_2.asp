<%
	dim typePage
	dim DetailHead
	dim DetailHeadSub
	dim PageName
	typePage = Request.QueryString("tp")
	if isempty(typePage) = true then
	else
			if typePage = 1 then
				DetailHead ="ใบขอให้ปฏิบัติการแก้ไข"
				DetailHeadSub="(Corrective Action Request : CAR)"
				DetailHeadNum="CAR NO. "
				PageName="CAR"
			else
				DetailHead ="ใบขอให้ปฏิบัติการป้องกัน"
				DetailHeadSub="(Preventive Action Request : PAR)"
				DetailHeadNum="PAR NO. "
				PageName="PAR"
			end if
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Form QS</title>
<style type="text/css">
<!--
.style1 {
font-size:12px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
</head>

<body>
<table width="85%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#000000" >
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%" align="center"><img src="images/aoryor.jpg" width="50" height="50" /></td>
         <td width="70%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td align="center"><%=DetailHead%></td>
           </tr>
           <tr>
             <td align="center"><%=DetailHeadSub%></td>
           </tr>
         </table></td>
         <td width="20%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
           <tr>
             <td><%=DetailHeadNum%></td>
           </tr>
           <tr>
             <td>&nbsp;</td>
           </tr>
         </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td  class="style1">ส่วนที่ 1 : ผู้ออกใบ <%=PageName%></td>
      </tr>
      <tr>
        <td>
        <table width="100%" cellpadding="3" cellspacing="0" border="0"><tr>
        <td width="10%" class="style1">ที่มา : </td>
          <td width="30%" class="style1"><label><input type="radio" name="radio"    value="1" checked="checked" /> การตรวจติดตามคุณภาพภายใน</label></td>
          <td width="30%" class="style1"><label><input type="radio" name="radio"   value="2" /> การตรวจประเมินจากภายนอก</label></td>
          <td width="30%" class="style1"><label><input type="radio" name="radio"   value="3" /> การประชุมทบทวนโดยฝ่ายบริหาร</label></td>
          </tr></table>          </td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="3" cellspacing="0" border="0"><tr>
        <td width="10%">&nbsp;</td>
          <td width="30%" valign="bottom" class="style1"><label><input type="radio" name="radio"    value="4" /> การปฏิบัติงาน</label></td>
          <td width="30%" valign="bottom" class="style1"><label><input type="radio" name="radio"   value="5" /> ข้อร้องเรียนจาก</label> <textarea name="SourceDetail1"  wrap="hard" rows="1" cols="20"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea></td>
          <td width="30%" valign="bottom" class="style1"><label><input type="radio" name="radio"   value="6" /> อื่นๆ</label> <textarea name="SourceDetail2"  wrap="hard" rows="1" cols="20" style="overflow:auto;resize:none; vertical-align:bottom" ></textarea></td>
          </tr></table></td>
      </tr>
      <tr>
        <td>
        <table width="100%" cellpadding="3" cellspacing="0" border="0"><tr>
        <td width="10%" class="style1">หน่วยงานที่พบ : </td>
          <td width="30%" class="style1"><textarea name="AuditDepart"  wrap="hard" rows="1" cols="40"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea></td>
          <td width="30%">&nbsp;</td>
          <td width="30%">&nbsp;</td>
          </tr></table>        </td>
      </tr>
      <tr>
        <td>
        <table width="100%" cellpadding="3" cellspacing="0" border="0"><tr>
        <td width="20%" colspan="2" class="style1">แนวโน้มข้อบกพร่องที่พบ : </td>
          <td width="80%" colspan="2" class="style1"><textarea name="AuditDepart"  wrap="hard" rows="3" cols="80"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea></td>
          </tr></table>
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><input type="button"  value="print"  onclick="javascript:{ window.print();}"/></td>
  </tr>
</table>
</body>
</html>
