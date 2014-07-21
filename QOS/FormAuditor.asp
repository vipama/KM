<%
	dim typePage
	dim DetailHead
	dim DetailHeadSub
	dim PageName
	dim Headpart2
	dim Textpart2_1
	dim Textpart2_2
	typePage = Request.QueryString("tp")
	if isempty(typePage) = true then
	else
			if typePage = 1 then
				DetailHead ="ใบขอให้ปฏิบัติการแก้ไข"
				DetailHeadSub="(Corrective Action Request : CAR)"
				DetailHeadNum="CAR NO. "
				PageName="CAR"
				Headpart2="ผู้ดำเนินการแก้ไข"
				Textpart2_1="แนวทางแก้ไข"
				Textpart2_2="แก้ไข"
			else
				DetailHead ="ใบขอให้ปฏิบัติการป้องกัน"
				DetailHeadSub="(Preventive Action Request : PAR)"
				DetailHeadNum="PAR NO. "
				PageName="PAR"
				Headpart2="ผู้ดำเนินการป้องกัน"
				Textpart2_1="แนวทางป้องกัน"
				Textpart2_2="ป้องกัน"
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
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
<script language="JavaScript" src="JScript/JS.js"></script>
</head>

<body>
<form  name="frmAuditor" id="frmAuditor"enctype="application/x-www-form-urlencoded" method="post">
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
          </tr></table>          
          </td>
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
        <td width="15%" colspan="2" class="style1">แนวโน้มข้อบกพร่องที่พบ : </td>
          <td width="85%" colspan="2" class="style1"><textarea name="AuditDepart"  wrap="hard" rows="3" cols="80"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea></td>
          </tr></table>        </td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="style1" width="50%">ลงชื่อผู้ออกใบ <%=PageName%> <input name="AuditLicense1" type="text"  id="AuditLicense1" size="60"  /></td>
            <td class="style1" width="50%">ลงชื่อ QMR <input name="AuditQMR1" type="text"  id="AuditQMR1" size="60"  /></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td class="style1">ส่วนที่ 2 : <%=Headpart2%></td>
      </tr>
      <tr>
        <td class="style1">
        <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
        <td width="10%" class="style1">สาเหตุของปัญหา</td>
        <td width="90%"></td>
        </tr>
        </table>        </td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="10%" class="style1">&nbsp;</td>
            <td width="90%"><span class="style1">
              <textarea name="AuditProblem"  wrap="hard" rows="3" cols="80"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea>
            </span></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="10%" class="style1"><%=Textpart2_1%></td>
            <td width="90%"></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="10%" class="style1">&nbsp;</td>
            <td width="90%"><span class="style1">
              <textarea name="AuditProtect"  wrap="hard" rows="3" cols="80"  style="overflow:auto;resize:none; vertical-align: bottom"  ></textarea>
            </span></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="2" cellspacing="0">
          <tr>
            <td width="50%" class="style1">กำหนดแล้วเสร็จภายในวันที่ :
              <label>วัน
              <select name="finishDay" id="finishDay">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="17">17</option>
                <option value="18">18</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
              เดือน
              <select name="finishMonth" id="finishMonth">
                <option value="1" selected="selected">มกราคม</option>
                <option value="2">กุมภาพันธ์</option>
                <option value="3">มีนาคม</option>
                <option value="4">เมษายน</option>
                <option value="5">พฤษภาคม</option>
                <option value="6">มิถุนายน</option>
                <option value="7">กรกฎาคม</option>
                <option value="8">สิงหาคม</option>
                <option value="9">กันยายน</option>
                <option value="10">ตุลาคม</option>
                <option value="11">พฤศจิกายน</option>
                <option value="12">ธันวาคม</option>
              </select>
              ปี
              <select name="finishYear" id="finishYear">
                <option value="2560">2560</option>
                <option value="2559">2559</option>
                <option value="2558">2558</option>
                <option value="2557" selected="selected">2557</option>
                                                        </select>
              </label></td>
            <td width="50%" class="style1">ลงชื่อผู้ดำเนินการ<%=Textpart2_2%> : 
            <input name="AuditEditname" type="text"  id="AuditEditname" size="40"  /></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td width="50%" class="style1">ลงชื่อหัวหน้าหน่วยงาน : <input name="AuditHeadDepart" type="text"  id="AuditHeadDepart" size="50"  /></td>
            <td width="50%"><span class="style1">ลงชื่อ QMR : 
              
                <input name="AuditQMR2" type="text"  id="AuditQMR2" size="60"  />
            </span></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><p>
      <input type="button"  value="print"  onclick="javascript:{ window.print();}"/>
    </p>
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td class="style1">ส่วนที่ 3 : ผู้ออกใบ <%=PageName%></td>
        </tr>
        <tr>
          <td>
          <table width="60%" cellpadding="3" cellspacing="0" border="0">
            <tr>
        <td width="30%" class="style1">ผลการตรวจติดตาม :</td>
          <td width="35%" class="style1"><label><input type="radio" name="AuditAccept"    value="0"  checked="checked" onclick="autoCheck('AuditAccept',this.value)" />
          ยอมรับ
          </label></td>
          <td width="35%" class="style1"><label><input name="OpenClose" type="radio"   value="0" checked="checked" onclick="autoCheck('OpenClose',this.value)" /> 
          ปิด </label><%=PageName%> No. : เลขปัจจุบันที่เปิดอยู่นี่</td>
          </tr></table>
          </td>
        </tr>
        <tr>
          <td><table width="70%" cellpadding="3" cellspacing="0" border="0">
            <tr>
              <td width="30%" class="style1">&nbsp;</td>
              <td width="35%" class="style1"><label>
                <input type="radio" name="AuditAccept"    value="1" onclick="autoCheck('AuditAccept',this.value)"  />
                ไม่ยอมรับ </label></td>
              <td width="35%" class="style1"><label>
                <input type="radio" name="OpenClose"   value="1"  onclick="autoCheck('OpenClose',this.value)" />
                เปิด </label>
                  <%=PageName%>No. : 
                  <input type="text" name="textfield" id="textfield" /></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>      <p>&nbsp; </p></td>
  </tr>
</table>
</form>
</body>
</html>
