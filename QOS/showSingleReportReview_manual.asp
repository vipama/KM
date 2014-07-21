<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if

'	if isEmpty(session("member")) = True then
'		Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
'	end if
		dim getR_Id,chkComport,chkUnComport,SOP_Q,SOP_P,SOP_W,showExpectedDate
		chkComport=""
		chkUnComport=""
		SOP_Q=""
		SOP_P=""
		SOP_W=""
		getR_Id = Request.Form("hidRID")
		getR_Id="2557-7-87"
		'response.write getR_Id
		SQL = "select * from Tb_Review where No_Review='"&getR_Id&"'"
		set RecReview = Server.CreateObject("ADODB.RECORDSET")
		RecReview.open SQL,ConQS,1,3
		 while not RecReview.EOF
		 
		 getR_Id=RecReview("R_Id")
		 getNo_Review=RecReview("No_Review")
		 getType_Sop=RecReview("Type_Sop")
		 getCurrent_Date=RecReview("CurrentReviewDate")
		 getD_Id=RecReview("D_Id")
		 getM_Code=RecReview("M_Code")
		 getM_Name=RecReview("M_Name")
		 getComport=RecReview("Comport")
		 getLogic_Comport1=RecReview("Logic_Comport1")
		 getLogic_Comport2=RecReview("Logic_Comport2")
		 getLogic_Comport3=RecReview("Logic_Comport3")
		 getLogic_Comport4=RecReview("Logic_Comport4")
		 getLogic_Comport5=RecReview("Logic_Comport5")
		 getUncomport=RecReview("Uncomport")
		 getMethodType=RecReview("MethodType")
		 getRemake_Date=RecReview("Remake_Date")
		 getEdit_Date=RecReview("Edit_Date")
		 getLogic_Uncomport1=RecReview("Logic_Uncomport1")
		 getLogic_Uncomport2=RecReview("Logic_Uncomport2")
		 getLogic_Uncomport3=RecReview("Logic_Uncomport3")
		 getLogic_Uncomport4=RecReview("Logic_Uncomport4")
		 getLogic_Uncomport5=RecReview("Logic_Uncomport5")
		 getName_Review=RecReview("Name_Review")
		 getLevel_Review=RecReview("Level_Review")
		 RecReview.MoveNext()
		 wend
		 
		 if getComport = True  then
			chkComport ="checked=""checked"""
			chkunComport=""
			
		else
			chkComport =""
			chkunComport ="checked=""checked"""
		end if
		dim chkRe,chkEd,chkCa
		if getMethodType = 1 then
			chkRe = "checked=""checked"""
			chkEd = ""
			chkCa = ""
		elseif getMethodType= 2 then
			chkRe = ""
			chkEd = "checked=""checked"""
			chkCa = ""
		elseif getMethodType=3 then
			chkRe = ""
			chkEd = ""
			chkCa = "checked=""checked"""
		end if
		if getcomport = True then
			chkComport = "checked=""checked"""
		end if
		if getUncomport = True then
			chkUncomport = "checked=""checked"""
		end if
		'------------------------------check for type of SOP-----------------------------
		if getType_Sop = "Q" then
			SOP_Q="checked=""checked"""
		elseif getType_Sop = "PC" or getType_Sop = "PS" then
			SOP_P="checked=""checked"""
		elseif getType_Sop = "W" then
			SOP_W="checked=""checked"""
		end if
		'------------------------------------------------------------------------------------
'end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>รายงานผลวิเคราะห์</title>
<style type="text/css">
<!--
.style1 {
font-size:10px;
font-family:Arial, Helvetica, sans-serif;
}
.style2 {
font-size:12px;
font-family:Arial, Helvetica, sans-serif;
}
-->
</style>
</head>

<body>
<table width="97%" border="2" cellpadding="1" cellspacing="0" bordercolor="#000000">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td width="10%"><img src="images/aoryor.jpg" width="50" height="50" /></td>
        <td align="center" width="70%"><font style="font-size:18px"><strong>รายงานการทบทวนเอกสาร<br />
          (Documentation Review Report)</strong></font></td>
      <td width="20%" class="style1" align="center"><b>No.</b> <%=getNo_Review%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>
    <!--Start of Part 1-->
    &nbsp;
    <!--End of Part 1-->
    <table width="100%" border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="25%" class="style2"><b>ส่วนที่ 1 : ผู้ทบทวน</b></td>
            <td width="55%">&nbsp;</td>
            <td width="20%" class="style1"><b>วันที่</b> <%=getCurrent_Date%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="8%" class="style1" ><b>ข้าพเจ้า :</b></td>
			<td width="42%" class="style1" ><%=getName_Review%></td>
            <td width="10%" class="style1" align="right">&nbsp;&nbsp;&nbsp;<b>หน่วยงาน :</b></td>
			<td width="40%" class="style1" ><%=getDepartmentname(getD_Id)%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="80%" border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td class="style1"><b>ดำเนินการทบทวน :</b></td>
            <td class="style1"><input type="radio" name="radioTypeSOP" id="radioTypeSOP_Q" value="Q"  <%=SOP_Q%> disabled="disabled"   />
              <label for="radioTypeSOP">คู่มือคุณภาพ (Q)</label></td>
            <td class="style1"><input type="radio" name="radioTypeSOP" id="radioTypeSOP_P" value="P"  <%=SOP_P%> disabled="disabled"/>
              <label for="radioTypeSOP">คู่มือขั้นตอนการปฏิบัติงาน (P)</label></td>
            <td class="style1"><input type="radio" name="radioTypeSOP" id="radioTypeSOP_W" value="W"  <%=SOP_W%>  disabled="disabled"/>
              <label for="radioTypeSOP">คู่มือวิธีการปฏิบัติงาน (W)</label></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td width="15%" class="style1"><b>รหัสเอกสาร :</b></td>
            <td width="10%"  class="style1"><%=getM_Code%></td>
            <td width="12%" class="style1" align="right" >&nbsp;&nbsp;&nbsp;<b>ชื่อเอกสาร :</b></td>
            <td width="63%"  class="style1"><%=getM_Name%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="15%" class="style1"><b>ผลการทบทวน :</b></td>
            <td width="85%" class="style1"><label>
              <input type="radio" name="radioPerfect" id="radioPerfect1" value="radioPerfect"  <%=chkcomport%> disabled="disabled"/>
              มีความเหมาะสมไม่ต้องดำเนินการใดๆ</label></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="12%">&nbsp;</td>
            <td width="88%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="7%" class="style1"><b>เหตุผล :</b></td>
                <td width="93%" class="style1"><%=getLogic_Comport1&" "&getLogic_Comport2&" "&getLogic_Comport3%>&nbsp;<% if getLogic_comport4 = True then response.write getLogic_comport5 end if%></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%">&nbsp;</td>
            <td width="90%" class="style1"><label>
              <input type="radio" name="radioPerfect" id="radioPerfect2" value="radioPerfect" <%=chkUncomport%> disabled="disabled" />
              ไม่มีความเหมาะสม ต้องดำเนินการ</label></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="17%">&nbsp;</td>
            <td width="88%"><table width="70%" border="0" cellspacing="0" cellpadding="1">
              <tr>
                <td width="34%" class="style1"><label>
                  <input type="radio" name="radio" id="radioRemake_R" value="Remake" <%=chkRe%> disabled="disabled" />
                  จัดทำใหม่</label></td>
                <td width="33%" class="style1"><label>
                  <input type="radio" name="radio" id="radioRemake_E" value="Edit" <%=chkEd%> disabled="disabled" />
                  แก้ไข</label></td>
                <td width="33%" class="style1"><label>
                  <input type="radio" name="radio" id="radioRemake_C" value="Cancel"  <%=chkCa%> disabled="disabled" />
                  ยกเลิก</label></td>
              </tr>
              <tr>
              <td class="style1" colspan="3">
              <table width="100%" cellpadding="3" cellspacing="0" border="0">
              <tr><td width="30%" class="style1">คาดว่าจะแล้วเสร็จ : </td>
              <td width="80%" class="style1">
              <%
			  if getMethodType = 1 then
			  		response.Write getRemake_Date
			  elseif  getMethodType = 2 then
			  		response.write getEdit_Date
			  end if
			  %>              </td>
              </tr>
              </table>              </td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="12%">&nbsp;</td>
            <td width="88%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="7%" class="style1"><b>เหตุผล :</b></td>
                <td width="93%" class="style1"><%=getLogic_Uncomport1&" "&getLogic_Uncomport2&" "&getLogic_Uncomport3%>&nbsp;<% if getLogic_Uncomport4 = True then response.write getLogic_Uncomport5 end if %></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr><td colspan="3">&nbsp;</td></tr>
          <tr>
            <td width="20%">&nbsp;</td>
            <td width="20%">&nbsp;</td>
            <td width="60%"><table width="90%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="70%">&nbsp;</td>
                <td width="20%" class="style1">ผู้ทบทวน</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center"><%=getName_Review%></td>
                    <td width="1%" class="style1">)</td>
                  </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">ตำแหน่ง</td>
                <td class="style1" align="center" >&nbsp;&nbsp;&nbsp;<%=getLevel_Review%></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=getCurrent_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">

    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="style2"><b>ส่วนที่ 2 : ผู้ตรวจสอบ</b></td>
          <td width="50%">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="15%" class="style1" align="center"><b>ผลการพิจารณา</b></td>
          <td width="43%" class="style1"><label>
            <input type="checkbox" name="chkAgree" id="chkAgree" />
            เห็นชอบ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="checkbox" name="chkNotAgree2" id="chkNotAgree" />
ไม่เห็นชอบ</label></td>
          <td width="42%">&nbsp;</td>
        </tr>
      </table></td>
    </tr>

    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5%">&nbsp;</td>
          <td width="60%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="10%" class="style1"><b>เหตุผล :</b></td>
              <td width="90%">&nbsp;</td>
            </tr>
          </table></td>
          <td width="42%">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="20%">&nbsp;</td>
          <td width="20%">&nbsp;</td>
          <td width="60%"><table width="90%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="70%">&nbsp;</td>
                <td width="20%" class="style1">ผู้ตรวจสอบ</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%">&nbsp;</td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">ตำแหน่ง</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
  </table></td>
  </tr>
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%" class="style2"><b>ส่วนที่ 3 : ผู้อนุมัติ</b></td>
            <td width="50%">&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" class="style1" align="center"><b>ผลการพิจารณา</b></td>
            <td width="43%" class="style1"><label>
              <input type="checkbox" name="chkAgree2" id="chkAgree2" />
              อนุมัติ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="checkbox" name="chkNotAgree" id="chkNotAgree2" />
ไม่อนุมัติ</label></td>
            <td width="42%">&nbsp;</td>
          </tr>
      </table></td>
    </tr>

    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5%">&nbsp;</td>
            <td width="60%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="10%" class="style1"><b>เหตุผล :</b></td>
                  <td width="90%">&nbsp;</td>
                </tr>
            </table></td>
            <td width="42%">&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%">&nbsp;</td>
            <td width="20%">&nbsp;</td>
            <td width="60%"><table width="90%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="10%" class="style1">ลงชื่อ</td>
                  <td width="70%">&nbsp;</td>
                  <td width="20%" class="style1">ผู้อนุมัติ</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td width="1%" class="style1">(</td>
                        <td width="98%">&nbsp;</td>
                        <td width="1%" class="style1">)</td>
                      </tr>
                  </table></td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td class="style1">ตำแหน่ง</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td class="style1">วันที่</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table></td>
  </tr>
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%" class="style2"><b>ส่วนที่ 4 : ผู้ควบคุมเอกสารกลาง</b></td>
            <td width="50%">&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%" class="style1" align="center">รับทราบผลการทบทวนเอกสาร</td>
            <td width="40%" class="style1">&nbsp;</td>
            <td width="50%">&nbsp;</td>
          </tr>
      </table></td>
    </tr>

    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%">&nbsp;</td>
            <td width="20%">&nbsp;</td>
            <td width="60%"><table width="95%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="10%" class="style1">ลงชื่อ</td>
                  <td width="65%">&nbsp;</td>
                  <td width="25%" class="style1">ผู้ควบคุมเอกสารกลาง</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td width="1%" class="style1">(</td>
                        <td width="98%">&nbsp;</td>
                        <td width="1%" class="style1">)</td>
                      </tr>
                  </table></td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td class="style1">ตำแหน่ง</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td class="style1">วันที่</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table></td>
  </tr>
</table>
<table width="97%" cellpadding="0" cellspacing="0" border="0"><tr>
  <th align="right" class="style1">F-FDA-T-7 (0-30/09/56) หน้า 1/1</th>
</tr></table>
<br />
<br />
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/>&nbsp;&nbsp;<input type="button"  value="กลับหน้ากรอกข้อมูล"  onclick="javascript:{ window.location.href='FReview.asp';}"/></div>
</body>
</html>
