<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
dim Dateddmmyyyy
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
Datemmddyyyy1=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
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
		getADT = Request.QueryString("ADT")
		getNCP = Request.QueryString("NCP")
		getMC = Request.QueryString("MC")
		'response.write getR_Id
		SQL = "select * from Tb_Internalaudit where Audit_DocType='"&getADT&"' and  No_Car_Par='"&getNCP&"' and  M_Code='"&getMC&"'  "
		'response.write SQL
		set RecC = Server.CreateObject("ADODB.RECORDSET")
		RecC.open SQL,ConQS,1,3
		 while not RecC.EOF
		 
		 get_ADT=RecC("Audit_DocType")
		 get_Audit_Level=RecC("Audit_Level")
		 get_No_Car_Par=RecC("No_Car_Par")
		 get_Audit_Date=RecC("Audit_Date")
		 get_Audit_Source=RecC("Audit_Source")
		 
		 
		 get_Audit_Depart=RecC("Audit_Depart")
		 get_Audit_SubDepart=RecC("Audit_SubDepart")
		 get_Audit_SubDepartElseName=RecC("Audit_SubDepartElseName")
		 get_M_Code=RecC("M_Code")
		 get_M_Name=RecC("M_Name")
		 get_Audit_Name1=RecC("Audit_Name1")
		 get_Audit_Name2=RecC("Audit_Name2")
		 get_Audit_Name3=RecC("Audit_Name3")
		 get_Audit_Name4=RecC("Audit_Name4")
		 get_Audit_Name5=RecC("Audit_Name5")
		 get_Audit_Name6=RecC("Audit_Name6")
		 
		 get_Audit_Descript=RecC("Audit_Descript")
		 get_Audit_Advantages=RecC("Audit_Advantages")
		 get_Audit_Disadvantages=RecC("Audit_Disadvantages")
		 get_Audit_Flag_Complete = RecC("Audit_Flag_Complete")
		 
		 RecC.MoveNext()
		 wend
			 dim selectDepart,selectSubDepart
			 selectDepart=""
			selectSubDepart=""
			if get_Audit_Level = "1" then
				selectDepart = "checked=""checked"""
			else
				selectSubDepart = "checked=""checked"""
			end if
			
		 RowCAR = GetCountRowQS("Tb_Internalaudit","ID","where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='NC' and Audit_Year='"&(year(Dateddmmyyyy)+543)&"' and Audit_Level='"&get_Audit_Level&"' ")
		 RowPAR = GetCountRowQS("Tb_Internalaudit","ID","where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='OBS' and Audit_Year='"&(year(Dateddmmyyyy)+543)&"' and Audit_Level='"&get_Audit_Level&"' ")	
			
		 dim chkSelectNotFind
		 chkSelectNotFind=""
		 if get_ADT = "C" and get_Audit_Flag_Complete = "0" then
		 chkSelectNotFind = "checked=""checked"""
		 RowCAR=0
		 RowPAR=0
		 else
		 	
		 end if
		 
		 
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
<table width="100%" border="2" align="center" cellpadding="1" cellspacing="0" bordercolor="#000000">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td width="10%"><img src="images/aoryor.jpg" width="50" height="50" /></td>
        <td align="center" width="70%"><font style="font-size:18px"><strong>รายงานการตรวจติดตามคุณภาพภายใน<br />
          (Audit Report)</strong></font></td>
      <td width="20%" class="style1" align="right"><table width="80%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td class="style2"><label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart1" value="1"  onClick="ChangeJobresultGroup('','')"  <%=selectDepart%> disabled="disabled" />
    ระดับกรม</label></td>
        </tr>
        <tr>
          <td class="style2"><label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart1" value="2"  onClick="ChangeJobresultGroup('','')"  <%=selectSubDepart%> disabled="disabled" />
    ระดับหน่วยงาน</label></td>
        </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr><td>
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2"><b>Report No :</b></td>
            <td width="70%" class="style2"><%=get_No_Car_Par%></td>
          </tr>
        </table>
      </td></tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2"><b>วันที่ตรวจ :</b></td>
            <td width="70%" class="style2"><%=get_Audit_Date%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2" ><b>หน่วยงานที่รับการตรวจ :</b></td>
			<td width="70%" class="style2" ><%=getDepartmentname(get_Audit_Depart)%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2"><b>กิจกรรมที่ตรวจ (รหัสเอกสาร/ชื่อ)</b></td>
            <td width="70%" class="style2"><%=get_M_Code&" "&get_M_Name%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td width="30%" class="style2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="style2"><b>ผู้ตรวจติดตาม :</b></td>
              </tr>
            </table></td>
            <td width="70%"  class="style1" valign="middle"><table width="85%" border="0" cellspacing="0" cellpadding="5">
              <% if get_Audit_Name1 <> "" then %>
              <tr>
                <td width="5%" class="style2">1</td>
                <td width="65%" class="style2"><%=get_Audit_Name1%></td>
                <td width="30%" class="style2">หัวหน้าผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name2 <> "" then %>
              <tr>
                <td class="style2">2</td>
                <td class="style2"><%=get_Audit_Name2%></td>
                <td class="style2">ผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name3 <> "" then %>
              <tr>
                <td class="style2">3</td>
                <td class="style2"><%=get_Audit_Name3%></td>
                <td class="style2">ผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name4 <> "" then %>
              <tr>
                <td class="style2">4</td>
                <td class="style2"><%=get_Audit_Name4%></td>
                <td class="style2">ผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name5 <> "" then %>
              <tr>
                <td class="style2">5</td>
                <td class="style2"><%=get_Audit_Name5%></td>
                <td class="style2">ผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name6 <> "" then %>
              <tr>
                <td class="style2">6</td>
                <td class="style2"><%=get_Audit_Name6%></td>
                <td class="style2">ผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style1" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="2" >
              <tr>
                <td class="style2"><b>ผลการตรวจติดตาม</b></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
            <td width="70%" class="style1"><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="style2">
                <label>
                  <input type="checkbox" name="chkCAR" id="chkCAR" disabled="disabled"  <% if RowCAR <> 0 then response.write "checked=""checked""" end if%> />
                  พบหลักฐานที่แสดงว่าเกิดข้อบกพร่องหรือความไม่สอดคล้องขึ้นในระบบคุณภาพ<br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงออกใบ CAR จำนวน  <%=RowCAR%>  ใบ <% 'if RowCAR > 0 then %>ดังนี้<% 'end if %><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <%
				  		set recShow = Server.CreateObject("ADODB.RECORDSET")
						sqlNoCARPAR = "select  No_Car_Par from Tb_InternalAudit where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='NC' and Audit_Year='"&(year(Dateddmmyyyy)+543)&"' and Audit_Level='"&get_Audit_Level&"' "
						recShow.open sqlNoCARPAR,ConQS,1,3
						countLP = 1
						getLP = recShow.recordcount
						 while not recShow.EOF
						 if countLP < getLP then
						 response.write recShow("No_Car_Par")&"&nbsp;&nbsp;,&nbsp;&nbsp;"
						 else
						  response.write recShow("No_Car_Par")
						 end if
						 countLP = countLP+1
						 recShow.MoveNext
						 wend
						 recShow.close
				  %>
                    </label></td>
              </tr>
              <tr>
                <td class="style2"><label>
                  <input type="checkbox" name="chkPAR" id="chkPAR" disabled="disabled" <% if RowPAR <> 0 then response.write "checked=""checked""" end if%> />
                  พบความมีแนวโน้มที่จะเกิดข้อบกพร่องหรือความไม่สอดคล้องขึ้นในระบบคุณภาพ<br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงออกใบ PAR จำนวน  <%=RowPAR%>  ใบ  <% 'if RowPAR > 0 then %>ดังนี้<%' end if %><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <%
                  		set recShow = Server.CreateObject("ADODB.RECORDSET")
						sqlNoCARPAR = "select  No_Car_Par from Tb_InternalAudit where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='OBS' and Audit_Year='"&(year(Dateddmmyyyy)+543)&"'  and Audit_Level='"&get_Audit_Level&"' "
						recShow.open sqlNoCARPAR,ConQS,1,3
						countLP = 1
						getLP = recShow.recordcount
						 while not recShow.EOF
						 if countLP < getLP then
						 response.write recShow("No_Car_Par")&"&nbsp;&nbsp;,&nbsp;&nbsp;"
						 else
						  response.write recShow("No_Car_Par")
						 end if
						 countLP = countLP+1
						 recShow.MoveNext
						 wend
						 recShow.close
                    %>
                    </label></td>
              </tr>
              <tr>
                <td class="style2"><label>
      <input type="checkbox" name="checkComplete" id="checkComplete" onClick="chkAllowCARPAR()" <%=chkSelectNotFind%>  disabled="disabled" />
      ไม่พบข้อบกพร่อง</label></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr><td height="15">&nbsp;</td></tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="100%" class="style2"><b>รายละเอียดเพิ่มเติม/ข้อคิดเห็นของผู้ตรวจติดตาม :</b></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="100%" height="39" class="style2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=get_Audit_Descript%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="6">
          <tr>
            <td width="100%" height="39" class="style2"><b>ข้อดี :</b>&nbsp;&nbsp;<%=get_Audit_Advantages%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="6">
          <tr>
            <td width="100%" height="42" class="style2"><b>ข้อเสีย :</b>&nbsp;&nbsp;<%=get_Audit_Disadvantages%></td>
          </tr>
        </table></td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td><table width="97%" border="0" align="center" cellpadding="5" cellspacing="0">
          <% if get_Audit_Name1 <> "" or get_Audit_Name2 <> "" then %>
          <tr>
            <td width="50%">
            <% if get_Audit_Name1 <> ""  then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="65%">&nbsp;</td>
                <td width="25%" class="style1">หัวหน้าผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name1%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <% end if %>
            </td>
            <td width="50%">
			<% if get_Audit_Name2 <> ""  then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="65%">&nbsp;</td>
                <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name2%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table>
			<% end if %>
            </td>
          </tr>
          <% end if %>
          <% if get_Audit_Name3 <> "" or get_Audit_Name4 <> "" then %>
          <tr>
            <td>
            <% if get_Audit_Name3 <> ""  then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="65%">&nbsp;</td>
                <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name3%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table>
                
                </td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <% end if %>
            </td>
            <td>
            <% if get_Audit_Name4 <> ""  then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="65%">&nbsp;</td>
                <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name4%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <% end if %>
            </td>
          </tr>
          <% end if %>
          <% if get_Audit_Name5 <> ""  or get_Audit_Name6 <> "" then %>
          <tr>
            <td>
            <% if get_Audit_Name5 <> "" then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="70%">&nbsp;</td>
                <td width="20%" class="style1">ผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name5%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">&nbsp;</td>
                <td class="style1" align="center">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <% end if %>
            </td>
            <td>
            <% if get_Audit_Name6 <> "" then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="10%" class="style1">ลงชื่อ</td>
                <td width="70%">&nbsp;</td>
                <td width="20%" class="style1">ผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style2" align="center"><%=get_Audit_Name6%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">&nbsp;</td>
                <td class="style1" align="center">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <% end if %>
            </td>
          </tr>
          <% end if %>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
  <th align="right" class="style1">F-FDA-T-15 (0-30/09/56) หน้า 1/1</th>
</tr></table>
<br />
<br />
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/></div>
</body>
</html>
