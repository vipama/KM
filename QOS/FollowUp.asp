<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
getCARPAR = Request.Form("txtCPId")
getFlagSave = Request.Form("hidS")
gethidCPId = Request.Form("hidCPId")
getCPId = Request.Form("txtCPId")
getchkPrint = Request.Form("hidPrint")
'	if isEmpty(session("member")) = True then
'		Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
'	end if
		dim Dateddmmyyyy
		Dateddmmyyyy=Now()
		Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
		Datemmddyyyy1=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
		Dateddmmyyyy=day(Dateddmmyyyy)&"/"&month(Dateddmmyyyy)&"/"&(year(Dateddmmyyyy)+543)
		dim getDateUpdate
		
		
		if isEmpty(gethidCPId) = False then
			
			getchkAccept = Request.Form("chkAccept")
			getchkOpenClose = Request.Form("chkOpenClose")
			gettxtReason = Request.Form("txtReason")
			gethidQMRpart3 = Request.Form("hidQMRpart3")
		
			SQL_Update = "Update Tb_InternalAudit set  Audit_Accept='"&getchkAccept&"' ,  Audit_OpenClose='"&getchkOpenClose&"' ,  Audit_Reason='"&gettxtReason&"' , Audit_License_P3='"&GetSingleField("member","NAME"," where email='"&Session("Member")&"' ")&" "&GetSingleField("member","SURNAME"," where email='"&Session("Member")&"' ")&"' , Audit_QMR_P3='"&gethidQMRpart3&"' , Audit_Date2='"&Datemmddyyyy&"'  where No_Car_Par='"&gethidCPId&"'   "
			'response.write SQL_Update&"<br>"
			ConQS.execute(SQL_Update)
			sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Update Part3','"&Datemmddyyyy1&"','"&getDepartmentname(GetSingleFieldQS("Tb_InternalAudit","Audit_Depart"," where  No_Car_Par='"&gethidCPId&"'"))&"','"&GetSingleFieldQS("Tb_InternalAudit","M_Code","where  No_Car_Par='"&gethidCPId&"'")&"','"&gethidCPId&"','NC')"
			'response.write sql_log&"<br />"
			ConQS.execute(sql_log)
				If Err.Number = 0 Then
					response.write "<script language=""javascript"">"
					response.write "alert(""Save Data Success"");"
					'response.write "window.location.href=""CalendarActivity.asp?date="&Day(getADate)&"/"&Month(getADate)&"/"&(year(getADate)+543)&" "" "
					'response.write "window.location.href=""EditCAR.asp"" "
					response.write "</script>"
				end if
		end if
		
		dim getR_Id,chkComport,chkUnComport,SOP_Q,SOP_P,SOP_W,showExpectedDate,PageName
		chkComport=""
		chkUnComport=""
		SOP_Q=""
		SOP_P=""
		SOP_W=""
		PageName = ""
		
		'response.write getR_Id
		if isEmpty(gethidCPId) = True then
			SQL = "select * from Tb_Internalaudit where  No_Car_Par='"&getCPId&"'   "
		else
			SQL = "select * from Tb_Internalaudit where  No_Car_Par='"&gethidCPId&"'  "
		end if
		'response.write SQL
		set RecC = Server.CreateObject("ADODB.RECORDSET")
		RecC.open SQL,ConQS,1,3
		if RecC.RecordCount <= 0 and isEmpty(getCARPAR) = False    then
					response.write "<script language=""javascript"">"
					response.write "alert(""No data in database"");"
					'response.write "window.location.href=""CalendarActivity.asp?date="&Day(getADate)&"/"&Month(getADate)&"/"&(year(getADate)+543)&" "" "
					response.write "window.location.href=""EditCAR.asp"" "
					response.write "</script>"
		end if
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
		 get_Audit_QMR_P1 = RecC("Audit_QMR_P1")
		 
		 get_Audit_Descript=RecC("Audit_Descript")
		 get_Audit_Advantages=RecC("Audit_Advantages")
		 get_Audit_Disadvantages=RecC("Audit_Disadvantages")
		 get_Audit_Defect = RecC("Audit_Defect")
		 get_Audit_License_P1 = RecC("Audit_License_P1")
		 
		 get_Audit_Problem = RecC("Audit_Problem")
		 get_Audit_Protect = RecC("Audit_Protect")
		 get_Audit_Finish_Date = RecC("Audit_Finish_Date")
		 get_Audit_Edit_Name = RecC("Audit_Edit_Name")
		 get_Audit_Head_Depart = RecC("Audit_Head_Depart")
		 get_Audit_QMR_P2 = RecC("Audit_QMR_P2")
		 get_Audit_Date2 = RecC("Audit_Date2")
		 
		 get_Audit_Accept = RecC("Audit_Accept")
		 get_Audit_OpenClose = RecC("Audit_OpenClose")
		 get_Audit_Reason = RecC("Audit_Reason")
		 get_Audit_License_P3 = RecC("Audit_License_P3")
		 get_Audit_QMR_P3 = RecC("Audit_QMR_P3")
		 get_Audit_Date3 = RecC("Audit_Date3")
		 get_Audit_Year = RecC("Audit_Year")
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
		 dim chkSelectNotFind
		 chkSelectNotFind=""
		 if get_ADT = "C" then
		 	chkSelectNotFind = "checked=""checked"""
		 else
		 	
		 end if
		 '-------------------------------------block set page name---------------------------------
		 if get_ADT = "NC" then
		 	PageName="CAR"
		 elseif get_ADT = "OBS" then
		 	PageName="PAR"
		 end if
		 '----------------------------------------------------------------------------------------------
		 
		 RowCAR = GetCountRowQS("Tb_Internalaudit","ID","where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='NC' ")
		 RowPAR = GetCountRowQS("Tb_Internalaudit","ID","where M_code='"&get_M_Code&"' and Audit_Depart='"&get_Audit_Depart&"' and Audit_Doctype='OBS' ")
		 
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
font-size:14px;
font-family:THSarabunPSK,Arial, Helvetica, sans-serif;
}
-->
</style>
<script language="javascript" src="jScript/JS.js"></script>
</head>

<body>
<% 

if isEmpty(getCARPAR) = True  and  isEmpty(getFlagSave) = True and isEmpty(gethidCPId) = True then
%>
<form name="frmEdit" enctype="application/x-www-form-urlencoded" method="post" action="FollowUp.asp">
<input type="hidden"  name="hidprint" id="hidprint" value=""/>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="80%">
<tr>
<td width="25%" align="right"><b>เลขที่ใบ CAR / PAR :</b></td>
<td width="25%">&nbsp;&nbsp;&nbsp;<input name="txtCPId" type="text" id="txtCPId" size="30" /></td>
<td width="50%">&nbsp;&nbsp;<input  type="submit" value="ค้นหาเพื่อแก้ไขเอกสาร" name="butSearch" id="butSearch"  />
&nbsp;&nbsp;<input  type="button" value="ค้นหาเพื่อพิมพ์เอกสาร" name="butSearchPrint" id="butSearchPrint" onClick="goFollowupPrint()"  />
&nbsp;&nbsp;&nbsp;<input type="button" value="ยกเลิก"  id="butCancel" name="butCancel" onClick="javascript:{ window.close(); }"/></td>
</tr>
</table>
</form>
<% else %>
<br />
<table width="100%" border="2" align="center" cellpadding="1" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td width="10%"><img src="images/aoryor.jpg" width="50" height="50" /></td>
        <td align="center" width="70%"><font style="font-size:18px"><strong>ใบขอให้ปฏิบัติการแก้ไข<br />
          <% if PageName = "CAR" then %>
          (Corrective Action Request : CAR)
          <% else %>
           (Preventive Action Request : PAR)
		  <% end if %>
          </strong></font></td>
      <td width="20%" class="style1" align="right"><table width="90%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td><label>
            <input type="checkbox" name="chkbox_department" id="chkbox_department" <% if get_Audit_Level = "1" then %>checked="checked" <% end if%> disabled="disabled" />
            ระดับกรม</label></td>
        </tr>
        <tr>
          <td><label>
            <input type="checkbox" name="chkbox_agencies" id="chkbox_agencies" <% if get_Audit_Level = "2" then %>checked="checked" <% end if%> disabled="disabled" />
            ระดับหน่วยงาน</label></td>
        </tr>
        <tr>
          <td class="style1"><b><%=PageName%> No.</b> <%=get_No_Car_Par%></td>
        </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr><td colspan="3" bgcolor="#CCCCCC" class="style1"><span style="font-weight: bold">ส่วนที่ 1 : ผลการติดตาม</span></td>
  </tr>
  <tr>
    <td valign="top" width="60%">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td class="style1"><b>ตรวจติดตามวันที่ : &nbsp;<%=get_Audit_Date%></b></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="20%" valign="top" class="style1"><b>ที่มา :</b></td>
            <td width="90%" class="style1" ><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkAuditIn" id="chkAuditIn"  <% if get_Audit_Source = "1" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  การตรวจติดตามคุณภาพภายใน</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkAuditOut" id="chkAuditOut" <% if get_Audit_Source = "2" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  การตรวจประเมินจากภายนอก</label></td>
                </tr>
              <tr>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkWork" id="chkWork" <% if get_Audit_Source = "4" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  การปฏิบัติงาน</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkRequest" id="chkRequest" <% if get_Audit_Source = "5" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  ข้อร้องเรียนจาก</label></td>
                </tr>
              <tr>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkMeetingReview" id="chkMeetingReview" <% if get_Audit_Source = "3" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  การประชุมทบทวนโดยฝ่ายบริหาร</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkElse" id="chkElse" <% if get_Audit_Source = "6" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  อื่นๆ</label></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="2" class="style1" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="20%"><b>หน่วยงานที่พบ :</b></td>
                  <td width="80%"><%=getDepartmentname(get_Audit_Depart)%>&nbsp;&nbsp;
                  <%
				if GetSingleFieldQS("Tb_Subdepart","Name_Subdepart","where Subdepart_ID="&get_Audit_SubDepart) <> "0" then
					response.write GetSingleFieldQS("Tb_Subdepart","Name_Subdepart","where Subdepart_ID="&get_Audit_SubDepart)
				else
					response.write get_Audit_SubDepartElseName
				end if
				 %>
                  </td>
                </tr>
              </table></td>
			</tr>
        </table></td>
      </tr>
      <tr><td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="25%" class="style1"><b>ขอบเขตการตรวจ :</b></td>
          <td width="65%" class="style1"><%=get_M_Name&" "&get_M_Code%></td>
        </tr>
      </table></td></tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="20%" class="style1"><b><% if PageName = "CAR" then %>ข้อบกพร่องที่พบ :<% else %>แนวโน้มข้อบกพร่องที่พบ :<% end if%></b></td>
            <td width="80%" class="style1"><%=get_Audit_Defect%></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td width="40%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="15%" class="style1">ลงชื่อ</td>
              <td width="60%" align="center"><span class="style1"><%=get_Audit_License_P1%></span></td>
              <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center"><b><%=get_Audit_License_P1%></b></td>
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
        </table></td>
      </tr>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="15%" class="style1">ลงชื่อ</td>
              <td width="60%">&nbsp;</td>
              <td width="25%" class="style1">ผู้รับการตรวจ</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center">&nbsp;</td>
                    <td width="1%" class="style1">)</td>
                  </tr>
              </table></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="style1">วันที่</td>
              <td class="style1" align="center">&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr>
                  <td colspan="3" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;รับทราบผลการตรวจติดตาม วันที่.......<%=get_Audit_Date%>.......</td>
                  </tr>
                <tr>
                  <td width="15%" class="style1">ลงชื่อ</td>
                  <td width="60%" align="center"><span class="style1"><%=get_Audit_QMR_P1%></span></td>
                  <td width="25%" class="style1">QMR</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td width="1%" class="style1">(</td>
                        <td width="98%"class="style1" align="center"><b><%=get_Audit_QMR_P1%></b></td>
                        <td width="1%" class="style1">)</td>
                      </tr>
                  </table></td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="3" class="style1"><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
                  </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr><td colspan="2" bgcolor="#CCCCCC" class="style1"><span style="font-weight: bold">ส่วนที่ 2 : การปฏิบัติการ<% if PageName = "CAR" then %>แก้ไข<% else %>ป้องกัน<% end if %></span> </td></tr>
  <tr>
  <td width="60%" valign="top">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20%" class="style1"><b>สาเหตุของปัญหา :</b></td>
          <td width="80%" class="style1">
          <%=get_Audit_Problem%>          </td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20%" class="style1"><b>
            <% if PageName = "CAR" then%>แนวทางแก้ไข<% else %>แนวทางป้องกัน<% end if %> :</b></td>
          <td width="80%" class="style1">
          <%=get_Audit_Protect%>          </td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="25%" class="style1"><b>กำหนดวันแล้วเสร็จ : 
            &nbsp;</b></td>
          <td width="75%" class="style1"><b><%=get_Audit_Finish_Date%></b></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="15%" valign="top" class="style1">หมายเหตุ : </td>
          <td width="85%" class="style1">ผู้รับการตรวจต้องดำเนินการวิเคราะห์สาเหตุของปัญหา กำหนดแนวทาง<% if PageName = "CAR" then%>แก้ไข<% else %>ป้องกัน<% end if %>
            และกำหนดวันแล้วเสร็จ ส่งให้ผู้ตรวจติดตามพิจารณา ภายใน 30 วันนับจากวันที่ตรวจติดตามแล้วเสร็จ</td>
        </tr>
      </table></td>
    </tr>
  </table>  </td>
  <td width="40%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%" align="center" class="style1"><%=get_Audit_Edit_Name%> </td>
            <td width="25%" class="style1"><%if PageName = "CAR" then %>
              ผู้ดำเนินการแก้ไข
              <% else %>
              ผู้ดำเนินการป้องกัน
              <% end if%></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><b>
                    <%
						response.write get_Audit_Edit_Name
					%>
                  </b></td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center"><%
			   if isEmpty(gethidCPId) = True  and getchkPrint <> "Print"  then
			  		response.write Dateddmmyyyy
			  else
			  		response.write  get_Audit_Date2 
			  end if
			  %></td>
            <td>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%" align="center" class="style1"><%=get_Audit_Head_Depart%></td>
            <td width="25%" class="style1">หัวหน้าหน่วยงาน</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><b><%=get_Audit_Head_Depart%></b></td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="15%" class="style1">ลงชื่อ</td>
                <td width="60%" align="center" class="style1"><%=get_Audit_QMR_P1%></td>
                <td width="25%" class="style1">QMR</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><b><%=get_Audit_QMR_P1%></b></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center"><%=get_Audit_Date2%></td>
                <td>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="3" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;เห็นชอบตามแนวทางการปฏิบัติการ
                <% if PageName = "CAR" then %>
              แก้ไข
              <% else %>
              ป้องกัน
              <% end if %>
              ที่เสนอมา</td>
          </tr>
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center">&nbsp;</td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center"><%=get_Audit_Date2%></td>
            <td>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
  </table></td>
  </tr>
  <tr><td colspan="2" bgcolor="#CCCCCC" class="style1" style="font-weight: bold">ส่วนที่ 3 : ผลการตรวจติดตามซ้ำ</td></tr>
  <tr>
  <td width="60%" valign="top"> <form name="frmEditPart3" id="frmEditPart3" enctype="application/x-www-form-urlencoded" action="FollowUp.asp" method="post">
  <input type="hidden"  name="hidCPId" id="hidCPId" value="<%=get_No_Car_Par%>"/>
  <input type="hidden" name="hidQMRpart3" id="hidQMRpart3"  value="<%=get_Audit_QMR_P2%>" /><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td><span class="style1">ตรวจติดตามซ้ำวันที่ : <%=Dateddmmyyyy%></span></td>
          </tr>
        <tr>
          <td><label>
            <input type="checkbox" name="chkOpenClose" id="chkClose"  value="2" 
			<% 
			if isEmpty(gethidCPId) = False or getchkPrint = "Print" then  
				response.write "disabled=""disabled"" "  
			end if  
			if isEmpty(gethidCPId) = False or getchkPrint = "Print" then
				    if get_Audit_OpenClose = "2" then
					  response.write " checked=""checked"" "
					else
					  response.write ""
					end if
		   else
		   		response.write ""
		   end if
			
			 %>  />
            มีประสิทธิผล ปิดประเด็น 
            <% if PageName = "CAR" then %>CAR<% else %>PAR<% end if %>
          </label></td>
          </tr>
        <tr>
          <td>
            <label>
            <input type="checkbox" name="chkOpenClose" id="chkOpen" value="1" 
			<% 
			if isEmpty(gethidCPId) = False or getchkPrint = "Print" then
				response.write "disabled=""disabled"" "
			end if  
			if isEmpty(gethidCPId) = False or getchkPrint = "Print" then
				    if get_Audit_OpenClose = "1" then
					  response.write " checked=""checked"" "
					else
					  response.write ""
					end if
		   else
		   		response.write ""
		   end if
			%>  />
            ยังพบข้อบกพร่องอยู่ เปิด 
            <% if PageName = "CAR" then %>CAR<% else %>PAR<% end if %> 
            No. :......................................            </label></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
        <td class="style1" width="5%">&nbsp;</td>
          <td class="style1" width="95%"><b>เหตุผล</b>&nbsp;
            <% if isEmpty(gethidCPId) = True and getchkPrint <> "Print" then %>
            <label>
            <textarea name="txtReason" cols="90" rows="3" id="txtReason"></textarea>
            </label>
            <% else  response.write get_Audit_Reason end if %>            </td>
          
        </tr>
      </table></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>

  </table></form></td>
  <td width="40%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%" class="style1" align="center" ><%=GetSingleField("member","NAME"," where email='"&Session("Member")&"' ")&" "&GetSingleField("member","SURNAME"," where email='"&Session("Member")&"' ")%></td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><b><%=GetSingleField("member","NAME"," where email='"&Session("Member")&"' ")&" "&GetSingleField("member","SURNAME"," where email='"&Session("Member")&"' ")%></b></td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center"><%=Dateddmmyyyy%></td>
            <td>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้รับการตรวจ</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center">&nbsp;</td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="1">
              <tr>
                <td colspan="3" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;รับทราบผลติดตามซ้ำ วันที่..............................................</td>
                </tr>
              <tr>
                <td width="15%" class="style1">ลงชื่อ</td>
                <td width="60%" class="style1" align="center"><%=get_Audit_QMR_P2%></td>
                <td width="25%" class="style1">QMR</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><b><%=get_Audit_QMR_P2%></b></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1">วันที่</td>
                <td class="style1" align="center"><%=Dateddmmyyyy%></td>
                <td>&nbsp;</td>
              </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
    </tr>
  </table></td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
  <th align="right" class="style1">F-FDA-T-<% if PageName = "CAR" then %>16<% else %>17<% end if %> 
  (1-08/07/57) หน้า .../...</th>
</tr></table>
<%  if isEmpty(gethidCPId) = True and getchkPrint <> "Print" then %>
<div><input type="button" value="บันทึกส่วนที่ 3" onClick="javascript:{ document.frmEditPart3.submit();}"/>
&nbsp;&nbsp;&nbsp;
<input type="button" value="กลับหน้าหลัก" name="butCancel" onClick="javascript:{ window.location.href='FollowUp.asp';}" /></div>
<% else%>
<div><input type="button" value="พิมพ์" onClick="javascript:{ window.print();}"/>&nbsp;&nbsp;&nbsp;<input type="button" value="กลับหน้าหลัก" name="butCancel" onClick="javascript:{ window.close();}" /></div>
<% end if %>
<% end if %>
</body>
</html>
