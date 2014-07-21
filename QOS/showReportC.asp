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
		 
		 get_Audit_QMR_P1 = RecC("Audit_QMR_P1")
		 
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
font-size:14px;
font-family:THSarabunPSK,Arial, Helvetica, sans-serif;
}
.style2 {
font-size:12px;
font-family:THSarabunPSK,Arial, Helvetica, sans-serif;
}
.style3 {font-size: 14px; font-family: THSarabunPSK,Arial, Helvetica, sans-serif; font-weight: bold; }
-->
</style>
</head>

<body>
<table width="100%" border="2" align="center" cellpadding="1" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
      <tr>
        <td width="10%"><img src="images/aoryor.jpg" width="50" height="50" /></td>
        <td align="center" width="65%"><font style="font-size:18px"><strong>รายงานการตรวจติดตามคุณภาพภายใน<br />
          (Audit Report)</strong></font></td>
      <td width="25%" class="style1" align="right"><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td class="style1"><label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart1" value="1"  onClick="ChangeJobresultGroup('','')"  <%=selectDepart%> disabled="disabled" />
    ระดับกรม</label></td>
        </tr>
        <tr>
          <td class="style1"><label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart1" value="2"  onClick="ChangeJobresultGroup('','')"  <%=selectSubDepart%> disabled="disabled" />
    ระดับหน่วยงาน</label></td>
        </tr>
        <tr>
          <td class="style1">Audit Report No. <span class="style2"><%=get_No_Car_Par%></span></td>
        </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td rowspan="2" width="60%" valign="top">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2"><b class="style1">ตรวจติดตามวันที่ :</b></td>
            <td width="70%" class="style1"><%=get_Audit_Date%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2" ><b class="style1">หน่วยงานที่รับการตรวจ :</b></td>
			<td width="70%" class="style1" ><%=getDepartmentname(get_Audit_Depart)%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="30%" class="style2"><b class="style1">ขอบเขตการตรวจ :</b></td>
            <td width="70%" class="style1"><%=get_M_Code&" "&get_M_Name%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td width="30%" class="style1" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="style1"><b>&nbsp;ผู้ตรวจติดตาม :</b></td>
              </tr>
            </table></td>
            <td width="70%"  class="style1" valign="middle"><table width="95%" border="0" cellspacing="0" cellpadding="5">
              <% if get_Audit_Name1 <> "" then %>
              <tr>
                <td width="5%" class="style2">1</td>
                <td width="65%" class="style1"><%=get_Audit_Name1%></td>
                <td width="30%" class="style1">หัวหน้าผู้ตรวจติดตาม</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name2 <> "" then %>
              <tr>
                <td class="style2">2</td>
                <td class="style1"><%=get_Audit_Name2%></td>
                <td class="style1">ผู้ตรวจติดตาม คนที่ 1</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name3 <> "" then %>
              <tr>
                <td class="style2">3</td>
                <td class="style1"><%=get_Audit_Name3%></td>
                <td class="style1">ผู้ตรวจติดตาม คนที่ 2</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name4 <> "" then %>
              <tr>
                <td class="style2">4</td>
                <td class="style1"><%=get_Audit_Name4%></td>
                <td class="style1">ผู้ตรวจติดตาม คนที่ 3</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name5 <> "" then %>
              <tr>
                <td class="style2">5</td>
                <td class="style1"><%=get_Audit_Name5%></td>
                <td class="style1">ผู้ตรวจติดตาม คนที่ 4</td>
              </tr>
              <% end if %>
              <% if get_Audit_Name6 <> "" then %>
              <tr>
                <td class="style2">6</td>
                <td class="style1"><%=get_Audit_Name6%></td>
                <td class="style1">ผู้ตรวจติดตาม คนที่ 5</td>
              </tr>
              <% end if %>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="style1" style="font-weight: bold">&nbsp;ข้อคิดเห็นของผู้ตรวจติดตาม</td>
          </tr>
          <tr>
            <td class="style1">&nbsp;<strong>จุดเด่น :</strong> <span class="style1"><%=get_Audit_Advantages%></span></td>
          </tr>
          <tr>
            <td class="style1">&nbsp;<strong>จุดด้อย :</strong> <%=get_Audit_Disadvantages%></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="4">
          <tr>
            <td width="30%" class="style1" valign="top"><b>ผลการตรวจติดตาม :</b></td>
            <td width="70%" class="style1"><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="style1"><label>
      <input type="checkbox" name="checkComplete" id="checkComplete" onClick="chkAllowCARPAR()" <%=chkSelectNotFind%>  disabled="disabled" />
      ไม่พบข้อบกพร่อง</label></td>
              </tr>
              <tr>
                <td class="style1">
                <label>
                  <input type="checkbox" name="chkCAR" id="chkCAR" disabled="disabled"  <% if RowCAR <> 0 then response.write "checked=""checked""" end if%> />
                  พบข้อบกพร่องหรือความไม่สอดคล้องขึ้นในระบบคุณภาพ (NC)<br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงออกใบ CAR จำนวน  <%=RowCAR%>  ใบ <% 'if RowCAR > 0 then %><% 'end if %><br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;เลขที่ดังนี้
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
                <td class="style1"><label>
                  <input type="checkbox" name="chkPAR" id="chkPAR" disabled="disabled" <% if RowPAR <> 0 then response.write "checked=""checked""" end if%> />
                  พบแนวโน้มที่จะเกิดข้อบกพร่องหรือความไม่สอดคล้องในระบบคุณภาพ (OBS)<br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงออกใบ PAR จำนวน  <%=RowPAR%>  ใบ  <% 'if RowPAR > 0 then %><%' end if %><br />
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;เลขที่ดังนี้
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
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="15" class="style1"><span style="font-weight: bold">&nbsp;หมายเหตุ :</span> เลขที่ใบ CAR และใบ PAR จะออกให้โดยอัตโนมัติ หากดำเนินการบันทึกข้อมูลในระบบ</td>
      </tr>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
    </table></td>
    <td width="40%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td class="style1">ผู้ตรวจติดตาม :</td>
      </tr>
       <% if get_Audit_Name1 <> ""  then %>
      <tr>
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="15%" class="style1">ลงชื่อ</td>
                <td width="60%">&nbsp;</td>
                <td width="25%" class="style1">หัวหน้าผู้ตรวจติดตาม</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><%=get_Audit_Name1%></td>
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
            </table>        </td>
      </tr>
      <% end if %>
      <% if get_Audit_Name2 <> ""  then %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td>
        <table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="15%" class="style1">ลงชื่อ</td>
                <td width="60%">&nbsp;</td>
                <td width="25%" class="style1">ผู้ตรวจติดตาม คนที่ 1</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><%=get_Audit_Name2%></td>
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
        </table>        </td>
      </tr>
      <% end if %>
      <% if get_Audit_Name3 <> ""  then %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม คนที่ 2</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><%=get_Audit_Name3%></td>
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
      <% end if %>
      <% if get_Audit_Name4 <> ""  then %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม คนที่ 3</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><%=get_Audit_Name4%></td>
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
      <% end if %>
      <% if get_Audit_Name5 <> ""  then %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม คนที่ 4</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><%=get_Audit_Name5%></td>
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
        </table></td>
      </tr>
      <% end if %>
	  <% if get_Audit_Name5 <> ""  then %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">ผู้ตรวจติดตาม คนที่ 5</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="1%" class="style1">(</td>
                  <td width="98%"class="style1" align="center"><%=get_Audit_Name6%></td>
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
        </table></td>
      </tr>
      <% end if %>
      <tr>
        <td><span style="height:5px"><img src="images/spacer.gif" alt=""  height="10"/></span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td class="style1"><span style="height:5px"><img src="images/spacer.gif" alt=""  height="5"/></span></td>
      </tr>
      <tr>
        <td class="style1">หน่วยงานที่รับการตรวจ :</td>
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
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="3" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;รับทราบรายงานการตรวจติดตาม วันที่..............................</td>
            </tr>
          <tr>
            <td width="15%" class="style1">ลงชื่อ</td>
            <td width="60%">&nbsp;</td>
            <td width="25%" class="style1">หัวหน้าหน่วยงาน</td>
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
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="3" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;รับทราบรายงานการตรวจติดตาม วันที่..............................</td>
          </tr>
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
                  <td width="98%"class="style1" align="center">&nbsp;</td>
                  <td width="1%" class="style1">)</td>
                </tr>
            </table></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td class="style1">วันที่</td>
            <td class="style1" align="center"><%=get_Audit_Date%></td>
            <td>&nbsp;</td>
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
  <th align="right" class="style1">F-FDA-T-15 (1-08/07/57) หน้า .../...</th>
</tr></table>
<br />
<br />
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/></div>
</body>
</html>
