<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if

'	if isEmpty(session("member")) = True then
'		Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
'	end if
		dim getR_Id,chkComport,chkUnComport,SOP_Q,SOP_P,SOP_W,showExpectedDate,PageName
		chkComport=""
		chkUnComport=""
		SOP_Q=""
		SOP_P=""
		SOP_W=""
		PageName = ""
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
		 get_Audit_Source_Details = RecC("Audit_Source_Details")
		 
		 
		 get_Audit_Depart=RecC("Audit_Depart")
		 get_Audit_SubDepart=RecC("Audit_SubDepart")
		 get_Audit_SubDepartElseName=RecC("Audit_SubDepartElseName")
		 get_M_Code=RecC("M_Code")
		 get_M_Name=RecC("M_Name")
		 get_Audit_Name1=RecC("Audit_Name1")
		 get_Audit_Name2=RecC("Audit_Name2")
		 get_Audit_Name3=RecC("Audit_Name3")
		 get_Audit_Name4=RecC("Audit_Name4")
		 
		 get_Audit_Descript=RecC("Audit_Descript")
		 get_Audit_Advantages=RecC("Audit_Advantages")
		 get_Audit_Disadvantages=RecC("Audit_Disadvantages")
		 get_Audit_Defect = RecC("Audit_Defect")
		 get_Audit_License_P1 = RecC("Audit_License_P1")
		 get_Audit_QMR_P1 = RecC("Audit_QMR_P1")
		 
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
<title>��§ҹ����������</title>
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
        <td align="center" width="70%"><font style="font-size:18px"><strong>㺢���黯Ժѵԡ�����<br />
          <% if PageName = "CAR" then %>
          (Corrective Action Request : CAR)
          <% else %>
           (Preventive Action Request : PAR)
		  <% end if %>
          </strong></font></td>
      <td width="20%" class="style1" align="right"><table width="90%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td class="style1"><b><%=PageName%> No.</b> <%=get_No_Car_Par%></td>
        </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="20%" class="style1"><b>��ǹ���1 : ����͡� <%=PageName%></b></td>
            <td width="80%" class="style1">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="20%" valign="top" class="style1"><b>����� :</b></td>
            <td width="80%" class="style1" ><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkAuditIn" id="chkAuditIn"  <% if get_Audit_Source = "1" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  ��õ�Ǩ�Դ����س�Ҿ����</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkAuditOut" id="chkAuditOut" <% if get_Audit_Source = "2" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  ��õ�Ǩ�����Թ�ҡ��¹͡</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkMeetingReview" id="chkMeetingReview" <% if get_Audit_Source = "3" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  ��û�Ъ�����ǹ�½��º�����</label></td>
              </tr>
              <tr>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkWork" id="chkWork" <% if get_Audit_Source = "4" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  ��û�Ժѵԧҹ</label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkRequest" id="chkRequest" <% if get_Audit_Source = "5" then response.write "checked=""checked""" end if%>  disabled="disabled" />
                  �����ͧ���¹�ҡ <% if get_Audit_Source = "5" then response.write get_Audit_Source_Details end if %></label></td>
                <td class="style1" ><label>
                  <input type="checkbox" name="chkElse" id="chkElse" <% if get_Audit_Source = "6" then response.write "checked=""checked""" end if%> disabled="disabled" />
                  ���� <% if get_Audit_Source = "6" then response.write get_Audit_Source_Details end if %></label></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td width="20%" class="style1" ><b>˹��§ҹ��辺 :</b></td>
			<td width="80%" class="style1" ><%=getDepartmentname(get_Audit_Depart)%></td>
          </tr>
        </table></td>
      </tr>
      <tr><td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="20%" class="style1"><b>�Ԩ��������Ǩ (�����͡���/����)</b></td>
          <td width="80%" class="style1"><%=get_M_Code&" "&get_M_Name%></td>
        </tr>
      </table></td></tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td width="20%" class="style1"><b><% if PageName = "CAR" then %>��ͺ����ͧ��辺 :<% else %>�������ͺ����ͧ��辺 :<% end if%></b></td>
            <td width="80%" class="style1">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr><td height="15"><table width="90%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td class="style1">&nbsp;&nbsp;&nbsp;&nbsp;<%=get_Audit_Defect%></td>
        </tr>
      </table></td></tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td><table width="90%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr>
            <td width="50%"><table width="90%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="40%" class="style1"><b>ŧ���ͼ���͡� <%=PageName%></b></td>
                <td width="59%">&nbsp;</td>
                <td width="1%" class="style1">&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><%=get_Audit_License_P1%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1"><b>�ѹ���</b></td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
            <td width="50%"><table width="90%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="30%" class="style1"><b>ŧ���� QMR</b></td>
                <td width="70%">&nbsp;</td>
                <td width="5%" class="style1">&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><table width="90%" border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td width="1%" class="style1">(</td>
                      <td width="98%"class="style1" align="center"><%=get_Audit_QMR_P1%></td>
                      <td width="1%" class="style1">)</td>
                    </tr>
                </table></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td class="style1"><b>�ѹ���</b></td>
                <td class="style1" align="center">&nbsp;<%=get_Audit_Date%></td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="30%" class="style1"><b>��ǹ��� 2 : 
              <% if PageName = "CAR" then%>�����Թ������<% else %>�����Թ��û�ͧ�ѹ<% end if %></b></td>
          <td width="70%" class="style1">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20%" class="style1"><b>���˵آͧ�ѭ�� :</b></td>
          <td width="80%" class="style1"><%=get_Audit_Problem%></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="25%" class="style1"><b>
            <% if PageName = "CAR" then%>�Ƿҧ���<% else %>�Ƿҧ��ͧ�ѹ<% end if %> :</b></td>
          <td width="75%" class="style1"><%=get_Audit_Protect%></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="50%" class="style1" valign="top"><b>��˹��������������ѹ��� : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=get_Audit_Finish_Date%></b></td>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td width="40%" class="style1"><b>ŧ����
                <%if PageName = "CAR" then %>�����Թ������<% else %>�����Թ��û�ͧ�ѹ<% end if%></b></td>
              <td width="59%">&nbsp;</td>
              <td width="1%" class="style1">&nbsp;</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center"><b><%=get_Audit_Edit_Name%></b></td>
                    <td width="1%" class="style1">)</td>
                  </tr>
              </table></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="style1"><b>�ѹ���</b></td>
              <td class="style1" align="center"><%=get_Audit_Date2%></td>
              <td>&nbsp;</td>
            </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td width="40%" class="style1"><b>ŧ�������˹��˹��§ҹ</b></td>
              <td width="59%">&nbsp;</td>
              <td width="1%" class="style1">&nbsp;</td>
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
              <td class="style1"><b>�ѹ���</b></td>
              <td class="style1" align="center"><%=get_Audit_Date2%></td>
              <td>&nbsp;</td>
            </tr>
          </table></td>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="2">
            <tr>
              <td width="40%" class="style1"><b>ŧ���� QMR</b></td>
              <td width="59%">&nbsp;</td>
              <td width="1%" class="style1">&nbsp;</td>
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
              <td class="style1"><b>�ѹ���</b></td>
              <td class="style1" align="center"><%=get_Audit_Date2%></td>
              <td>&nbsp;</td>
            </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>
  </table></td>
  </tr>
  <tr>
  <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="20%" class="style1"><b>��ǹ��� 3 :
            ����͡� <%=PageName%>
          </b></td>
          <td width="80%" class="style1">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20%">�š�õ�Ǩ�Դ���</td>
          <td width="40%"><label>
            <input type="checkbox" name="chkAccept" id="chkAccept" disabled="disabled"  <% if get_Audit_Accept = "1" then  response.write " checked=""checked"" " end if %> />
            ����Ѻ</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <label>
            <input type="checkbox" name="chkClose" id="chkClose"  disabled="disabled" <% if get_Audit_OpenClose = "2" then  response.write " checked=""checked"" " end if %> />
            �Դ <% if PageName = "CAR" then %>CAR<% else %>PAR<% end if %> No. :</label></td>
          <td width="40%">&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><label>
            <input type="checkbox" name="chkNotAccept" id="chkNotAccept" disabled="disabled" <% if get_Audit_Accept = "2" then  response.write " checked=""checked"" " end if %> />
            �������Ѻ</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <label>
            <input type="checkbox" name="chkOpen" id="chkOpen" disabled="disabled" <% if get_Audit_OpenClose = "1" then  response.write " checked=""checked"" " end if %> />
            �Դ <% if PageName = "CAR" then %>CAR<% else %>PAR<% end if %> No. :</label></td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr>
          <td width="20%">&nbsp;</td>
          <td width="80%"><b>�˵ؼ�</b> &nbsp;&nbsp;&nbsp;<%=get_Audit_Reason%></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="34%" class="style1"><b>ŧ���ͼ���͡�<% if PageName = "CAR" then %>CAR<% else %>PAR<% end if %></b></td>
              <td width="65%">&nbsp;</td>
              <td width="1%" class="style1">&nbsp;</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center"><%=get_Audit_License_P3%></td>
                    <td width="1%" class="style1">)</td>
                  </tr>
              </table></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="style1"><b>�ѹ���</b></td>
              <td class="style1" align="center"><%=get_Audit_Date3%></td>
              <td>&nbsp;</td>
            </tr>
          </table></td>
          <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="40%" class="style1"><b>ŧ���� QMR</b></td>
              <td width="59%">&nbsp;</td>
              <td width="1%" class="style1">&nbsp;</td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr>
                    <td width="1%" class="style1">(</td>
                    <td width="98%"class="style1" align="center"><%=get_Audit_QMR_P3%></td>
                    <td width="1%" class="style1">)</td>
                  </tr>
              </table></td>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td class="style1"><b>�ѹ���</b></td>
              <td class="style1" align="center"><%=get_Audit_Date3%></td>
              <td>&nbsp;</td>
            </tr>
          </table></td>
        </tr>
      </table></td>
    </tr>

  </table></td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
  <th align="right" class="style1">F-FDA-T-<% if PageName = "CAR" then %>16<% else %>17<% end if %> (0-30/09/56) ˹�� 1/1</th>
</tr></table>
<div><input type="button" value="Print" onClick="javascript:{ window.print();}"/></div>
</body>
</html>
