<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<!--#include file="../../_ClassModule/SdbsResultCode.Class.asp"-->
<!--#include file="../../_ClassModule/SdbsShowData.Class.asp"-->
<!--#include file="../../_ClassModule/SdbsWebUI.Class.asp"-->
<!--#include file="../../_Function/sendmail.asp"-->
<script language="JavaScript" src="../../_java/main.js"></script>
<script language="javascript" src="../../_java/chk_date.js"></script>
<script language="JavaScript" src="../../_java/chkProfile.js"></script>
<script language="JavaScript" src="../../_java/previewImage1.js"></script>
<script language="JavaScript" src="../../_java/formattemp.js"></script>
<script language="JavaScript" src="../../_java/valid.js"></script>
<script language="JavaScript" src="../../_java/tree.js"></script>
<script language='javascript' src='../../_java/menus.js'></script>
<%
			cmd = request("cmd")
			Email = request("Email")&"@fda.moph.go.th"
			Fname = request("Fname")
			Lname= request("Lname")
			department2= request("department")
			
			IF cmd = 1 Then
						StrSQL = "Select * From Member Where Email='"&Email&"'"
						Set rs_c = server.createobject("adodb.recordset")
						rs_c.open StrSQL,con,1,3
							if rs_c.recordcount=0 Then
							If (department2="�ʨ.") Then
								act = 0
							Else
								act = 1
							End if
							rs_c.addnew
							rs_c("Email") = Email
							Randomize
							rs_c("KMPassword")=int(((99999 - 10000) + 1)*Rnd() + 10000)
							rs_c("Name") = Fname
							rs_c("SurName") = Lname
							rs_c("Depart") = department2
							rs_c("ACT") = act
							rs_c("ConfirmStatus")=0
							rs_c.update
							
									' ---------------------------------------Start block for send  password by email to user but now this function can not use---------------------------------------------
									'	sender 		= 		GetSingleField("tabconfigmail","Configmail","Where Id = 1")
									'	receiver 	=	 	Email
									'	subject		=		"Your Password For KM Member"
									'	body =  "<font face='Ms sans Serif' size='2'>Your Password is "&GetSingleField("Member","KMPassword"," Where Email='"&Email&"'")&"</font>"
									'	Call SendMail(sender,receiver,subject,body)
									' ---------------------------------------------------------------------End block for send email--------------------------------------------------------------------------------
							str_status = 1
							else
							str_status = 2
							end if 					
						Call CloseRecord(rs_c)						
						
			End IF
%>
<script language="JavaScript" src="../../_java/main.js" type="text/javascript"></script>
<script language="JavaScript">
function chk_val(f,lang)
{
	if (Trim(f.Email.value))
		WarnInput(f.Email,"  ������ ",lang)
	else if (Trim(f.Fname.value))
		WarnInput(f.Fname,"  ���ͨ�ԧ ",lang)
	else if (Trim(f.Lname.value))
		WarnInput(f.Lname,"  ���ʡ�� ",lang)
	else if (Trim(f.department.value))
		Warnselbox(f.department,"  ���˹�/˹��§ҹ",lang)
	else
		f.submit()
	

}
</script>
<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" > 

    <table width="780" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF"  align="center">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
              <td colspan="3"><img src="images/violet_head_web_2_1.jpg" width="1264"  /></td>
        </tr>
        <tr>
        <td width="10%" bgcolor="#f2e3ff">&nbsp;</td>
          <td width="80%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%
						if Isempty(session("member")) = True  then  	
				%>
				<tr>
				  <td bgcolor="#f2e3ff" >&nbsp;</td>
			  </tr>
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="1">
				  <input type="hidden" name="qs_chk" value="QS">
                  <tr>
                    <td bgcolor="#f2e3ff" align="center" ><font size="2" face="Ms Sans Serif" color="#000000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E-Mail&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input name="Email" type="text" size="20" value="yourname@fda.moph.go.th" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff">
                      &nbsp;���ʼ�ҹ :&nbsp;&nbsp;&nbsp;
                      <input name="Password" type="password" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff"  size="20"></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit"  value="�������к�"></td>
                  </tr>
                  </form>
				  <tr><td bgcolor="#f2e3ff">&nbsp;</td></tr>
				  <% end if %>
                  <tr>
                    <td>
                    <br>
                    <!--------------------------------------Block register-------------------------------------->
                    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u><b>��鹵͹���ŧ����¹</b></u><br>
                      <br>
                    </p>
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">1. ��͡��������´�ͧ��ҹ ���� E-mail ���������Ѻ ��. � 
                    �Ѩ�غѹ, ���ͨ�ԧ, ���ʡ�� ��е��˹�/˹��§ҹ����ѧ�Ѵ ŧ㹪�ͧ��ҧ � ���ú��ǹ��ж١��ͧ 
                    �ҡ��鹡����� <b>��ŧ</b>
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">
                    2. �к��зӡ�õ�Ǩ�ͺ��������������ʼ�ҹ����Ѻ Login ��Ҵ��͡����к��س�Ҿ�ͧ ��.�����ҹ�ҧ E-mail Address �ͧ��ҹ (��觵�ͧ�� E-mail Address �ͧ 
                    ��. ��ҹ��)
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">
                    3. ����ҹ�ӡ�õ�Ǩ�ͺ���ʼ�ҹ�¡���Դ E-mail �ͧ��ҹ ��Ш����ʼ�ҹ�ѧ������������Ѻ���� 
                    Login ��Ҵ��͡����к��س�Ҿ�ͧ ��. ����
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">
                    <!--4. �����ѧ�ҡ�����Ѥ���Ҫԡ FDA KM ���� ��ҹ����ö��������ʴ������Դ��� �š����¹������� 
                    �����������Ԩ������ҧ � �Ѻ���䫵� FDA KM ��<dd>-->
                    4. �óշ���ҹ��ŧ����¹������зӡ��ŧ����¹��� �к����բ�ͤ�������ҷ�ҹ��ŧ����¹�ҡ�͹���� 
                    ���ͻ�ͧ�ѹ�ѭ�Ң����ū�ӫ�͹  
                    <p class="text">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>�����˵�</u>: 
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">
                    1. �óշ���ҹŧ����¹���ǵ�ͧ��������ҹ���ҧ��觴�ǹ  ��سҵԴ��� �س��  �س�� ����������Ѿ�� 7254
                    <dd style="font-size:14px; font-family:Arial, Helvetica, sans-serif">
                    2. <b>੾�����˹�ҷ������ E-mail Address �ͧ��. ��� @fda.moph.go.th ��ҹ��</b>�������öŧ����¹���
                      <!--2. ���ʼ�ҹ�ѧ����Ƿ�ҹ����ö�ӡ������¹�ŧ������䫵� FDA KM ��Ǣ�� 
                    <b>����¹���ʼ�ҹ</b>--></p>
                    
                      <form method=post name=form>
                      <table width="390" height="220" align="center" cellpadding="0" cellspacing="0" background="<%=path_link%>admin/image/bg_NewUser.gif" class="text1" style="background-repeat: no-repeat">
                        <tr> 
                          <td height="210" class="FontNorMal"> <table border="0" cellpadding="3" cellspacing="0" align=center width="90%" class="text3">
                              <tr> 
                                <td colspan="2" align="right"><table border="0" cellpadding="0" cellspacing="0" align=center width="100%" class="text3">
                                    <tr> 
                                      <td width="30%" align="right">&nbsp;</td>
                                      <td width="70%"><font size="5"><strong><img src="../../_images/Register.jpg"></strong></font></td>
                                    </tr>
                                    <tr> 
                                      <td align="right">&nbsp;</td>
                                      <td width="80%">&nbsp;</td>
                                    </tr>
                                    <% if Isempty(str_status) = False then %>
                                    <tr> 
                                      <td colspan="2" ><!--<strong>��Ѥ���Ҫԡ</strong>-->&nbsp;&nbsp;&nbsp;&nbsp; 
                                        <%
                                                    IF str_status = 1 Then
                                                        response.write "<font color='green'>�����Ѥ����º���� ��سҵ�Ǩ�ͺ���ʼ�ҹ���������ͧ��ҹ</font>"
														response.write "<script type=""text/javascript"">setTimeout(function(){window.location.href=""http://elib.fda.moph.go.th/kmfda/_block/qos/""},3000)</script>"
                                                    elseif str_status = 2 then
                                                        response.write "<font color='#CC0000'><img src='_images/icon_err.gif' align='absmiddle'>&nbsp;&nbsp;��������������Ѥ���Ҫԡ���� �������ö��Ѥ����ա</font>"
                                                    end if
                                              %> </td>
                                    </tr>
                                    <% end if %>
                                  </table></td>
                              </tr>
                              <tr> 
                                <td height="19" align="right">������&nbsp;:&nbsp;</td>
                                <td><input type="text" name="Email"  size="10" class="FieldSkin" style="font-family: Ms sans serif; font-size: 8pt;">
                                  @fda.moph.go.th</td>
                              </tr>
                              <tr> 
                                <td align="right">���ͨ�ԧ :&nbsp;</td>
                                <td><input type="text" name="Fname"  size="20" class="FieldSkin" style="font-family: Ms sans serif; font-size: 8pt;"></td>
                              </tr>
                              <tr> 
                                <td width="46%" align="right">���ʡ��&nbsp;:&nbsp;</td>
                                <td width="54%"> <input type="text" name="Lname"  size="20" class="FieldSkin" style="font-family: Ms sans serif; font-size: 8pt;">                                </td>
                              </tr>
                              <tr> 
                                <td align="right">���˹�/˹��§ҹ&nbsp;:&nbsp;</td>
                                <td> <select name="department" class="textbox">
                                    <option value="">== ���͡ ==</option>
                                    <%call ListBoxArrayList(department,department2,1)%>
                                  </select> </td>
                              </tr>
                              <tr> 
                                <td colspan="2" align="right"><div align="center"> 
                                    <input type=button value="   ��ŧ  "   style="font-family: Ms sans serif; font-size: 8pt;" onClick="chk_val(this.form,'TH')">
                                    &nbsp; 
                                    <input type=button value="¡��ԡ"   style="font-family: Ms sans serif; font-size: 8pt;" >
                                    <input type="hidden" name="cmd" value="1">
                                  </div></td>
                              </tr>
                            </table></td>
                        </tr>
                      </table>
                      <% '=MenuBottomLogin(1,0,1,1)%> 
                    </form>
					<br /><div align="center"><a href="default.asp" style=" text-decoration:none">��Ѻ˹����ѡ</a></div><br /><br /><br />
                    <!------------------------------------------------------------------------------------------->
                    </td>
                  </tr>
          </table></td>
               <td width="10%" bgcolor="#f2e3ff">&nbsp;</td>
        </tr>
        <tr><td bgcolor="#f2e3ff" colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="3" align="center" valign="top" bgcolor="#f2e3ff">&nbsp;</td>
        </tr>
        <tr>
              <td  colspan="3" bgcolor="#d8b4fe" height="30px" align="center">
<strong><img src="http://elib.fda.moph.go.th/library/_images/icon_register.gif" width="31" height="21" align="absmiddle" /><font size="1">����颳й�� <%
numcount=Application("OnlineUser")
'numcount=right("0000000000000"&numcount,7)		
l=len(numcount)
for i=1 to l
num=mid(numcount,i,1)
display2=display2&num '"<img src=_images/counter/"&num&".gif align='absmiddle'>"
next
response.Write("<font color=000000><b>"&display2&"</b></font>")%></font></strong>
&nbsp;&nbsp;&nbsp;
<font size="2" face="Ms Sans Serif" color="#3300ff">&quot;੾�Т���Ҫ��� ��. ����ѧ��������ʼ�ҹ㹡����Ҵ��͡����к��س�Ҿ������駪��� ���ʡ�� �ҷ�� library@fda.moph.go.th , qsfda@fda.moph.go.th ��з�ҹ�����Ѻ��������ʼ�ҹ��Ѻ�ҧ�������������&quot;</font>
<%

set rs_counter=server.CreateObject("ADODB.recordset")
rs_counter.open "Select * From tabUniIP ",con,1,3


if session("check")="" then
session("check")="x"
rs_counter.addnew
rs_counter("ip")=request.ServerVariables("REMOTE_HOST")
rs_counter("date")=date
rs_counter.update
end if
rs_counter.close

%>
<a href="javascript:openWin('counter.asp','counter',500,500)" >Uni IP <img src="../../_images/folder_stats.gif" width="30" height="30" border=0></a>
</td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
