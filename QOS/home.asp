<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>�͡����к��س�Ҿ ��.</title>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <!--<tr> 
    <td colspan="2" align="center" valign="top">
	<%'call openrecord(rs,"Select Desc from TabData Where Id=9",con,1,1)%>
	<%'=rs("Desc")%>
	<%'closerecord(rs)%>
      <br>
    </td>
  </tr>-->
 <!--<tr>
  <td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr><td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>��º�¤س�Ҿ ��. : </strong>�ӹѡ�ҹ��С���������������� �Ѳ�Һ�ԡ�����������Ѻ�дѺ�ҡ� �ִ���㹼��ç�� �繸���</font>
  </td></tr></table>
  </td>
  </tr>-->
  
 <!-- <tr>
  <td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr><td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>�к��س�Ҿ (Quality System)</strong> ���¶֧ �к����������ͧ���㹡�äǺ�����л�Сѹ�س�Ҿ�ͧ˹��§ҹ ��觻�Сͺ仴����ç���ҧ�ͧͧ��� ˹�ҷ������Ѻ�Դ�ͺ �Ըմ��Թ��� ��кǹ��ô��Թ��� ��Ѿ�ҡ� ���͹ӹ�º�¡�ú����çҹ��ҹ�س�Ҿ任�Ժѵ� ��ô��Թ��ôѧ����Ǩ��繵�ͧ�Ѵ�����͡��� ��������ö���Թ����ѡ���к��س�Ҿ�����ҧ������� �������ö����������ҧ�ջ���Է���Ҿ</font>
  </td></tr></table>  </td>
  </tr>-->
  
  <tr><td align="center"><!--<img src="images/head1.jpg" alt="" height="95" />-->&nbsp;</td></tr>
  
  <tr>
    <td colspan="2" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>��º�¤س�Ҿ</b></td>
  </tr>
  <tr><td>&nbsp;</td></tr>
   <tr>
    <td colspan="2" align="left" height="60">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/��º��.jpg" border="0" width="650" /></td>
  </tr>
 <tr><td>&nbsp;</td></tr>
  
<!--********************************************************************************************************************************************-->
	<tr><td colspan="2"><table width="100%" cellpadding="2" cellspacing="2" border="0">
    <tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>�Ԩ�����к��س�Ҿ��ѹ���</b></td></tr>
    <tr><td>&nbsp;</td></tr>
    <%
	Dateddmmyyyy=Now()
	Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
	'if isEmpty(Request.QueryString("date")) = false then
	set Rec = Server.CreateObject("ADODB.RECORDSET")
	set RecAc = Server.CreateObject("ADODB.RECORDSET")
	FDate = Month(Request.QueryString("date"))&"/"&Day(Request.QueryString("date"))&"/"&Year(Request.QueryString("date"))
	FDate = Datemmddyyyy
	SQL = "Select * from Tb_Book where B_Flag = True and B_StartDate=#"&FDate&"# and B_EndDate >= #"&FDate&"#"
	'SQLActivity = "Select * from Tb_Activity where  A_Flag = True and A_Date=#"&FDate&"# or A_StartDate <= #"&FDate&"# and A_EndDate >= #"&FDate&"#"
	SQLActivity = "Select * from Tb_Activity where  A_Flag = true  and  A_StartDate <= #"&FDate&"# and A_EndDate >= #"&FDate&"#"
	'response.write SQLActivity&"<br />"
	Rec.open SQL,ConActivity,1,3
	if Rec.RecordCount <= 0 then
	'response.write "No Data"
	end if 
	RecAc.open SQLActivity,ConActivity,1,3
	if RecAc.RecordCount <= 0 then
	'response.write "No Data"
	end if
	
%>
	<%
	while not RecAc.EOF
%>
<TR><TD ALIGN="center">
<table width="85%" cellpadding="3" cellspacing="0" border="0" bgcolor="#FFFF99">
<tr>
<td width="100%" colspan="2" align="left" bgcolor="#b27ee0" style="font-size:12px;COLOR=#FFFF00"><b><%=RecAc("A_Topic")%></b></td>
</tr>
<tr>
  <td align="left" width="90%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="font-size:12px">��ͧ��Ъ�� / ʶҹ��� :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF; font-size:12px"><%=RecAc("A_RoomName")%></span></td>
  <td width="10%" rowspan="4" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<!--<tr><td width="80%" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��Ǣ�� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%'=RecAc("A_Topic")%></td></tr>-->
<tr>
  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="font-size:12px">�ѹ���  :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF; font-size:12px"><%=RecAc("A_StartDate")%><% if RecAc("A_EndDate") <> "" and RecAc("A_StartDate") <> RecAc("A_EndDate") then response.write "&nbsp;&nbsp;&nbsp;<b style=""color:#000000"">�֧</b>&nbsp;&nbsp;&nbsp;"&RecAc("A_EndDate") end if %></span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="font-size:12px">����  :</strong>&nbsp;
 <span style="color:#0000FF; font-size:12px">
  <%
if Minute(RecAc("A_StartTime")) >=0 and Minute(RecAc("A_StartTime")) < 10  then
response.write Hour(RecAc("A_StartTime"))&":0"&Minute(RecAc("A_StartTime"))
elseif Minute(RecAc("A_StartTime")) > 10 then
response.write Hour(RecAc("A_StartTime"))&":"&Minute(RecAc("A_StartTime"))
else
response.write ""
end if 
%>
</span>&nbsp;&nbsp;
<strong style="font-size:12px">�֧</strong>&nbsp;&nbsp;
<span style="color:#0000FF; font-size:12px">
<%
if  Minute(RecAc("A_EndTime")) >=0 and Minute(RecAc("A_EndTime")) < 10 then
response.write Hour(RecAc("A_EndTime"))&":0"&Minute(RecAc("A_EndTime"))
elseif  Minute(RecAc("A_EndTime")) > 10 then
response.write Hour(RecAc("A_EndTime"))&":"&Minute(RecAc("A_EndTime"))
else
response.write ""
end if
%>
</span></td>
</tr>

<tr>
  <td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="font-size:12px">���ͼ���Ѻ�Դ�ͺ / ˹��§ҹ :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF; font-size:12px"><%=RecAc("A_Name")%></span></td></tr>
<tr><td align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong style="font-size:12px">����Դ��� :</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#0000FF; font-size:12px"><%=RecAc("A_Tel")%></span></td></tr>
</table>
</TD></TR>
<%
	RecAc.MoveNext
	wend
	
%>
</table>
<!----------------------------------------------------------------------------------------------Start show blank block------------------------------------------------------------------------------------------------------>
                      <% if RecAc.RecordCount < 1  then %>
                      <table width="85%" cellpadding="3" cellspacing="0" border="0" bgcolor="#FFFF99" align="center">
                        <tr><td width="100%"  align="left" bgcolor="#b27ee0" style="font-size:12px">&nbsp;</td>
                        </tr>
                        <tr><td align="center" class="textsmall">---------------------------����աԨ������ѹ���----------------------------</td></tr>
                        </table>
					  <% 
					  end if 
					  	RecAc.close
						set RecAc = Nothing
						Rec.close
						set Rec = Nothing
					  %>
<!----------------------------------------------------------------------------------------------End show blank block------------------------------------------------------------------------------------------------------->
    </td></tr>
<!--********************************************************************************************************************************************-->
 <tr><td colspan="2" align="center"><br /><br />
 <table width="100%" border="0">
 <tr><td align="left">
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>��Ъ�����ѹ��</b><br>
 </td></tr>
 <tr><td>&nbsp;&nbsp;<img src="images/newthai2.gif">&nbsp;
 <a href="pdf/��ԷԹ��õ�Ǩ.pdf" target="_blank">��˹���õ�Ǩ�Դ����س�Ҿ���㹢ͧ�ӹѡ�ҹ��С���������������һ�Шӻ�������ҳ �.�.2557</a>
 </td></tr>
 <!--<tr><td><img src="images/newthai2.gif">&nbsp;
 <a href="pdf/��˹����ͺ����͡�˹� ��.pdf" target="_blank">���ͺ����͡�˹��к��س�Ҿ�ͧ�ӹѡ�ҹ��С����������������: 2557 ��Шӻէ�����ҳ �.�. 2557 �ѹ��� 14 ����Ҥ� 2557 ���� 08.30 - 12.00 �. � ��ͧ��Ъ����� 6 �Ҥ�� 4 �ӹѡ�ҹ��С����������������
</a>
 </td></tr>
  <tr>
   <td><img src="images/newthai2.gif" alt="" /> <a href="pdf/Binder21.pdf" target="_blank">����ª��ͼ���������ͺ����ѡ�ٵ��ҵðҹ�к��س�Ҿ����Ѻ��õ�Ǩ�ͺ����Ѻ�ͧ ��Шӻէ�����ҳ  �.�. 2557 (Introduction to ISO/IEC 17021:2011) �ѹ��� 24-25 �Զع�¹ 2557 � �ç��� ��������ù� ��� �.�Ժ����ʧ���� �.���ͧ �.�������</a></td>
 </tr>
 <!--<tr>
   <td><img src="images/newthai2.gif" alt="" /> <a href="pdf/Binder65.pdf" target="_blank">����ª��ͼ���������ͺ����ѡ�ٵ��ҵðҹ�к��س�Ҿ����Ѻ��õ�Ǩ�ͺ����Ѻ�ͧ ��Шӻէ�����ҳ  �.�. 2557 (Introduction to ISO/IEC 17065:2012) �ѹ��� 16-17 �Զع�¹ 2557 � �ç��� ��������ù� ��� �.�Ժ����ʧ���� �.���ͧ �.�������</a></td>
 </tr>-->
<!-- <tr>
   <td><img src="images/newthai2.gif" alt="" /> <a href="pdf/��ª��ͼ���������ͺ�� Lead auditor_binder.pdf" target="_blank">����ª��ͼ���������ͺ����ѡ�ٵ����˹�Ҽ���Ǩ�����Թ�к������äس�Ҿ  (Lead Auditor) ��Шӻէ�����ҳ  �.�. 2557 �ѹ��� 9-12 �Զع�¹ 2557</a></td>
 </tr>
 <tr>
   <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="pdf/��˹����ͺ��_Lead Auditor.pdf" target="_blank">���ͺ����ѡ�ٵ����˹�Ҽ���Ǩ�����Թ�к������äس�Ҿ (Quality Management System Lead auditor) ��Шӻէ�����ҳ �.�. 2557 �ѹ��� 9 - 10 �Զع�¹ 2557 � �ç��� ��������ù� ��� �.�Ժ��ʧ���� �.���ͧ �.������� ����ѹ��� 11 - 12 �Զع�¹ 2557 � ��ͧ��Ъ�� 1 ��� 6 �֡�Թ�ع��ع���¹���ʾ�Դ �ӹѡ�ҹ��С����������������</a> </td>
 </tr>-->
 <!--<tr><td><img src="images/newthai2.gif">&nbsp;
 <a href="pdf/��˹����ͺ��_iso 17065.pdf" target="_blank">���ͺ����ѡ�ٵä���������ͧ������ǡѺ��͡�˹�����Ѻ˹����Ѻ�ͧ��Ե�ѳ�� �к� ��к�ԡ�� (Introduction to ISO/IEC 17065:2012) ��Шӻէ�����ҳ �.�. 2557 �ѹ��� 16 - 17 �Զع�¹ 2557 � �ç��� ��������ù� ��� �.�Ժ��ʧ���� �.���ͧ �.�������</a>
 </td></tr>
 <tr><td><img src="images/newthai2.gif">&nbsp;
 <a href="pdf/��˹����ͺ��_iso 17021.pdf" target="_blank">���ͺ����ѡ�ٵä���������ͧ������ǡѺ��͡�˹�����Ѻ˹����Ѻ�ͧ�к���èѴ��� (Introduction to ISO/IEC 17021:2011) ��Шӻէ�����ҳ �.�. 2557 �ѹ��� 24 - 25 �Զع�¹ 2557 � �ç��� ��������ù� ��� �.�Ժ��ʧ���� �.���ͧ �.�������</a>
 </td></tr>-->
 </table>
<!-- <table  width="50%" align="center" cellpadding="3" cellspacing="0" border="0" >
 <tr><td>&nbsp;</td></tr>
 <tr><td align="center">
 <img src="images/newthai2.gif"> <a href="��ͺ��С�ȹ�ºѵ�.pdf" target="_blank">Click DownLoad ��º�¤س�Ҿ</a>
 </td></tr>
 <tr><td>&nbsp;</td></tr>
 </table>-->
 </td></tr>
  
<!--**********************************************************************************************************************************************************-->
  
  <tr> 
    <td align="center" valign="top">&nbsp;<!--<table width="39%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/images_qs/qs_fda_03.jpg" width="302" height="47"></td>
        </tr>
        <tr> 
          <td  height="100" valign="top">&nbsp; <!--<table width="100%" border="0" cellpadding="3" cellspacing="0" class="text">
              <tr> 
                <td> <%call OpenRecord(rs2,"Select top 1 * from TabData_L2 Where Id_L1=9 Order By Numberlist Desc",con,1,1)
	for r2=1 to rs2.recordcount
	
	 					call InsertIink(rs2,link2,endlink2,"ID_L2")%> <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" class="text">
                    <tr> 
                      <td width="1%">&nbsp;&nbsp;</td>
                      <td width="100%"><b> 
                        <%'=link2%>
                        <%'=rs2("Topic")%>
                        <%'=endlink2%>
                        </b></td>
                    </tr>
                    <tr > 
                      <td></td>
                      <td> <%call OpenRecord(rs3,"Select  * from TabData_L3 Where Id_L2=57 Order By Numberlist Desc",con,1,1)
		   	for r3=1 to rs3.recordcount
			
	 					call InsertIink(rs3,link3,endlink3,"ID_L3")%> <table width="100%" border="0" cellpadding="3" cellspacing="0" class="text">
                          <tr> 
                            <td width="1%"><img src="<%=path_link%>_images/arrowL2.gif" width="15" height="13">&nbsp;&nbsp;</td>
                            <td width="100%"><%=link3%><%=rs3("Topic")%><%=endlink3%></td>
                          </tr>
                          <tr> 
                            <td></td>
                            <td> <%call OpenRecord(rs4,"Select * from TabData_L4 Where Id_L3="&rs3("Id_L3")&" Order By Numberlist Desc",con,1,1)
						for r4=1 to rs4.recordcount
						
	 					call InsertIink(rs4,link4,endlink4,"ID_L4")
						%> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                                <tr> 
                                  <td width="1%"><img src="<%=path_link%>_images/arrowL3.gif" width="15" height="13">&nbsp;&nbsp;</td>
                                  <td width="100%"><%=link4%><%=rs4("Topic")%><%=endlink4%> </td>
                                </tr>
                              </table>
                              <%rs4.movenext
				  next
					   closerecord(rs4)%> </td>
                          </tr>
                        </table>
                        <%rs3.movenext
			next
		   closerecord(rs3)%> </td>
                    </tr>
                  </table>
                  <%rs2.movenext
	next
	closerecord(rs2)%> </td>
              </tr>
            </table>-->
			<!--<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" class="text"><tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" >
          <table width="90%" border="0" align="left" cellpadding="3" cellspacing="0" class="text">
          <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%">&nbsp;</td></tr>
          <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=101" target="_self">�͡����к��س�Ҿ (Quality System Documentation)</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=102" target="_self">�ѡɳ���л���ª��ͧ�͡����к��س�Ҿ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=104" target="_self">��鹵͹��èѴ���͡����к��س�Ҿ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=108" target="_self">��äǺ����͡�����Т����� (Document and Data Control)</a></td>
          </tr>
          </table>		  </td></tr></table>		  </td>
        </tr>
        <tr> 
          <td><img src="images/qs_fda_03.gif" width="302" height="64"></td>
        </tr>
      </table>--></td>
    <td valign="top" align="center">&nbsp;<!--
    <table width="39%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/images_qs/qs_fda_05.jpg" width="302" height="47"></td>
        </tr>
        <tr> 
          <td background="images/qs_fda_02.gif" height="125" valign="top">
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" class="text"><tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" >
          <table width="90%" border="0" align="left" cellpadding="3" cellspacing="0" class="text">
          <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%">&nbsp;</td></tr>
          <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/library5/fda_standard.pdf" target="_blank">�ҵðҹ�к��س�Ҿ-��͡�˹����������Ѻ�ӹѡ�ҹ��С����������������</a></td>
          </tr>
          </table>          </td></tr></table>
          <!-- Start Original code -->
		  <!--<table width="100%" border="0" cellpadding="3" cellspacing="0" class="text">
              <tr> 
                <td> <%'call OpenRecord(rs2,"Select top 1 * from TabData_L2 Where Id_L1=9 Order By Numberlist Desc",con,1,1)
	'for r2=1 to rs2.recordcount
	
	 					'call InsertIink(rs2,link2,endlink2,"ID_L2")%> <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" class="text">
                    <tr> 
                      <td width="1%">&nbsp;&nbsp;</td>
                      <td width="100%"><b> 
                        <%'=link2%>
                        <%'=rs2("Topic")%>
                        <%'=endlink2%>
                        </b></td>
                    </tr>
                    <tr > 
                      <td></td>
                      <td> <%'call OpenRecord(rs3,"Select  * from TabData_L3 Where Id_L2=58 Order By Numberlist Desc",con,1,1)
		   	'for r3=1 to rs3.recordcount
			
	 					'call InsertIink(rs3,link3,endlink3,"ID_L3")%> <table width="100%" border="0" cellpadding="3" cellspacing="0" class="text">
                          <tr> 
                            <td width="1%"><img src="<%'=path_link%>_images/arrowL2.gif" width="15" height="13">&nbsp;&nbsp;</td>
                            <td width="100%"><%'=link3%><%'=rs3("Topic")%><%'=endlink3%></td>
                          </tr>
                          <tr> 
                            <td></td>
                            <td> <%'call OpenRecord(rs4,"Select * from TabData_L4 Where Id_L3="&rs3("Id_L3")&" Order By Numberlist Desc",con,1,1)
						'for r4=1 to rs4.recordcount
						
	 					'call InsertIink(rs4,link4,endlink4,"ID_L4")
						%> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                                <tr> 
                                  <td width="1%"><img src="<%'=path_link%>_images/arrowL3.gif" width="15" height="13">&nbsp;&nbsp;</td>
                                  <td width="100%"><%'=link4%><%'=rs4("Topic")%><%'=endlink4%> </td>
                                </tr>
                              </table>
                              <%'rs4.movenext
				  'next
					   'closerecord(rs4)%> </td>
                          </tr>
                        </table>
                        <%'rs3.movenext
			'next
		   'closerecord(rs3)%> </td>
                    </tr>
                  </table>
                  <%'rs2.movenext
	'next
	'closerecord(rs2)%> </td>
              </tr>
            </table>-->
			<!-- End Original code -->		  </td>
        <!--</tr>
        <tr> 
          <td><img src="images/qs_fda_03.gif" width="302" height="64"></td>
        </tr>
      </table> -->   </td>
  </tr>
</table>
</body>
</html>