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
  <tr><td>&nbsp;</td><td>&nbsp;</td></tr>
 <!-- <tr>
  <td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr><td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>�к��س�Ҿ (Quality System)</strong> ���¶֧ �к����������ͧ���㹡�äǺ�����л�Сѹ�س�Ҿ�ͧ˹��§ҹ ��觻�Сͺ仴����ç���ҧ�ͧͧ��� ˹�ҷ������Ѻ�Դ�ͺ �Ըմ��Թ��� ��кǹ��ô��Թ��� ��Ѿ�ҡ� ���͹ӹ�º�¡�ú����çҹ��ҹ�س�Ҿ任�Ժѵ� ��ô��Թ��ôѧ����Ǩ��繵�ͧ�Ѵ�����͡��� ��������ö���Թ����ѡ���к��س�Ҿ�����ҧ������� �������ö����������ҧ�ջ���Է���Ҿ</font>
  </td></tr></table>  </td>
  </tr>-->
  <tr><td>&nbsp;</td><td>&nbsp;</td></tr>
  <tr>
    <td colspan="2" align="center"><!--#include file="QS_Policy.asp"--></td>
  </tr>
  
<!--**********************************************************************************************************************************************************-->
 <tr><td colspan="2" align="center"><br /><br />
 <table width="100%" border="0">
 <tr><td align="left">
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>��Ъ�����ѹ��</b><br>
 </td></tr>
 <tr><td><img src="images/newthai2.gif">&nbsp;
 <a href="pdf/��˹����ͺ����͡�˹� ��.pdf" target="_blank">���ͺ����͡�˹��к��س�Ҿ�ͧ�ӹѡ�ҹ��С����������������: 2557 ��Шӻէ�����ҳ �.�. 2557 �ѹ��� 14 ����Ҥ� 2557 ���� 08.30 - 12.00 �. � ��ͧ��Ъ����� 6 �Ҥ�� 4 �ӹѡ�ҹ��С����������������
</a>
 
 </td></tr>
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