<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <!--<tr> 
    <td colspan="2" align="center" valign="top">
	<%'call openrecord(rs,"Select Desc from TabData Where Id=9",con,1,1)%>
	<%'=rs("Desc")%>
	<%'closerecord(rs)%>
      <br>
    </td>
  </tr>-->
  <tr>
  <td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr><td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>นโยบายคุณภาพ อย. : </strong>สำนักงานคณะกรรมการอาหารและยา พัฒนาบริการสู่การยอมรับระดับสากล ยึดมั่นในผล โปร่งใส เป็นธรรม
  </font>
  </td></tr></table>
  </td>
  </tr>
  <tr><td colspan="0">&nbsp;</td></tr>
  <tr>
  <td colspan="2" align="center" valign="top"><table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr><td>
  <font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>ระบบคุณภาพ (Quality System)</strong> หมายถึง ระบบที่เป็นเครื่องมือในการควบคุมและประกันคุณภาพของหน่วยงาน ซึ่งประกอบด้วยโครงสร้างขององค์กร หน้าที่ความรับผิดชอบ วิธีดำเนินการ กระบวนการดำเนินการ ทรัพยากร เพื่อนำนโยบายการบริหารงานด้านคุณภาพไปปฏิบัติ การดำเนินการดังกล่าวจำเป็นต้องจัดทำเป็นเอกสาร เพื่อสามารถดำเนินการรักษาระบบคุณภาพได้อย่างเหมาะสม และสามารถนำไปใช้ได้อย่างมีประสิทธิภาพ
  </font>
  </td></tr></table>
  </td>
  </tr>
  <tr><td colspan="0">&nbsp;</td></tr>
  <tr> 
    <td align="center" valign="top"><table width="39%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="/kmfda/_block/QOS/images/images_qs/qs_fda_03.jpg" width="302" height="45"></td>
        </tr>
        <tr> 
          <td background="/kmfda/_block/QOS/images/qs_fda_02.gif" height="100" valign="top"> <!--<table width="100%" border="0" cellpadding="3" cellspacing="0" class="text">
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
			<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" class="text"><tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" >
          <table width="90%" border="0" align="left" cellpadding="3" cellspacing="0" class="text">
          <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%">&nbsp;</td></tr>
          <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=101" target="_self">เอกสารระบบคุณภาพ(Quality System Documentation)</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=102" target="_self">ลักษณะและประโยชน์ของเอกสารระบบคุณภาพ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=104" target="_self">ขั้นตอนการจัดทำเอกสารระบบคุณภาพ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=108" target="_self">การควบคุมเอกสารและข้อมูล (Document and Data Control)</a></td>
          </tr>
          </table>
		  </td></tr></table>
			</td>
        </tr>
        <tr> 
          <td><img src="/kmfda/_block/QOS/images/qs_fda_03.gif" width="302" height="64"></td>
        </tr>
      </table></td>
    <td valign="top" align="center">
    <table width="39%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="/kmfda/_block/QOS/images/images_qs/qs_fda_05.jpg" width="302" height="45"></td>
        </tr>
        <tr> 
          <td background="/kmfda/_block/QOS/images/qs_fda_02.gif" height="125" valign="top">
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" class="text"><tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" >
          <table width="90%" border="0" align="left" cellpadding="3" cellspacing="0" class="text">
          <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%">&nbsp;</td></tr>
          <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/library5/fda_standard.pdf" target="_blank">มาตรฐานระบบคุณภาพ-ข้อกำหนดทั่วไปสำหรับสำนักงานคณะกรรมการอาหารและยา</a></td>
          </tr>
          </table>
          </td></tr></table>
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
			<!-- End Original code -->
			</td>
        </tr>
        <tr> 
          <td><img src="/kmfda/_block/QOS/images/qs_fda_03.gif" width="302" height="64"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
