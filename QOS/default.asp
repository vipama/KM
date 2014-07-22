<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
'===========================Login Member=====================================
Email=BlockSqlConjection(request("Email"))
password=BlockSqlConjection(request("password"))
login=request("login")
chk_login = Request("chk_login")
qs_chk = Request.Form("qs_chk")
if login="1"or chk_login="1"  then
		set rs_login=server.createobject("ADODB.recordset")
		rs_login.open "Select * from member Where Email='"&Email&"' And KMPassword='"&password&"' and ConfirmStatus=1",con,1,1
		if rs_login.recordcount<>0 then
				session("member")=rs_login("Email")
				session("act") = rs_login("ACT")
				if  Patchdate2(rs_login("LastLogin"),0,"EN")<>Patchdate2(Date,0,"EN") Then call AddScoreBoard(session("member"),"logincount")
					con.execute("Update Member Set LastLogin=now() Where Email='"&session("member")&"'")
					con.execute("INSERT INTO TabMemberLog (memberid,[date]) VALUES ("&rs_login("id")&",'"&FormatDate(Patchdate2(date,0,"EN"),2)&"')")
					message=""
				else
					message="????????????????"
				end if		
				
				if qs_chk = "QS"	 and rs_login.recordcount<>0 then
					'Response.Redirect "http://filing.fda.moph.go.th/kmfda/default.asp?page=doc"	
					Response.Redirect "http://filing.fda.moph.go.th/kmfda/_block/qos/"
				else
					Response.Redirect "http://filing.fda.moph.go.th/kmfda/_block/qos/"
				end if
		closerecord(rs_login)
elseif login="0" or chk_login="0" then
session("login")=0
session.Contents.Remove("member")
Response.Redirect "default.asp"
End if 
'==============================end of login====================================
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>ระบบคุณภาพของ อย.</title>

<link href="../../_Css/Styte.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
function Middle_showhide(elementName)
{
if (eval("document.all."+elementName+".style.display")=='')
eval("document.all."+elementName+".style.display='none'")
else
eval("document.all."+elementName+".style.display=''")
}
</script>
<script language="JavaScript" src="../../_java/main.js"></script>
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<link href="SpryAssets/SpryMenuBarVertical.css" rel="stylesheet" type="text/css" />
</head>
<!--------------------------------------------------------Start Code for get flag show  system  popup ------------------------------------------------------>
<%
	Set ConPlanweb = Server.CreateObject("ADODB.Connection")
	PlanwebStr_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source=E:\Planweb\_db\DBPopup.mdb"
	ConPlanweb.open PlanwebStr_Connect
	Set   RecFlagpopupshow = Server.CreateObject("ADODB.RECORDSET")
	SQL_Flagpopupshow = "Select * from tb_popup "
	RecFlagpopupshow.open SQL_Flagpopupshow,ConPlanweb,1,3
	 While Not RecFlagpopupshow.EOF
	 	getID =  RecFlagpopupshow("ID")
		getPopName = RecFlagpopupshow("PopName")
		getPopValue =  RecFlagpopupshow("PopValue")
		getPopObject = RecFlagpopupshow("PopObject")
		getSessionName = RecFlagpopupshow("SessionName")
		getPath = RecFlagpopupshow("Path")
		getFlagShowDes =  RecFlagpopupshow("FlagShowDes")
		if getFlagShowDes = 1 then
			getPopName = getPopName&"?id="&getID
			'response.write getPopName
		end if
		 
		 if getPopValue <> 0 and getPopValue <> "" then
				if  getPopValue = 1 then
		  			if session(getSessionName) = ""  then
		 				session(getSessionName)="show"
%>
		<script>
var <%=getPopObject%>  =  window.open("<%=getPath%><%=getPopName%>","<%=getPopObject%>","location=no,status=no,menubar=no,scrollbars=no,resizable=no,titlebar=no,width=400,height=300");
</script>
<%
					end if
				else
%>
<script>
var <%=getPopObject%>  =  window.open("<%=getPath%><%=getPopName%>","<%=getPopObject%>","location=no,status=no,menubar=no,scrollbars=no,resizable=no,titlebar=no,width=400,height=300");
</script>
<%			
				end if
		end if
		RecFlagpopupshow.MoveNext		
	 Wend
	 RecFlagpopupshow.Close()
	 set RecFlagpopupshow = Nothing
%>
<!------------------------------------------------------End Code for get flag show system  popup--------------------------------------------------------->
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF"> 
<center>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td colspan="3" bgcolor="#cd9efe" align="right"><img src="images/violet_head_web_2_1.jpg" width="1264"  /></td></tr>
<!--<tr>
<td width="20%" align="center" bgcolor="#ecdaff"><img src="images/head_logo.jpg" height="150" /></td>
  <td width="60%"  align="center"><img src="images/head1.jpg" alt="" height="150" /></td>
  <td width="20%" bgcolor="#8208ff">&nbsp;</td>
</tr>-->
<tr>
<td valign="top" bgcolor="#FFFFFF"><br />
<div align="center" style="font-size:16px; background-color:#8208ff; height:27px; vertical-align:middle; color:#FFFFFF"><b>เมนูหลัก : ระบบคุณภาพ</b></div>
<!--Left menu-->
<ul id="MenuBar1" class="MenuBarVertical">
  <li><a href="pdf/โครงสร้างการบริหารงานระบบคุณภาพ.pdf" target="contain">โครงสร้างและบทบาทหน้าที่</a></li>
  <li><a href="#" class="MenuBarItemSubmenu">คำสั่งผู้รับผิดชอบระบบคุณภาพ</a>
      <ul>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">คณะกรรมการบริหารระบบคุณภาพและคณะกรรมการประสานงานระบบคุณภาพ (QMC &amp; QMR)</a>
          <ul>
            <li><a href="pdf/คำสั่งQS/คณะ QMC&amp;QMR.pdf">1. คำสั่งที่ 127/2557</a></li>
            <li><a href="pdf/คณะ QMR_เพิ่มเติมเงินทุน.pdf" target="_blank">2. คำสั่งที่ 168/2557 (แก้ไขเพิ่มเติม)</a></li>
          </ul>
           </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">คณะทำงานระบบคุณภาพ (QS Team)</a>
          <ul>
            <li><a href="pdf/คำสั่งQS/คณะ QS กอง.pdf" target="_blank">1.คำสั่งที่    128/2557 </a></li>
            <li><a href="pdf/คณะ QS กอง_เพิ่มเติมเงินทุน.pdf" target="_blank">2.คำสั่งที่ 169/2557 (แก้ไขเพิ่มเติม)</a></li>
          </ul>
           </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">คณะทำงานประสานการจัดการเอกสารระบบคุณภาพ (DC Team)</a>
          <ul>
            <li><a href="pdf/คำสั่งQS/คณะ DC.pdf" target="_blank">1.คำสั่งที่ 129/2557</a></li>
            <li><a href="pdf/คณะ DC_เพิ่มเติมเงินทุน.pdf" target="_blank">2.คำสั่งที่  170/2557 (แก้ไขเพิ่มเติม)</a></li>
          </ul>
          </li>
        <li><a href="pdf/คำสั่งQS/จัดตั้งศูนย์คุณภาพ.pdf" target="_blank">จัดตั้งศูนย์คุณภาพ (Quality Center)</a> </li>
        <li><a href="pdf/คำสั่งแต่งตั้งคณะLead Auditor.pdf" target="_blank">หัวหน้าผู้ตรวจติดตามคุณภาพภายใน (Lead Auditor Team)</a></li>
        <li><a href="pdf/คณะผู้ตรวจติดตาม.pdf" target="_blank">คณะผู้ตรวจติดตามคุณภาพภายใน (Auditor Team)</a></li>
      </ul>
  </li>
  <li><a href="http://filing.fda.moph.go.th/kmfda/_block/qos/กรอบประกาศนียบัตร.pdf" target="_blank">นโยบายและวัตถุประสงค์คุณภาพ</a> </li>
  <li><a href="#" class="MenuBarItemSubmenu">ข้อกำหนดระบบคุณภาพ</a>
      <ul>
        <li><a href="pdf/หนังสือข้อกำหนดระบบคุณภาพ.pdf" target="_blank">ข้อกำหนดระบบคุณภาพของสำนักงานคณะกรรมการอาหารและยา : 2557</a></li>
        <li><a href="#" class="MenuBarItemSubmenu">มาตรฐานระบบคุณภาพ-ข้อกำหนดทั่วไปสำหรับสำนักงานคณะกรรมการอาหารและยา : 2552</a>
          <ul>
            <li><a href="http://filing.fda.moph.go.th/library5/fda_standard.pdf" target="_blank">PDF File</a></li>
            <li><a href="ppt/มาตรฐานระบบคุณภาพ ปี2552.ppt" target="_blank">PPT File</a></li>
          </ul>
          </li>
      </ul>
  </li>
  <li><a href="http://filing.fda.moph.go.th/library/e-file/TPD/ศูนย์วิทยบริการ/เอกสารระบบคุณภาพ/อย/Q/Q-FDA-T-1_1.pdf" target="_blank">คู่มือคุณภาพ</a></li>
  <li><a href="pdf/04_เป้าหมาย &amp; ตัวชี้วัด_57.pdf" target="_blank">เป้าหมายและตัวชี้วัด</a></li>
  <li><a href="#" class="MenuBarItemSubmenu">Road Map</a>
      <ul>
        <li><a href="pdf/03_QS Roadmap 2014-2016_Thai.pdf" target="_blank">Thai version</a></li>
        <li><a href="pdf/02_QS Roadmap 2014-2016_English.pdf" target="_blank">English version</a></li>
      </ul>
  </li>
  <li><a href="pdf/05_แผนปฏิบัติการ_57.pdf" target="_blank" class="MenuBarItemSubmenu">แผนการดำเนินงาน</a>
    <ul>
      <li><a href="pdf/แผนปฏิบัติการ_57.pdf" target="_blank">แผนปฏิบัติการประจำปี</a></li>
      <li><a href="pdf/แผนการตรวจติดตาม.pdf" target="_blank">แผนการตรวจติดตามคุณภาพภายใน</a></li>
      <li><a href="#">แผนการประชุมทบทวนโดยฝ่ายบริหาร</a></li>
    </ul>
    </li>
  <li><a href="#" class="MenuBarItemSubmenu">ผลการดำเนินงาน</a>
      <ul>
        <li><a href="#">ผลการดำเนินงานตามวัตถประสงค์คุณภาพ</a></li>
        <li><a href="#">ผลการดำเนินงานตามเป้าหมายและตัวชี้วัด</a></li>
        <li><a href="KPIReport.asp" target="_blank">ผลการวิเคราะห์ตาม KPI</a></li>
        <li><a href="ReportSOP.asp" target="_blank">ผลการวิเคราะห์กระบวนการ Core&amp;Support Process</a></li>
        <li><a href="ReportAnalaysis.asp" target="_blank">ผลการวิเคราะห์ความสอดคล้องและความต้องการ</a></li>
        <li><a href="ReviewReport.asp" target="_blank">ผลการดำเนินงานทบทวนเอกสารคุณภาพ</a> </li>
        <li><a href="AnalaysisInternalAuditReport.asp" target="_blank">ผลการดำเนินงานตรวจติดตามคุณภาพภายใน</a></li>
        <li><a href="ManagementReviewReport.asp" target="_self">ผลการดำเนินงานประชุมทบทวนโดยฝ่ายบริหาร</a></li>
      </ul>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu" title="<% if isEmpty(session("member")) = true then response.write "ต้อง Login ก่อนใช้งาน" else   response.write "ใช้งานได้เลย" end if %>" >ระบบรายงาน</a>
      <% if isEmpty(session("member")) = false then %>
      <ul>
        <li><a href="showProcedure.asp" target="_self">1. บันทึกและพิมพ์การเลือกกระบวนการ Core&amp;Support Process</a></li>
        <li><a href="analaysis.asp" target="_self">2. บันทึกและพิมพ์ความสอดคล้องและความต้องการ</a></li>
        <li><a href="FReview.asp" target="_self">3. บันทึกและพิมพ์การทบทวนเอกสารคุณภาพ</a></li>
        <li><a href="#" class="MenuBarItemSubmenu">4. บันทึกและพิมพ์การตรวจติดตามคุณภาพภายใน</a>
            <ul>
              <li><a href="InternalAudit.asp" target="_self">รายงานการตรวจติดตามคุณภาพภายใน</a></li>
              <li><a href="EditCAR.asp" target="_blank">การปฏิบัติการแก้ไข (CAR)</a></li>
              <li><a href="EditPAR.asp" target="_blank">การปฏิบัติการป้องกัน (PAR)</a></li>
              <li><a href="FollowUp.asp" target="_blank">การตรวจติดตามซ้ำ (Follow up)</a></li>
            </ul>
        </li>
        <li><a href="ManagementReview.asp" target="_self">5. บันทึกและพิมพ์การประชุมทบทวนโดยฝ่ายบริหาร</a></li>
      </ul>
      <% end if %>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu"  title="<% if isEmpty(session("member")) = true then response.write "ต้อง Login ก่อนใช้งาน" else   response.write "ใช้งานได้เลย" end if %>"> เอกสารระบบคุณภาพของ อย.</a>
      <% if isEmpty(session("member")) = False then %>
      <ul>
        <li><a href="http://filing.fda.moph.go.th/kmfda/default.asp?page=doc">สำนักงานคณะกรรมการอาหารและยา : คู่มือขั้นตอนการปฏิบัติงานกลาง</a> </li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=F596021&amp;bid=137&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@322,^,@323,^,@290,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%a7%d2%b9%e0%c5%a2%d2%b9%d8%a1%d2%c3%a1%c3%c1&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">สำนักงานเลขานุการกรม</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E30368&amp;bid=131&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">สำนักยา</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E31380&amp;bid=135&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@346,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%cd%d2%cb%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">สำนักอาหาร</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3239D&amp;bid=134&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@580,^,@581,^,@451,^,@582,^,@50,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%b4%e8%d2%b9%cd%d2%cb%d2%c3%e1%c5%d0%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">สำนักด่านอาหารและยา</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E323B6&amp;bid=132&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@48,^,@355,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a4%c7%ba%a4%d8%c1%e0%a4%c3%d7%e8%cd%a7%ca%d3%cd%d2%a7&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กลุ่มควบคุมเครื่องสำอาง</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E333D0&amp;bid=128&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@48,^,@67,^,@483,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a4%c7%ba%a4%d8%c1%c7%d1%b5%b6%d8%cd%d1%b9%b5%c3%d2%c2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กลุ่มควบคุมวัตถุอันตราย</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E343EA&amp;bid=130&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@48,^,@207,^,@208,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%a4%c7%ba%a4%d8%c1%e0%a4%c3%d7%e8%cd%a7%c1%d7%cd%e1%be%b7%c2%ec&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กองควบคุมเครื่องมือแพทย์</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E34405&amp;bid=129&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@48,^,@67,^,@234,^,@235,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%a4%c7%ba%a4%d8%c1%c7%d1%b5%b6%d8%e0%ca%be%b5%d4%b4&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กองควบคุมวัตถุเสพติด</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E35420&amp;bid=59&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@41,^,@279,^,@50,^,@312,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%e1%bc%b9%a7%d2%b9%e1%c5%d0%c7%d4%aa%d2%a1%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กองแผนงานและวิชาการ</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3643A&amp;bid=122&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@35,^,@273,^,@274,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%be%d1%b2%b9%d2%c8%d1%a1%c2%c0%d2%be%bc%d9%e9%ba%c3%d4%e2%c0%a4&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กองพัฒนาศักยภาพผู้บริโภค</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E032F7&amp;bid=142&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@278,^,@279,^,@280,^,@274,^,@281,^,@282,^,@283,^,@98,^,@284,^,@50,^,@285,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%ca%e8%a7%e0%ca%c3%d4%c1%a7%d2%b9%a4%d8%e9%c1%a4%c3%cd%a7%bc%d9%e9%ba%c3%d4%e2%c0%a4%b4%e9%d2%b9%bc%c5%d4%b5%c0%d1%b3%b1%ec%ca%d8%a2%c0%d2%be%e3%b9%ca%e8%c7%b9%c0%d9%c1%d4%c0%d2%a4%e1%c5%d0%b7%e9%cd%a7%b6%d4%e8%b9&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กองส่งเสริมงานคุ้มครองผู้บริโภคฯ</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E37455&amp;bid=141&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@352,^,@346,^,@50,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a1%ae%cb%c1%d2%c2%cd%d2%cb%d2%c3%e1%c5%d0%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กลุ่มกฎหมายอาหารและยา</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3846E&amp;bid=139&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@398,^,@292,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%b5%c3%c7%a8%ca%cd%ba%c0%d2%c2%e3%b9&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กลุ่มตรวจสอบภายใน</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E38488&amp;bid=140&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@35,^,@454,^,@562,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%be%d1%b2%b9%d2%c3%d0%ba%ba%ba%c3%d4%cb%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">กลุ่มพัฒนาระบบบริหาร</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E394A4&amp;bid=138&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@476,^,@429,^,@282,^,@283,^,@797,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%c8%d9%b9%c2%ec%ba%c3%d4%a1%d2%c3%bc%c5%d4%b5%c0%d1%b3%b1%ec%ca%d8%a2%c0%d2%be%e0%ba%e7%b4%e0%ca%c3%e7%a8&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">ศูนย์บริการผลิตภัณฑ์สุขภาพเบ็ดเสร็จ</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=F545FA2&amp;bid=136&amp;qst=@478&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%c8%d9%b9%c2%ec%a2%e9%cd%c1%d9%c5%e1%c5%d0%ca%d2%c3%ca%b9%e0%b7%c8&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">ศูนย์ข้อมูลและสารสนเทศ</a></li>
      </ul>
      <% end if %>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu">ประวัติการฝึกอบรม</a>
      <ul>
        <li><a href="#" class="MenuBarItemSubmenu">Introduction Auditor</a>
            <ul>
              <li><a href="http://filing.fda.moph.go.th/kmfda/_block/qos/pdf/History_Intro&amp;InternalAuditor.pdf" target="_blank">Introduction and Internal Auditor ISO 9001:2008 ประจำปีงบประมาณ 2557 </a></li>
              <li><a href="pdf/17065.pdf" target="_blank">Introduction to ISO/IEC 17065:2012 ประจำปีงบประมาณ 2557</a></li>
              <li><a href="pdf/17021.pdf" target="_blank">Introduction to ISO/IEC 17021:2012 ประจำปีงบประมาณ 2557</a></li>
            </ul>
        </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">Lead Auditor</a>
          <ul>
            <li><a href="pdf/ประวัติ Lead.pdf" target="_blank">ปีงบประมาณ พ.ศ.2557</a></li>
          </ul>
          </li>
      </ul>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu" title="<% if isEmpty(session("member")) = true then response.write "ต้อง Login ก่อนใช้งาน" else   response.write "ใช้งานได้เลย" end if %>" >รายงานการประชุม</a>
    <% if isEmpty(session("member")) = false then %>
    <ul>
      <li><a href="#" class="MenuBarItemSubmenu">คณะกรรมการบริหารระบบคุณภาพ</a>
        <ul>
          <li><a href="pdf/รายงานการประชุมคณะกรรมการบริหารระบบคุณภาพ_ครั้งที่ 1-57.pdf" target="_blank">ครั้งที่ 1/2557</a></li>
        </ul>
        </li>
      <li><a href="#" class="MenuBarItemSubmenu">คณะกรรมการประสานงานระบบคุณภาพ</a>
        <ul>
          <li><a href="pdf/รายงานการประชุมคณะกรรมการประสานสานงานระบบคุณภาพ _ครั้งที่ 1-2557.pdf" target="_blank">ครั้งที่ 1/2557</a></li>
          <li><a href="pdf/รายงานการประชุมคณะกรรมการประสานสานงานระบบคุณภาพ _ครั้งที่ 2-2557.pdf" target="_blank">ครั้งที่ 2/2557</a></li>
        </ul>
        </li>
    </ul>
    <% end if %>
    </li>
</ul>
<script type="text/javascript">
<!--
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", {imgRight:"SpryAssets/SpryMenuBarRightHover.gif"});
//-->
</script>
<!--Left menu-->

<br />
<!--------------------------------------------------Start Calendar---------------------------------------------->
<%
Function GetDaysInMonth(iMonth, iYear)
	Dim dTemp
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
	GetDaysInMonth = Day(dTemp)
End Function

' Previous implementation on GetDaysInMonth
'Function GetDaysInMonth(iMonth, iYear)
'	Select Case iMonth
'		Case 1, 3, 5, 7, 8, 10, 12
'			GetDaysInMonth = 31
'		Case 4, 6, 9, 11
'			GetDaysInMonth = 30
'		Case 2
'			If IsDate("February 29, " & iYear) Then
'				GetDaysInMonth = 29
'			Else
'				GetDaysInMonth = 28
'			End If
'	End Select
'End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
	Dim dTemp
	dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
	SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
	AddOneMonth = DateAdd("m", 1, dDate)
End Function
' ***End Function Declaration***


Dim dDate     ' Date we're displaying calendar for
Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table


' Get selected date.  There are two ways to do this.
' First check if we were passed a full date in RQS("date").
' If so use it, if not look for seperate variables, putting them togeter into a date.
' Lastly check if the date is valid...if not use today
If IsDate(Request.QueryString("date")) Then
	dDate = CDate(Request.QueryString("date"))
Else
	If IsDate(Request.QueryString("month") & "/" & Request.QueryString("day") & "/" & Request.QueryString("year")) Then
		dDate = CDate(Request.QueryString("month") & "/" & Request.QueryString("day") & "/" & Request.QueryString("year"))
	Else
		dDate = Date()
		' The annoyingly bad solution for those of you running IIS3
		If Len(Request.QueryString("month")) <> 0 Or Len(Request.QueryString("day")) <> 0 Or Len(Request.QueryString("year")) <> 0 Or Len(Request.QueryString("date")) <> 0 Then
			Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
		End If
		' The elegant solution for those of you running IIS4
		'If Request.QueryString.Count <> 0 Then Response.Write "The date you picked was not a valid date.  The calendar was set to today's date.<BR><BR>"
	End If
End If
'Now we've got the date.  Now get Days in the choosen month and the day of the week it starts on.
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
iDOW = GetWeekdayMonthStartsOn(dDate)
%>
<table BORDER=0 CELLSPACING=2 CELLPADDING=2 align="center" width="80%"><tr><td align="center"><b>ปฏิทินกิจกรรมระบบคุณภาพ</b></td></tr></table>
<TABLE BORDER=10 CELLSPACING=0 CELLPADDING=0 align="center" width="80%">
<TR>
<TD>
<TABLE BORDER=1 CELLSPACING=0 CELLPADDING=1 BGCOLOR=#facdf2 width="100%">
	<TR>
		<TD BGCOLOR=#000099 ALIGN="center" COLSPAN=7>
			<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0 bgcolor="#9445d4" >
				<TR>
					<TD ALIGN="right"><A HREF="./calendaractivity.asp?date=<%= SubtractOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="-1">&lt;&lt;</FONT></A></TD>
					<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1"><%= MonthName(Month(dDate)) & "  " & (Year(dDate)+543) %></FONT></TD>
					<TD ALIGN="left"><A HREF="./calendaractivity.asp?date=<%= AddOneMonth(dDate) %>"><FONT COLOR=#FFFF00 SIZE="-1">&gt;&gt;</FONT></A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR bgcolor="#b27ee0">
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">อา</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">จ</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">อ</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">พ</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">พฤ</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">ศ</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">ส</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
	</TR>
<%
' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
If iDOW <> 1 Then
	Response.Write vbTab & "<TR>" & vbCrLf
	iPosition = 1
	Do While iPosition < iDOW
		Response.Write vbTab & vbTab & "<TD>&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
End If

' Write days of month in proper day slots
iCurrent = 1
iPosition = iDOW
Do While iCurrent <= iDIM
	' If we're at the begginning of a row then write TR
	If iPosition = 1 Then
		Response.Write vbTab & "<TR>" & vbCrLf
	End If
	'--------------------------Code for get data from DB-------------------------------
	GDate = Month(dDate) & "/" &iCurrent  & "/" &Year(dDate)
	get_DataAc = getDataCalendarActivity(GDate)
	'get_DataBook = getDataCalendarBooking(GDate)
	if get_DataAc > 0 then
		setColor="yellow"
	else
		setColor="#facdf2"
	end if 
	'---------------------------------------------------------------------------------------
	' If the day we're writing is the selected day then highlight it somehow.
	If iCurrent = Day(dDate) Then
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD BGCOLOR="""&setColor&"""><A HREF=""./calendaractivity.asp?date=" &iCurrent  & "/" &  Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""+1"" color=""red""><b>" & iCurrent & "</b></FONT></A><BR></TD>" & vbCrLf
	Else
		'Response.Write vbTab & vbTab & "<TD><A HREF=""./calendar.asp?date=" & Month(dDate) & "-" & iCurrent & "-" & (Year(dDate)+543) & """><FONT SIZE=""-3"">" & iCurrent & "</FONT></A><BR><BR></TD>" & vbCrLf
		Response.Write vbTab & vbTab & "<TD BGCOLOR="""&setColor&""" ><A HREF=""./calendaractivity.asp?date=" &iCurrent  & "/" & Month(dDate) & "/" & (Year(dDate)+543) & """><FONT SIZE=""-1""  style=""text-decoration:none; color:#000000"">" & iCurrent & "</FONT></A><BR></TD>" & vbCrLf
	End If
	
	' If we're at the endof a row then write /TR
	If iPosition = 7 Then
		Response.Write vbTab & "</TR>" & vbCrLf
		iPosition = 0
	End If
	
	' Increment variables
	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
Loop

' Write spacer cells at end of last row if month doesn't end on a Saturday.
If iPosition <> 1 Then
	Do While iPosition <= 7
		Response.Write vbTab & vbTab & "<TD bgcolor=""#facdf2"">&nbsp;</TD>" & vbCrLf
		iPosition = iPosition + 1
	Loop
	Response.Write vbTab & "</TR>" & vbCrLf
End If
%>
</TABLE>
</TD>
</TR>
</TABLE>
<!--------------------------------------------------End Calendar------------------------------------------------>
<br />
<!--Left menu number2-->
<div align="center" style="font-size:16px; background-color:#8208ff; height:23px; vertical-align:middle; color:#FFFFFF"><b>เมนูหลัก : PMQA</b></div>
<!--Left menu-->
<ul id="MenuBar2" class="MenuBarVertical">
  <li><a href="#">ลักษณะสำคัญองค์กร</a></li>
  <li><a href="#">หมวด 1 การนำองค์กร</a> </li>
  <li><a href="#" target="_blank">หมวด 2 การวางแผนเชิงยุทธศาสตร์</a></li>
  <li><a href="#">หมวด 3 การให้ความสำคัญกับผู้รับบริการและผู้มีส่วนได้เสีย</a>    </li>
  <li><a href="#" target="_blank">หมวด 4 การวัด การวิเคราะห์ และการจัดการความรู้</a></li>
  <li><a href="#">หมวด 5 การมุ่งเน้นทรัพยากรบุคคล</a>      </li>
  <li><a href="#" class="MenuBarItemSubmenu">หมวด 6 การจัดการกระบวนการ</a>
    <ul>
      <li><a href="pdf/แผนพัฒนาองค์การหมวด6_2557.pdf" target="_blank">แผนพัฒนาองค์กรหมวด 6</a></li>
      <li><a href="#">กระบวนการกำหนดกระบวนสร้างคุณค่าและกระบวนการสนับสนุน (PM1, PM7)</a></li>
      <li><a href="#">กระบวนการจัดทำข้อกำหนดและตัวชี้วัดที่สำคัญของกระบวนการสร้างคุณค่าและกระบวนการสนับสนุน (PM2, PM4,PM8, PM10)</a></li>
      <li><a href="#">กระบวนการออกแบบกระบวนการ (PM3, PM9)</a></li>
      <li><a href="#">กระบวนการจัดทำมาตรฐานการปฏิบัติงานของกระบวนการสร้างคุณค่าและกระบวนการสนับสนุน(PM5, PM11)</a></li>
      <li><a href="#">กระบวนการปรับปรุงกระบวนการสร้างคุณค่าและกระบวนการสนับสนุน (PM6, PM12)</a></li>
    </ul>
          </li>
  <li><a href="#">หมวด 7 ผลลัพธ์การดำเนินการ</a></li>
  <li><a href="#">คำสั่งต่างๆ</a> </li>
</ul>
<script type="text/javascript">
<!--
var MenuBar2 = new Spry.Widget.MenuBar("MenuBar2", {imgRight:"SpryAssets/SpryMenuBarRightHover.gif"});
//-->
</script>
<!--Left menu-->
<!--Left menu number2-->

</td>
<td align="center" valign="top">
    <table width="780" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
  <tr><td>&nbsp;</td></tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!--<tr>
              <td><img src="images/qs_fda.gif" width="780" height="150"></td>
        </tr>-->
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
				<%
						if Isempty(session("member")) = True  then  	
				%>
				
				<!--<tr><td  bgcolor="#a27dff" >&nbsp;&nbsp;&nbsp;<span style="color: #FFFFFF"><font size="2" face="Ms Sans Serif"><strong>Login</strong></font></span>&nbsp;&nbsp;&nbsp;&nbsp;<span style="color: #FFFFFF"><font size="2" face="Ms Sans Serif"><strong>ผู้ที่สนใจต้องการเข้าดูเอกสารระบบคุณภาพของ อย. สามารถล๊อกอินได้ที่นี่</strong></font></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				</tr>-->
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="1">
				  <input type="hidden" name="qs_chk" value="QS">
                  <tr>
                    <td bgcolor="#f2e3ff" ><span style="color: #000000"><font size="2" face="Ms Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E-Mail&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;
                          <input name="Email" type="text" size=20 value="yourname@fda.moph.go.th" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff">
&nbsp;รหัสผ่าน :&nbsp;&nbsp;&nbsp;
                          <input name="Password" type="password" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff"  size=20>
                    </font></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit"  value="เข้าสู่ระบบ">
&nbsp;&nbsp;&nbsp;<input type="button" value="สมัครสมาชิก"onclick="javascript:{window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos/register.asp';}" style="cursor:pointer; cursor:hand" ></td>
                  </tr>
                  </form>
				  <!--<tr><td>&nbsp;sdfsdf</td></tr>-->
				  <% else%>
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="0">
				  <input type="hidden" name="qs_chk" value="QS">
                  
                  <tr>
                    <td bgcolor="#a27dff" align="right"  height="27"><font size="2" face="Ms Sans Serif" color="#3300ff">
                      <input type="submit"  value="ออกจากระบบ">&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                  </form>
				  
				  <% end if %>
                  <tr>
                    <td><!--#include file="checksession.asp"--></td>
                  </tr>
                </table></td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr><td colspan="2" align="center" valign="top">&nbsp;<!--<table align="center" border="0" cellpadding="0" cellspacing="0" width="98%"><tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <font size="2" face="Ms Sans Serif" color="#3300ff">"เฉพาะข้าราชการ อย. ที่ยังไม่มีรหัสผ่านในการเข้าดูเอกสารระบบคุณภาพขอให้แจ้งชื่อ นามสกุล มาที่ library@fda.moph.go.th และท่านจะได้รับการแจ้งรหัสผ่านกลับทางอีเมล์ที่แจ้งไว้"</font>
  <!--<font size="2" face="Ms Sans Serif" color="#3300ff">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*** ???????????????????????????????????????? ??. ????????????????????????? <a href="http://elib.fda.moph.go.th/kmfda/default.asp" target="_blank">FDA KM</a> ?????? "?????????????????????????" ???????????????????????????? e-mail address ?? ??. ??? @fda.moph.go.th ???????? ???????????????????? ????????????????????????????????????? (Password) .??????????????????????  ...*** </font></td>
        </tr></table>-->
  </td></tr>
        <tr>
              <td  height="50" valign="bottom">
			  <!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
<%

'set rs_counter=server.CreateObject("ADODB.recordset")
'rs_counter.open "Select * From tabUniIP ",con,1,3


'if session("check")="" then
'session("check")="x"
'rs_counter.addnew
'rs_counter("ip")=request.ServerVariables("REMOTE_HOST")
'rs_counter("date")=date
'rs_counter.update
'end if
'rs_counter.close

%>
                      <a href="javascript:openWin('counter.asp','counter',500,500)" >Uni IP <img src="../../_images/folder_stats.gif" width="30" height="30" border=0></a></td>
  </tr>
</table>-->

			  </td>
        </tr>
      </table></td>
  </tr>
</table>
</td>
<td valign="top">
<!--Right menu-->
<br />
<div align="center" style="font-size:16px; background-color:#8208ff; height:27px;color:#FFFFFF; vertical-align:middle">ความรู้ทั่วไป</div>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
  <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" >
          <table width="90%" border="0" align="left" cellpadding="3" cellspacing="0" class="text">
          <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%">&nbsp;</td></tr>
          <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=101" target="_self">เอกสารระบบคุณภาพ (Quality System Documentation)</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=102" target="_self">ลักษณะประโยชน์ของเอกสารระบบคุณภาพ</a></td>
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

<!--Right menu-->
</td>
</tr>
<tr><td  colspan="3" bgcolor="#d8b4fe" height="30px" align="center">
<strong><img src="http://elib.fda.moph.go.th/library/_images/icon_register.gif" width="31" height="21" align="absmiddle" /><font size="1">ผู้ใช้ขณะนี้ <%
numcount=Application("OnlineUser")
'numcount=right("0000000000000"&numcount,7)		
l=len(numcount)
for i=1 to l
num=mid(numcount,i,1)
display2=display2&num '"<img src=_images/counter/"&num&".gif align='absmiddle'>"
next
response.Write("<font color=000000><b>"&display2&"</b></font>")%></font></strong>
&nbsp;&nbsp;&nbsp;
<font size="2" face="Ms Sans Serif" color="#3300ff">&quot;เฉพาะข้าราชการ อย. ที่ยังไม่มีรหัสผ่านในการเข้าดูเอกสารระบบคุณภาพขอให้แจ้งชื่อ นามสกุล มาที่ library@fda.moph.go.th , qsfda@fda.moph.go.th และท่านจะได้รับการแจ้งรหัสผ่านกลับทางอีเมล์ที่แจ้งไว้&quot;</font>
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
</table>
</center>
</body>
</html>
