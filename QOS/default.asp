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
<title>�к��س�Ҿ�ͧ ��.</title>

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
<div align="center" style="font-size:16px; background-color:#8208ff; height:27px; vertical-align:middle; color:#FFFFFF"><b>������ѡ : �к��س�Ҿ</b></div>
<!--Left menu-->
<ul id="MenuBar1" class="MenuBarVertical">
  <li><a href="pdf/�ç���ҧ��ú����çҹ�к��س�Ҿ.pdf" target="contain">�ç���ҧ��к��ҷ˹�ҷ��</a></li>
  <li><a href="#" class="MenuBarItemSubmenu">����觼���Ѻ�Դ�ͺ�к��س�Ҿ</a>
      <ul>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">��С�����ú������к��س�Ҿ��Ф�С�����û���ҹ�ҹ�к��س�Ҿ (QMC &amp; QMR)</a>
          <ul>
            <li><a href="pdf/�����QS/��� QMC&amp;QMR.pdf">1. ����觷�� 127/2557</a></li>
            <li><a href="pdf/��� QMR_��������Թ�ع.pdf" target="_blank">2. ����觷�� 168/2557 (����������)</a></li>
          </ul>
           </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">��зӧҹ�к��س�Ҿ (QS Team)</a>
          <ul>
            <li><a href="pdf/�����QS/��� QS �ͧ.pdf" target="_blank">1.����觷��    128/2557 </a></li>
            <li><a href="pdf/��� QS �ͧ_��������Թ�ع.pdf" target="_blank">2.����觷�� 169/2557 (����������)</a></li>
          </ul>
           </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">��зӧҹ����ҹ��èѴ����͡����к��س�Ҿ (DC Team)</a>
          <ul>
            <li><a href="pdf/�����QS/��� DC.pdf" target="_blank">1.����觷�� 129/2557</a></li>
            <li><a href="pdf/��� DC_��������Թ�ع.pdf" target="_blank">2.����觷��  170/2557 (����������)</a></li>
          </ul>
          </li>
        <li><a href="pdf/�����QS/�Ѵ����ٹ��س�Ҿ.pdf" target="_blank">�Ѵ����ٹ��س�Ҿ (Quality Center)</a> </li>
        <li><a href="pdf/������觵�駤��Lead Auditor.pdf" target="_blank">���˹�Ҽ���Ǩ�Դ����س�Ҿ���� (Lead Auditor Team)</a></li>
        <li><a href="pdf/��м���Ǩ�Դ���.pdf" target="_blank">��м���Ǩ�Դ����س�Ҿ���� (Auditor Team)</a></li>
      </ul>
  </li>
  <li><a href="http://filing.fda.moph.go.th/kmfda/_block/qos/��ͺ��С�ȹ�ºѵ�.pdf" target="_blank">��º������ѵ�ػ��ʧ��س�Ҿ</a> </li>
  <li><a href="#" class="MenuBarItemSubmenu">��͡�˹��к��س�Ҿ</a>
      <ul>
        <li><a href="pdf/˹ѧ��͢�͡�˹��к��س�Ҿ.pdf" target="_blank">��͡�˹��к��س�Ҿ�ͧ�ӹѡ�ҹ��С���������������� : 2557</a></li>
        <li><a href="#" class="MenuBarItemSubmenu">�ҵðҹ�к��س�Ҿ-��͡�˹����������Ѻ�ӹѡ�ҹ��С���������������� : 2552</a>
          <ul>
            <li><a href="http://filing.fda.moph.go.th/library5/fda_standard.pdf" target="_blank">PDF File</a></li>
            <li><a href="ppt/�ҵðҹ�к��س�Ҿ ��2552.ppt" target="_blank">PPT File</a></li>
          </ul>
          </li>
      </ul>
  </li>
  <li><a href="http://filing.fda.moph.go.th/library/e-file/TPD/�ٹ���Էº�ԡ��/�͡����к��س�Ҿ/��/Q/Q-FDA-T-1_1.pdf" target="_blank">�����ͤس�Ҿ</a></li>
  <li><a href="pdf/04_������� &amp; ��Ǫ���Ѵ_57.pdf" target="_blank">���������е�Ǫ���Ѵ</a></li>
  <li><a href="#" class="MenuBarItemSubmenu">Road Map</a>
      <ul>
        <li><a href="pdf/03_QS Roadmap 2014-2016_Thai.pdf" target="_blank">Thai version</a></li>
        <li><a href="pdf/02_QS Roadmap 2014-2016_English.pdf" target="_blank">English version</a></li>
      </ul>
  </li>
  <li><a href="pdf/05_Ἱ��Ժѵԡ��_57.pdf" target="_blank" class="MenuBarItemSubmenu">Ἱ��ô��Թ�ҹ</a>
    <ul>
      <li><a href="pdf/Ἱ��Ժѵԡ��_57.pdf" target="_blank">Ἱ��Ժѵԡ�û�Шӻ�</a></li>
      <li><a href="pdf/Ἱ��õ�Ǩ�Դ���.pdf" target="_blank">Ἱ��õ�Ǩ�Դ����س�Ҿ����</a></li>
      <li><a href="#">Ἱ��û�Ъ�����ǹ�½��º�����</a></li>
    </ul>
    </li>
  <li><a href="#" class="MenuBarItemSubmenu">�š�ô��Թ�ҹ</a>
      <ul>
        <li><a href="#">�š�ô��Թ�ҹ����ѵ����ʧ��س�Ҿ</a></li>
        <li><a href="#">�š�ô��Թ�ҹ������������е�Ǫ���Ѵ</a></li>
        <li><a href="KPIReport.asp" target="_blank">�š������������ KPI</a></li>
        <li><a href="ReportSOP.asp" target="_blank">�š�����������кǹ��� Core&amp;Support Process</a></li>
        <li><a href="ReportAnalaysis.asp" target="_blank">�š��������������ʹ���ͧ��Ф�����ͧ���</a></li>
        <li><a href="ReviewReport.asp" target="_blank">�š�ô��Թ�ҹ���ǹ�͡��äس�Ҿ</a> </li>
        <li><a href="AnalaysisInternalAuditReport.asp" target="_blank">�š�ô��Թ�ҹ��Ǩ�Դ����س�Ҿ����</a></li>
        <li><a href="ManagementReviewReport.asp" target="_self">�š�ô��Թ�ҹ��Ъ�����ǹ�½��º�����</a></li>
      </ul>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu" title="<% if isEmpty(session("member")) = true then response.write "��ͧ Login ��͹��ҹ" else   response.write "��ҹ�����" end if %>" >�к���§ҹ</a>
      <% if isEmpty(session("member")) = false then %>
      <ul>
        <li><a href="showProcedure.asp" target="_self">1. �ѹ�֡��о���������͡��кǹ��� Core&amp;Support Process</a></li>
        <li><a href="analaysis.asp" target="_self">2. �ѹ�֡��о��������ʹ���ͧ��Ф�����ͧ���</a></li>
        <li><a href="FReview.asp" target="_self">3. �ѹ�֡��о�����÷��ǹ�͡��äس�Ҿ</a></li>
        <li><a href="#" class="MenuBarItemSubmenu">4. �ѹ�֡��о�����õ�Ǩ�Դ����س�Ҿ����</a>
            <ul>
              <li><a href="InternalAudit.asp" target="_self">��§ҹ��õ�Ǩ�Դ����س�Ҿ����</a></li>
              <li><a href="EditCAR.asp" target="_blank">��û�Ժѵԡ����� (CAR)</a></li>
              <li><a href="EditPAR.asp" target="_blank">��û�Ժѵԡ�û�ͧ�ѹ (PAR)</a></li>
              <li><a href="FollowUp.asp" target="_blank">��õ�Ǩ�Դ������ (Follow up)</a></li>
            </ul>
        </li>
        <li><a href="ManagementReview.asp" target="_self">5. �ѹ�֡��о�����û�Ъ�����ǹ�½��º�����</a></li>
      </ul>
      <% end if %>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu"  title="<% if isEmpty(session("member")) = true then response.write "��ͧ Login ��͹��ҹ" else   response.write "��ҹ�����" end if %>"> �͡����к��س�Ҿ�ͧ ��.</a>
      <% if isEmpty(session("member")) = False then %>
      <ul>
        <li><a href="http://filing.fda.moph.go.th/kmfda/default.asp?page=doc">�ӹѡ�ҹ��С���������������� : �����͢�鹵͹��û�Ժѵԧҹ��ҧ</a> </li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=F596021&amp;bid=137&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@322,^,@323,^,@290,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%a7%d2%b9%e0%c5%a2%d2%b9%d8%a1%d2%c3%a1%c3%c1&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ӹѡ�ҹ�Ţҹء�á��</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E30368&amp;bid=131&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ӹѡ��</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E31380&amp;bid=135&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@346,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%cd%d2%cb%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ӹѡ�����</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3239D&amp;bid=134&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@841,^,@580,^,@581,^,@451,^,@582,^,@50,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%ca%d3%b9%d1%a1%b4%e8%d2%b9%cd%d2%cb%d2%c3%e1%c5%d0%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ӹѡ��ҹ����������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E323B6&amp;bid=132&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@48,^,@355,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a4%c7%ba%a4%d8%c1%e0%a4%c3%d7%e8%cd%a7%ca%d3%cd%d2%a7&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">������Ǻ�������ͧ���ҧ</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E333D0&amp;bid=128&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@48,^,@67,^,@483,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a4%c7%ba%a4%d8%c1%c7%d1%b5%b6%d8%cd%d1%b9%b5%c3%d2%c2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">������Ǻ����ѵ���ѹ����</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E343EA&amp;bid=130&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@48,^,@207,^,@208,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%a4%c7%ba%a4%d8%c1%e0%a4%c3%d7%e8%cd%a7%c1%d7%cd%e1%be%b7%c2%ec&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ͧ�Ǻ�������ͧ���ᾷ��</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E34405&amp;bid=129&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@48,^,@67,^,@234,^,@235,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%a4%c7%ba%a4%d8%c1%c7%d1%b5%b6%d8%e0%ca%be%b5%d4%b4&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ͧ�Ǻ����ѵ���ʾ�Դ</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E35420&amp;bid=59&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@41,^,@279,^,@50,^,@312,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%e1%bc%b9%a7%d2%b9%e1%c5%d0%c7%d4%aa%d2%a1%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ͧἹ�ҹ����Ԫҡ��</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3643A&amp;bid=122&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@35,^,@273,^,@274,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%be%d1%b2%b9%d2%c8%d1%a1%c2%c0%d2%be%bc%d9%e9%ba%c3%d4%e2%c0%a4&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ͧ�Ѳ���ѡ��Ҿ��������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E032F7&amp;bid=142&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@1,^,@278,^,@279,^,@280,^,@274,^,@281,^,@282,^,@283,^,@98,^,@284,^,@50,^,@285,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%cd%a7%ca%e8%a7%e0%ca%c3%d4%c1%a7%d2%b9%a4%d8%e9%c1%a4%c3%cd%a7%bc%d9%e9%ba%c3%d4%e2%c0%a4%b4%e9%d2%b9%bc%c5%d4%b5%c0%d1%b3%b1%ec%ca%d8%a2%c0%d2%be%e3%b9%ca%e8%c7%b9%c0%d9%c1%d4%c0%d2%a4%e1%c5%d0%b7%e9%cd%a7%b6%d4%e8%b9&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ͧ��������ҹ������ͧ���������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E37455&amp;bid=141&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@352,^,@346,^,@50,^,@32,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%a1%ae%cb%c1%d2%c2%cd%d2%cb%d2%c3%e1%c5%d0%c2%d2&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">���������������������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E3846E&amp;bid=139&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@398,^,@292,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%b5%c3%c7%a8%ca%cd%ba%c0%d2%c2%e3%b9&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�������Ǩ�ͺ����</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E38488&amp;bid=140&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@351,^,@35,^,@454,^,@562,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%a1%c5%d8%e8%c1%be%d1%b2%b9%d2%c3%d0%ba%ba%ba%c3%d4%cb%d2%c3&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">������Ѳ���к�������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=E394A4&amp;bid=138&amp;qst=@10,@454,^,@49,^,@211,^,@572,^,@56,^,@182,^,@740,^,@476,^,@429,^,@282,^,@283,^,@797,^&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%c8%d9%b9%c2%ec%ba%c3%d4%a1%d2%c3%bc%c5%d4%b5%c0%d1%b3%b1%ec%ca%d8%a2%c0%d2%be%e0%ba%e7%b4%e0%ca%c3%e7%a8&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ٹ���ԡ�ü�Ե�ѳ���آ�Ҿ������</a></li>
        <li><a href="http://filing.fda.moph.go.th/elib/cgi-bin/opacexe.exe?op=dsp&amp;wa=F545FA2&amp;bid=136&amp;qst=@478&amp;lang=1&amp;db=Efileall&amp;pat=%e0%cd%a1%ca%d2%c3%c3%d0%ba%ba%a4%d8%b3%c0%d2%be%b5%d2%c1%a2%e9%cd%a1%d3%cb%b9%b4%b7%d1%e8%c7%e4%bb+%a2%cd%a7+%cd%c2.+:+%c8%d9%b9%c2%ec%a2%e9%cd%c1%d9%c5%e1%c5%d0%ca%d2%c3%ca%b9%e0%b7%c8&amp;cat=gen&amp;skin=u&amp;lpp=20&amp;catop=&amp;scid=zzz" target="_blank">�ٹ�������������ʹ��</a></li>
      </ul>
      <% end if %>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu">����ѵԡ�ý֡ͺ��</a>
      <ul>
        <li><a href="#" class="MenuBarItemSubmenu">Introduction Auditor</a>
            <ul>
              <li><a href="http://filing.fda.moph.go.th/kmfda/_block/qos/pdf/History_Intro&amp;InternalAuditor.pdf" target="_blank">Introduction and Internal Auditor ISO 9001:2008 ��Шӻէ�����ҳ 2557 </a></li>
              <li><a href="pdf/17065.pdf" target="_blank">Introduction to ISO/IEC 17065:2012 ��Шӻէ�����ҳ 2557</a></li>
              <li><a href="pdf/17021.pdf" target="_blank">Introduction to ISO/IEC 17021:2012 ��Шӻէ�����ҳ 2557</a></li>
            </ul>
        </li>
        <li><a href="#" target="_blank" class="MenuBarItemSubmenu">Lead Auditor</a>
          <ul>
            <li><a href="pdf/����ѵ� Lead.pdf" target="_blank">�է�����ҳ �.�.2557</a></li>
          </ul>
          </li>
      </ul>
  </li>
  <li><a href="#" class="MenuBarItemSubmenu" title="<% if isEmpty(session("member")) = true then response.write "��ͧ Login ��͹��ҹ" else   response.write "��ҹ�����" end if %>" >��§ҹ��û�Ъ��</a>
    <% if isEmpty(session("member")) = false then %>
    <ul>
      <li><a href="#" class="MenuBarItemSubmenu">��С�����ú������к��س�Ҿ</a>
        <ul>
          <li><a href="pdf/��§ҹ��û�Ъ����С�����ú������к��س�Ҿ_���駷�� 1-57.pdf" target="_blank">���駷�� 1/2557</a></li>
        </ul>
        </li>
      <li><a href="#" class="MenuBarItemSubmenu">��С�����û���ҹ�ҹ�к��س�Ҿ</a>
        <ul>
          <li><a href="pdf/��§ҹ��û�Ъ����С�����û���ҹ�ҹ�ҹ�к��س�Ҿ _���駷�� 1-2557.pdf" target="_blank">���駷�� 1/2557</a></li>
          <li><a href="pdf/��§ҹ��û�Ъ����С�����û���ҹ�ҹ�ҹ�к��س�Ҿ _���駷�� 2-2557.pdf" target="_blank">���駷�� 2/2557</a></li>
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
<table BORDER=0 CELLSPACING=2 CELLPADDING=2 align="center" width="80%"><tr><td align="center"><b>��ԷԹ�Ԩ�����к��س�Ҿ</b></td></tr></table>
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
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">��</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">�</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">�</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">�</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">��</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">�</B></FONT><BR>
		  <IMG SRC="./images/spacer.gif" WIDTH=20 HEIGHT=1 BORDER=0></TD>
		<TD ALIGN="center"><FONT COLOR=#FFFF00 size="-1">�</B></FONT><BR>
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
<div align="center" style="font-size:16px; background-color:#8208ff; height:23px; vertical-align:middle; color:#FFFFFF"><b>������ѡ : PMQA</b></div>
<!--Left menu-->
<ul id="MenuBar2" class="MenuBarVertical">
  <li><a href="#">�ѡɳ��Ӥѭͧ���</a></li>
  <li><a href="#">��Ǵ 1 ��ù�ͧ���</a> </li>
  <li><a href="#" target="_blank">��Ǵ 2 ����ҧἹ�ԧ�ط���ʵ��</a></li>
  <li><a href="#">��Ǵ 3 ����������Ӥѭ�Ѻ����Ѻ��ԡ����м������ǹ������</a>    </li>
  <li><a href="#" target="_blank">��Ǵ 4 ����Ѵ ����������� ��С�èѴ��ä������</a></li>
  <li><a href="#">��Ǵ 5 �������鹷�Ѿ�ҡúؤ��</a>      </li>
  <li><a href="#" class="MenuBarItemSubmenu">��Ǵ 6 ��èѴ��á�кǹ���</a>
    <ul>
      <li><a href="pdf/Ἱ�Ѳ��ͧ������Ǵ6_2557.pdf" target="_blank">Ἱ�Ѳ��ͧ�����Ǵ 6</a></li>
      <li><a href="#">��кǹ��á�˹���кǹ���ҧ�س�����С�кǹ���ʹѺʹع (PM1, PM7)</a></li>
      <li><a href="#">��кǹ��èѴ�Ӣ�͡�˹���е�Ǫ���Ѵ����Ӥѭ�ͧ��кǹ������ҧ�س�����С�кǹ���ʹѺʹع (PM2, PM4,PM8, PM10)</a></li>
      <li><a href="#">��кǹ����͡Ẻ��кǹ��� (PM3, PM9)</a></li>
      <li><a href="#">��кǹ��èѴ���ҵðҹ��û�Ժѵԧҹ�ͧ��кǹ������ҧ�س�����С�кǹ���ʹѺʹع(PM5, PM11)</a></li>
      <li><a href="#">��кǹ��û�Ѻ��ا��кǹ������ҧ�س�����С�кǹ���ʹѺʹع (PM6, PM12)</a></li>
    </ul>
          </li>
  <li><a href="#">��Ǵ 7 ���Ѿ���ô��Թ���</a></li>
  <li><a href="#">����觵�ҧ�</a> </li>
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
				
				<!--<tr><td  bgcolor="#a27dff" >&nbsp;&nbsp;&nbsp;<span style="color: #FFFFFF"><font size="2" face="Ms Sans Serif"><strong>Login</strong></font></span>&nbsp;&nbsp;&nbsp;&nbsp;<span style="color: #FFFFFF"><font size="2" face="Ms Sans Serif"><strong>�����ʹ㨵�ͧ�����Ҵ��͡����к��س�Ҿ�ͧ ��. ����ö��͡�Թ������</strong></font></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				</tr>-->
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="1">
				  <input type="hidden" name="qs_chk" value="QS">
                  <tr>
                    <td bgcolor="#f2e3ff" ><span style="color: #000000"><font size="2" face="Ms Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E-Mail&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;
                          <input name="Email" type="text" size=20 value="yourname@fda.moph.go.th" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff">
&nbsp;���ʼ�ҹ :&nbsp;&nbsp;&nbsp;
                          <input name="Password" type="password" style=" font-size:14px; font-family:'Ms Sans Serif',Georgia, 'Times New Roman', Times, serif; color:#3300ff"  size=20>
                    </font></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit"  value="�������к�">
&nbsp;&nbsp;&nbsp;<input type="button" value="��Ѥ���Ҫԡ"onclick="javascript:{window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos/register.asp';}" style="cursor:pointer; cursor:hand" ></td>
                  </tr>
                  </form>
				  <!--<tr><td>&nbsp;sdfsdf</td></tr>-->
				  <% else%>
                  <form name="formlogin" method="post" action="default.asp">
				  <input type="hidden" name="login" value="0">
				  <input type="hidden" name="qs_chk" value="QS">
                  
                  <tr>
                    <td bgcolor="#a27dff" align="right"  height="27"><font size="2" face="Ms Sans Serif" color="#3300ff">
                      <input type="submit"  value="�͡�ҡ�к�">&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
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
		  <font size="2" face="Ms Sans Serif" color="#3300ff">"੾�Т���Ҫ��� ��. ����ѧ��������ʼ�ҹ㹡����Ҵ��͡����к��س�Ҿ������駪��� ���ʡ�� �ҷ�� library@fda.moph.go.th ��з�ҹ�����Ѻ��������ʼ�ҹ��Ѻ�ҧ�������������"</font>
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
<div align="center" style="font-size:16px; background-color:#8208ff; height:27px;color:#FFFFFF; vertical-align:middle">�����������</div>
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
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=101" target="_self">�͡����к��س�Ҿ (Quality System Documentation)</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=102" target="_self">�ѡɳл���ª��ͧ�͡����к��س�Ҿ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=104" target="_self">��鹵͹��èѴ���͡����к��س�Ҿ</a></td>
          </tr>
		  <tr>
          <td width="2%"><img src="../../_images/arrowL2.gif" width="15" height="13"></td>
          <td width="98%"><a href="http://filing.fda.moph.go.th/kmfda/_block/QOS/default.asp?page=data_detail&ID_L3=108" target="_self">��äǺ����͡�����Т����� (Document and Data Control)</a></td>
          </tr>
          </table>
        </td></tr></table>

<!--Right menu-->
</td>
</tr>
<tr><td  colspan="3" bgcolor="#d8b4fe" height="30px" align="center">
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
</table>
</center>
</body>
</html>
