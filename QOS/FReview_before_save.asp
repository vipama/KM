<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->
<%
'--------------------------------------------------------start block for sava data---------------------------------------------------

'--------------------------------------------------------end block for sava data----------------------------------------------------

dim chkPS,chkPC,chkQ,chkW
 chkPS=""
 chkPC=""
 chkQ=""
 chkW=""
if isEmpty(Request.QueryString("id")) = true then
	 if isEmpty(Request.Form("hidDid")) = false then
	 	getDid=Request.Form("hidDid")
	 else
	 	getDid = "1"
	 end if
else
	getDid=Request.QueryString("id")
end if

if isEmpty(Request.QueryString("tid")) <> true then
	if Request.QueryString("tid") = "PC" then
		getTid = "M_Main=1"
		chkPC = "checked=""checked"""
	elseif Request.QueryString("tid") = "PS" then
		getTid = "M_Reserve=1"
		chkPS= "checked=""checked"""
	end if
	if Request.QueryString("tid") = "W" then
		getTid = "M_Main=1"
		chkW= "checked=""checked"""
	elseif  Request.QueryString("tid") = "Q" then
		getTid = "M_Main=1"
		chkQ= "checked=""checked"""
	end if
else
	getTid = "M_Main=1"
	chkPC = "checked=""checked"""
end if
dim getMID
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>���ǹ��кǹ�ҹ</title>
<style type="text/css">
<!--
.style1 {
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
<script language="javascript">
function ChangeJobresultGroup(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			var typeID = getRadioValue("radioReviewType");
			window.location.href="FReview.asp?id="+val+"&oid="+val1+"&tid="+typeID;
		}else{
			var typeID = getRadioValue("radioReviewType");
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="FReview.asp?id="+strUser+"&oid="+val1+"&tid="+typeID;
		}
		
}
function goSave()
{
		if ((document.frmFReview.txtName.value != "")&&(document.frmFReview.txtPosition.value != "") )
		{
			document.frmFReview.hidS.value="S"
			document.frmFReview.submit();	
		}else{
			alert("��سҡ�͡������е��˹觴��¤��");
			document.frmFReview.txtName.focus();
		}
}
</script>
<script type="text/javascript" src="JScript/JS.js"></script>
</head>

<body>
<form name="frmFReview" method="post" enctype="application/x-www-form-urlencoded" action="FReview.asp">
<input type="hidden"  name="hidS" id="hidS" value=""/>
<div align="center"  style="font-size:24px;"><strong>Ẻ���ǹ��кǹ�ҹ</strong></div><br />
<table width="85%" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr><th align="right">No Review : <input type="text"  name="txtReviewNumber" id="txtReviewNumber" readonly="readonly" value="<%=(year(Now)+543)%>-7" /></th></tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>���Թ��÷��ǹ</td>
        <td><input type="radio" name="radioReviewType" id="radioReviewType1" value="Q" <%=chkQ%>onclick="ChangeJobresultGroup('','')"  />
          <label>�����ͤس�Ҿ</label>
&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType2" value="PC"  <%=chkPC%> onclick="ChangeJobresultGroup('','')"   />
          <label >�����͢�鹵͹��û�Ժѵԧҹ (P) (Core Process)</label></td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType3" value="W" <%=chkW%> onclick="ChangeJobresultGroup('','')" />
          <label >�����͢���Ըա�û�Ժѵԧҹ (W)</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;&nbsp;
          <input type="radio" name="radioReviewType" id="radioReviewType4" value="PS" <%=chkPS%>  onclick="ChangeJobresultGroup('','')" />
          <label >�����͢�鹵͹��û�Ժѵԧҹ (P) (Support Process)</label></td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="15%">˹��§ҹ :</td>
        <td width="85%">
             <%
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroup(this.value,1)" style="font-size:14px"   >
			  <%
			  while not recDepart.EOF
			  if recDepart("D_Id") = getDid then
			  selected = "selected=""selected"""
			  else
			  selected = ""
			  end if
			  %>
			  <option value="<%=recDepart("D_Id")%>" <%=selected%> ><%=recDepart("D_Name")%></option>
			  <%
			  recDepart.MoveNext
			  wend
			  recDepart.Close
			  Set recDepart = Nothing
			  %>
			  </select> 
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="15%">��кǹ�ҹ :</td>
        <td width="85%">
        <%
	  Set   recSOP = Server.CreateObject("ADODB.RECORDSET")
	  sqlSOP = "select  *  from  Tb_Manual where  D_Id='"&getDid&"' and "&getTid&" order by M_Id  asc"
	  recSOP.open sqlSOP,ConQS,1,3
	  %>
	  <select name="Manual" style="font-size:14px"  >
	  <%
	  while not recSOP.EOF
	'  if recSOP("M_Id") = getDid then
	'  selected = "selected=""selected"""
	'  else
	'  selected = ""
	'  end if
	  %>
	  <option value="<%=recSOP("M_Id")%>" <%=selected%> ><%response.write recSOP("M_Code")&" "&recSOP("M_Name")%></option>
	  <%
	  recSOP.MoveNext
	  wend
	  recSOP.Close
	  Set recSOP = Nothing
	  %>
      </select>    
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>���� : 
      <input name="txtName" type="text" id="txtName" size="60" /></td>
        <td>�˹� : 
      <input type="text" name="txtPosition" id="txtPosition" size="60" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">�š�÷��ǹ :</td>
        <td width="90%"><input name="radioPerfect" type="radio" id="radioPerfect" value="radioPerfect" checked="checked" />
          <label for="radioPerfect">�դ���������� ����ͧ���Թ�����</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">&nbsp;</td>
        <td>�˵ؼ� :</td>
        <td><input type="checkbox" name="chkCurrent" id="chkCurrent" />
          <label for="chkCurrent">�繻Ѩ�غѹ</label></td>
        <td><input type="checkbox" name="chkSupportWork" id="chkSupportWork" />
          <label for="chkSupportWork">�ʹ���ͧ�Ѻ��û�Ժѵԧҹ</label></td>
        <td><input type="checkbox" name="chkBelongManual" id="chkBelongManual" />
          <label for="chkBelongManual">�ա�ô��Թ��õ��������</label></td>
      </tr>
      <tr>
        <td width="10%">&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse" id="chkElse" />
          <label for="chkElse">����&nbsp;&nbsp;&nbsp;&nbsp;</label>
          <label for="txtElse"></label>
          <input name="txtElse" type="text" id="txtElse" size="70" /></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">�š�÷��ǹ :</td>
        <td width="90%"><input type="radio" name="radioPerfect" id="radioPerfect2" value="radioPerfect" />
          <label for="radioPerfect2">����դ���������� ��ͧ���Թ���</label></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%"><input type="radio" name="radioRemake" id="radioRemake1" value="radioRemake" checked="checked" />
          <label for="radioRemake">�Ѵ������</label></td>
        <td width="15%">�Ҵ��Ҩ����������ѹ���</td>
        <td>
        <label>
              <select name="finishDay" id="finishDay">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="17">17</option>
                <option value="18">18</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
              ��͹
              <select name="finishMonth" id="finishMonth">
                <option value="1" selected="selected">���Ҥ�</option>
                <option value="2">����Ҿѹ��</option>
                <option value="3">�չҤ�</option>
                <option value="4">����¹</option>
                <option value="5">����Ҥ�</option>
                <option value="6">�Զع�¹</option>
                <option value="7">�á�Ҥ�</option>
                <option value="8">�ԧ�Ҥ�</option>
                <option value="9">�ѹ��¹</option>
                <option value="10">���Ҥ�</option>
                <option value="11">��Ȩԡ�¹</option>
                <option value="12">�ѹ�Ҥ�</option>
              </select>
              ��
              <select name="finishYear" id="finishYear">
                <option value="2560">2560</option>
                <option value="2559">2559</option>
                <option value="2558">2558</option>
                <option value="2557" selected="selected">2557</option>
                                                        </select>
              </label>
        </td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake2" value="radioRemake" />
          <label for="radioRemake">���</label></td>
        <td width="15%">�Ҵ��Ҩ����������ѹ���</td>
        <td><label>
              <select name="finishDay" id="finishDay">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
                <option value="13">13</option>
                <option value="14">14</option>
                <option value="15">15</option>
                <option value="16">16</option>
                <option value="17">17</option>
                <option value="18">18</option>
                <option value="19">19</option>
                <option value="20">20</option>
                <option value="21">21</option>
                <option value="22">22</option>
                <option value="23">23</option>
                <option value="24">24</option>
                <option value="25">25</option>
                <option value="26">26</option>
                <option value="27">27</option>
                <option value="28">28</option>
                <option value="29">29</option>
                <option value="30">30</option>
                <option value="31">31</option>
              </select>
              ��͹
              <select name="finishMonth" id="finishMonth">
                <option value="1" selected="selected">���Ҥ�</option>
                <option value="2">����Ҿѹ��</option>
                <option value="3">�չҤ�</option>
                <option value="4">����¹</option>
                <option value="5">����Ҥ�</option>
                <option value="6">�Զع�¹</option>
                <option value="7">�á�Ҥ�</option>
                <option value="8">�ԧ�Ҥ�</option>
                <option value="9">�ѹ��¹</option>
                <option value="10">���Ҥ�</option>
                <option value="11">��Ȩԡ�¹</option>
                <option value="12">�ѹ�Ҥ�</option>
              </select>
              ��
              <select name="finishYear" id="finishYear">
                <option value="2560">2560</option>
                <option value="2559">2559</option>
                <option value="2558">2558</option>
                <option value="2557" selected="selected">2557</option>
                                                        </select>
              </label></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="radio" name="radioRemake" id="radioRemake-" value="radioRemake" />
          <label for="radioRemake">¡��ԡ</label></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="15%">�˵ؼ� :</td>
        <td><input type="checkbox" name="chkNotNow" id="chkNotNow" />
          <label for="chkNotNow">����繻Ѩ�غѹ</label></td>
        <td><input type="checkbox" name="chkNotSupportWork" id="chkNotSupportWork" />
          <label for="chkNotSupportWork">����ʹ���ͧ�Ѻ��û�Ժѵԧҹ</label></td>
        <td><input type="checkbox" name="chkNewWayWork" id="chkNewWayWork" />
          <label for="chkNewWayWork">���Ƿҧ��û�Ժѵԧҹ����</label></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3"><input type="checkbox" name="chkElse2" id="chkElse2" />
          <label for="chkElse2">���� 
            <input name="txtElse2" type="text" id="txtElse2" size="70" />
          </label></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>
          <input type="button" name="butSave" id="butSave" value="�ѹ�֡������" onclick="goSave()" /></td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
</body>
</html>
