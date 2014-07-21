// JavaScript Document
function alltrim(str)
{
	    var i;
	    var ch;
	    var retStr = '';
	    var retStr1 = '';
	    
	    if (str.length == 0)
	        return str;
	    for (i=0;i<str.length;i++)
	    {
	        ch = str.charAt(i);
	        if (ch == ' ' || ch == '\r' || ch == '\n') 
	            continue;
			 retStr += ch;
	    }
	    return retStr;
}
function isNumber(str){
	 	var i;
	    var ch;
	    var retStr=0;
	    var flag;
	    if (str.length == 0)
	        return false;
	    for (i=0;i<str.length;i++)
	    {
	        ch = str.charAt(i);
	        if (ch >= "0" && ch <= "9" ) 
			{
				 retStr++;
			}
		}
	if(str.length == retStr)
		{
			flag=true;	
		}
		else
		{
	    	flag=false;
		}
		return flag;
}
function autoCheck(typeName,getVal)
{
			
			if (typeName == "AuditAccept")
			{
					if (getVal == 0)
					{
							
							document.frmAuditor.OpenClose[0].checked=true;
					}else{
							document.frmAuditor.OpenClose[1].checked=true;
					}
					
			}else{
					if (getVal == 0)
					{
							
							document.frmAuditor.AuditAccept[0].checked=true;
					}else{
							document.frmAuditor.AuditAccept[1].checked=true;
					}
			}

}
function showSum(getTY,getval)
{
		
		if (getTY == "1")
		{
				document.frmAnalaysis.hidSum1.value=getval;
				document.frmAnalaysis.txtSumAll.value=(parseInt(document.frmAnalaysis.hidSum1.value)+parseInt(document.frmAnalaysis.hidSum2.value)+parseInt(document.frmAnalaysis.hidSum3.value)+parseInt(document.frmAnalaysis.hidSum4.value));
		}
		else if (getTY == "2")
		{
				document.frmAnalaysis.hidSum2.value=getval;
				document.frmAnalaysis.txtSumAll.value=(parseInt(document.frmAnalaysis.hidSum1.value)+parseInt(document.frmAnalaysis.hidSum2.value)+parseInt(document.frmAnalaysis.hidSum3.value)+parseInt(document.frmAnalaysis.hidSum4.value));
		}else if (getTY == "3")
		{
				document.frmAnalaysis.hidSum3.value=getval;
				document.frmAnalaysis.txtSumAll.value=(parseInt(document.frmAnalaysis.hidSum1.value)+parseInt(document.frmAnalaysis.hidSum2.value)+parseInt(document.frmAnalaysis.hidSum3.value)+parseInt(document.frmAnalaysis.hidSum4.value));
		}else if (getTY == "4")
		{
				document.frmAnalaysis.hidSum4.value=getval;
				document.frmAnalaysis.txtSumAll.value=(parseInt(document.frmAnalaysis.hidSum1.value)+parseInt(document.frmAnalaysis.hidSum2.value)+parseInt(document.frmAnalaysis.hidSum3.value)+parseInt(document.frmAnalaysis.hidSum4.value));
		}
		
}
function openReport()
{
	document.frmAnalaysis.action="showSingleReportAnalaysis.asp";
	document.frmAnalaysis.method="post";
	document.frmAnalaysis.submit();
}
function getRadioValue(name) {
    var group = document.getElementsByName(""+name+"");

    for (var i=0;i<group.length;i++) {
        if (group[i].checked) {
            return group[i].value;
        }
    }

    return '';
}
function openReportReview()
{
	document.frmFReview.action="showSingleReportReview.asp";
	document.frmFReview.method="post";
	document.frmFReview.submit();
}
function setPageEnable(valType)
{
	
	if(valType == "1")
	{
		//var objchkCurrent = document.getElementByID("chkCurrent");
		//alert("section 1");
		document.getElementById("chkCurrent").disabled=true;
     	document.getElementById("chkSupportWork").disabled=true;
		document.getElementById("chkBelongManual").disabled=true;
		document.getElementById("chkElse").disabled=true;
		document.getElementById("txtElse").disabled=true;
		document.getElementById("radioRemake1").disabled=false;
		document.getElementById("radioRemake2").disabled=false;
		document.getElementById("radioRemake3").disabled=false;
		document.getElementById("RemakefinishDay").disabled=false;
		document.getElementById("RemakefinishMonth").disabled=false;
		document.getElementById("RemakefinishYear").disabled=false;
		document.getElementById("EditfinishDay").disabled=false;
		document.getElementById("EditfinishMonth").disabled=false;
		document.getElementById("EditfinishYear").disabled=false;
		document.getElementById("chkNotNow").disabled=false;
		document.getElementById("chkNotSupportWork").disabled=false;
		document.getElementById("chkNewWayWork").disabled=false;
		document.getElementById("chkElse2").disabled=false;
		document.getElementById("txtElse2").disabled=false;
		
		//document.frmFReview.getElementByID("chkCurrent").disabled=true;
	}
	else if(valType == "2")
	{
		//alert("section 2");
		document.getElementById("chkCurrent").disabled=false;
     	document.getElementById("chkSupportWork").disabled=false;
		document.getElementById("chkBelongManual").disabled=false;
		document.getElementById("chkElse").disabled=false;
		document.getElementById("txtElse").disabled=false;
		document.getElementById("radioRemake1").disabled=true;
		document.getElementById("radioRemake2").disabled=true;
		document.getElementById("radioRemake3").disabled=true;
		document.getElementById("RemakefinishDay").disabled=true;
		document.getElementById("RemakefinishMonth").disabled=true;
		document.getElementById("RemakefinishYear").disabled=true;
		document.getElementById("EditfinishDay").disabled=true;
		document.getElementById("EditfinishMonth").disabled=true;
		document.getElementById("EditfinishYear").disabled=true;
		document.getElementById("chkNotNow").disabled=true;
		document.getElementById("chkNotSupportWork").disabled=true;
		document.getElementById("chkNewWayWork").disabled=true;
		document.getElementById("chkElse2").disabled=true;
		document.getElementById("txtElse2").disabled=true;
		
	}
}
function checkCountNum(typeP,PPosition)
{
			/*var mybrowser=navigator.userAgent;
			if(mybrowser.indexOf('MSIE')>0){
				
				var getValNum
				var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
				var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
				if(alltrim(obj.value) != "" && parseInt(obj1.value) == 0)
				{
					getValNum = parseInt(obj1.value);
					getValNum = getValNum+1;
					
					obj1.value=getValNum;
				}
				//checkText(typeP,PPosition);
				//showSumCARPAR(typeP);
			}
			
			if(mybrowser.indexOf('Firefox')>0){
				var getValNum
				var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
				var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
				if(obj.value.trim() != "" && parseInt(obj1.value) == 0)
				{
					getValNum = parseInt(obj1.value);
					getValNum = getValNum+1;
					
					obj1.value=getValNum;
				}
				checkText(typeP,PPosition);
				showSumCARPAR(typeP);
			}
			
			if(mybrowser.indexOf('Presto')>0){
				alert("Opera");
			}
			
			if(mybrowser.indexOf('Chrome')>0){
				var getValNum
				var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
				var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
				alert(obj.value);
				if(obj.value.trim() != "" && parseInt(obj1.value) == 0)
				{
					getValNum = parseInt(obj1.value);
					getValNum = getValNum+1;
					
					obj1.value=getValNum;
				}
				checkText(typeP,PPosition);
				showSumCARPAR(typeP);
			}*/
				var getValNum
				var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
				var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
				//alert(obj.value);
				if(alltrim(obj.value) != "" && parseInt(obj1.value) == 0)
				{
					getValNum = parseInt(obj1.value);
					getValNum = getValNum+1;
					
					obj1.value=getValNum;
					
				}
				checkText(typeP,PPosition);
				showSumCARPAR(typeP);
			
		
}
function checkText(typeP,PPosition)
{
			
		/*var mybrowser=navigator.userAgent;
		if(mybrowser.indexOf('MSIE')>0){
			//alert("IE");
			var getValNum
			var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
			var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
			getValNum = obj.value;
			if(getValNum == "" && parseInt(obj1.value) > 0)
			{
				getValNum = parseInt(obj1.value);
				getValNum = getValNum-1;
				//alert(getValNum);
				obj1.value=getValNum;
			}
		}
		if(mybrowser.indexOf('Firefox')>0){
			//alert("Firefox");
			var getValNum
			var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
			var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
			getValNum = obj.value;
			if(getValNum.trim() == "" && parseInt(obj1.value) > 0)
			{
				getValNum = parseInt(obj1.value);
				getValNum = getValNum-1;
				//alert(getValNum);
				obj1.value=getValNum;
			}
		}	
		if(mybrowser.indexOf('Presto')>0){
			alert("Opera");
		}			
		if(mybrowser.indexOf('Chrome')>0){
			//alert("Chrome");
			var getValNum
			var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
			var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
			getValNum = obj.value;
			if(getValNum.trim() == "" && parseInt(obj1.value) > 0)
			{
				getValNum = parseInt(obj1.value);
				getValNum = getValNum-1;
				//alert(getValNum);
				obj1.value=getValNum;
			}
		}*/
			var getValNum
			var obj = eval("document.getElementById('txt"+typeP+"Descript"+PPosition+"')");
			var obj1 = eval("document.getElementById('hidCountNum"+typeP+PPosition+"')");
			getValNum = obj.value;
			if(alltrim(getValNum) == "" && parseInt(obj1.value) > 0)
			{
				getValNum = parseInt(obj1.value);
				getValNum = getValNum-1;
				//alert(getValNum);
				obj1.value=getValNum;
			}

			
			
		 
}
function showSumCARPAR(typeP)
{
	var obj = eval("document.getElementById('txtNum"+typeP+"')");
	var sumNum;
	sumNum=0;
	for(var i=1;i<6;i++)
	{
		var obj1 = eval("document.getElementById('hidCountNum"+typeP+i+"')");
		sumNum = sumNum+parseInt(obj1.value);
	}
	obj.value = sumNum;
}
function chkAllowCARPAR()
{
	chkVal = document.getElementById("checkComplete").checked;
	var txtCar1,txtCar2,txtCar3,txtCar4,txtCar5;
	var txtPar1,txtPar2,txtPar3,txtPar4,txtPar5;
	var CheckFind,CheckNotFind;
	
	CheckFind = document.getElementById("checkFind");
	
	CheckNotFind = document.getElementById("checkNotFind");
	if(CheckNotFind.checked == true || CheckFind.checked == true )
	{
	CheckFind.disabled = !CheckFind.disabled;
	CheckNotFind.disabled = !CheckNotFind.disabled;
	
	txtCar1 = document.getElementById("txtCARDescript1");
	txtCar1.readOnly = !txtCar1.readOnly;
	txtCar2 = document.getElementById("txtCARDescript2");
	txtCar2.readOnly = !txtCar2.readOnly;
	txtCar3 = document.getElementById("txtCARDescript3");
	txtCar3.readOnly = !txtCar3.readOnly;
	txtCar4 = document.getElementById("txtCARDescript4");
	txtCar4.readOnly = !txtCar4.readOnly;
	txtCar5 = document.getElementById("txtCARDescript5");
	txtCar5.readOnly = !txtCar5.readOnly;
	
	txtPar1 = document.getElementById("txtPARDescript1");
	txtPar1.readOnly = !txtPar1.readOnly;
	txtPar2 = document.getElementById("txtPARDescript2");
	txtPar2.readOnly = !txtPar2.readOnly;
	txtPar3 = document.getElementById("txtPARDescript3");
	txtPar3.readOnly = !txtPar3.readOnly;
	txtPar4 = document.getElementById("txtPARDescript4");
	txtPar4.readOnly = !txtPar4.readOnly;
	txtPar5 = document.getElementById("txtPARDescript5");
	txtPar5.readOnly = !txtPar5.readOnly;
	
	objselctCARModerator1 = document.getElementById("selctCARModerator1");
	objselctCARModerator1.disabled = !objselctCARModerator1.disabled;
	objselctCARModerator2 = document.getElementById("selctCARModerator2");
	objselctCARModerator2.disabled = !objselctCARModerator2.disabled;
	objselctCARModerator3 = document.getElementById("selctCARModerator3");
	objselctCARModerator3.disabled = !objselctCARModerator3.disabled;
	objselctCARModerator4 = document.getElementById("selctCARModerator4");
	objselctCARModerator4.disabled = !objselctCARModerator4.disabled;
	objselctCARModerator5 = document.getElementById("selctCARModerator5");
	objselctCARModerator5.disabled = !objselctCARModerator5.disabled;
	
	objselctPARModerator1 = document.getElementById("selctPARModerator1");
	objselctPARModerator1.disabled = !objselctPARModerator1.disabled;
	objselctPARModerator2 = document.getElementById("selctPARModerator2");
	objselctPARModerator2.disabled = !objselctPARModerator2.disabled;
	objselctPARModerator3 = document.getElementById("selctPARModerator3");
	objselctPARModerator3.disabled = !objselctPARModerator3.disabled;
	objselctPARModerator4 = document.getElementById("selctPARModerator4");
	objselctPARModerator4.disabled = !objselctPARModerator4.disabled;
	objselctPARModerator5 = document.getElementById("selctPARModerator5");
	objselctPARModerator5.disabled = !objselctPARModerator5.disabled;
	}
/*	var objCheck;
	var objCAR;
	if(chkVal == true)
	{
		for(var i=1;i<6;i++)
		{
			objCAR = eval("document.getElementById('txtCARDescript'"+i+")");
			objCAR.readOnly=true;
		}
		for(var i=1;i<6;i++)
		{
			objPAR = eval("document.getElementById('txtPARDescript'"+i+")");
			objPAR.readOnly=true;
		}
	}
	else if(chkVal == false)
	{
		for(var i=1;i<6;i++)
		{
			objCAR = eval("document.getElementById('txtCARDescript'"+i+")");
			objCAR.readOnly=false;
		}
		for(var i=1;i<6;i++)
		{
			objPAR = eval("document.getElementById('txtPARDescript'"+i+")");
			objPAR.readOnly=false;
		}
	}*/
}
function chkAllowCAR()
{
	var txtCar1,txtCar2,txtCar3,txtCar4,txtCar5;
	txtCar1 = document.getElementById("txtCARDescript1");
	txtCar1.readOnly = !txtCar1.readOnly;
	txtCar2 = document.getElementById("txtCARDescript2");
	txtCar2.readOnly = !txtCar2.readOnly;
	txtCar3 = document.getElementById("txtCARDescript3");
	txtCar3.readOnly = !txtCar3.readOnly;
	txtCar4 = document.getElementById("txtCARDescript4");
	txtCar4.readOnly = !txtCar4.readOnly;
	txtCar5 = document.getElementById("txtCARDescript5");
	txtCar5.readOnly = !txtCar5.readOnly;
	
	objselctCARModerator1 = document.getElementById("selctCARModerator1");
	objselctCARModerator1.disabled = !objselctCARModerator1.disabled;
	objselctCARModerator2 = document.getElementById("selctCARModerator2");
	objselctCARModerator2.disabled = !objselctCARModerator2.disabled;
	objselctCARModerator3 = document.getElementById("selctCARModerator3");
	objselctCARModerator3.disabled = !objselctCARModerator3.disabled;
	objselctCARModerator4 = document.getElementById("selctCARModerator4");
	objselctCARModerator4.disabled = !objselctCARModerator4.disabled;
	objselctCARModerator5 = document.getElementById("selctCARModerator5");
	objselctCARModerator5.disabled = !objselctCARModerator5.disabled;
}
function chkAllowPAR()
{
	var txtPar1,txtPar2,txtPar3,txtPar4,txtPar5;
	txtPar1 = document.getElementById("txtPARDescript1");
	txtPar1.readOnly = !txtPar1.readOnly;
	txtPar2 = document.getElementById("txtPARDescript2");
	txtPar2.readOnly = !txtPar2.readOnly;
	txtPar3 = document.getElementById("txtPARDescript3");
	txtPar3.readOnly = !txtPar3.readOnly;
	txtPar4 = document.getElementById("txtPARDescript4");
	txtPar4.readOnly = !txtPar4.readOnly;
	txtPar5 = document.getElementById("txtPARDescript5");
	txtPar5.readOnly = !txtPar5.readOnly;
	
	objselctPARModerator1 = document.getElementById("selctPARModerator1");
	objselctPARModerator1.disabled = !objselctPARModerator1.disabled;
	objselctPARModerator2 = document.getElementById("selctPARModerator2");
	objselctPARModerator2.disabled = !objselctPARModerator2.disabled;
	objselctPARModerator3 = document.getElementById("selctPARModerator3");
	objselctPARModerator3.disabled = !objselctPARModerator3.disabled;
	objselctPARModerator4 = document.getElementById("selctPARModerator4");
	objselctPARModerator4.disabled = !objselctPARModerator4.disabled;
	objselctPARModerator5 = document.getElementById("selctPARModerator5");
	objselctPARModerator5.disabled = !objselctPARModerator5.disabled;
}
function IAuditSave()
{
	var objchkSave,objCheckfind,objCheckNotfind,checkComplete,flagcheck,flagcount;
	flagcheck = true;
	flagcount = 0;
	if(alltrim(document.getElementById("txtchif").value) == "" )
	{
		flagcount = flagcount+1;
	}
	if(alltrim(document.getElementById("txtFollower1").value) == "" )
	{
		flagcount = flagcount+1;
	}
	if(alltrim(document.getElementById("txtFollower2").value) == "" )
	{
		flagcount = flagcount+1;
	}
	if(alltrim(document.getElementById("txtFollower3").value) == "" )
	{
		flagcount = flagcount+1;
	}
	checkComplete = document.getElementById("checkComplete");
	if(checkComplete.checked == true && flagcount < 4)
	{
		alert("save");
		objchkSave = document.getElementById("hidSave");
		objchkSave.value="Save";
		document.frmInAudit.action="InternalAudit.asp";
		document.frmInAudit.submit();
	}else{
	   if(flagcount < 4)
	   {
		//-------------------------block check value in CAR and PAR textbox-------------------------
			var chkFlag,chkFlagCar,chkFlagPar,objCar,objPar;
			chkFlagCar=0;
			chkFlagPar=0;
			chkFlag=true;
			for(var i=1;i<6;i++)
			{
				objCar = eval("document.getElementById('txtCARDescript"+i+"')");
				objPar = eval("document.getElementById('txtPARDescript"+i+"')");
				if(alltrim(objCar.value) != "")
				{
					chkFlagCar = chkFlagCar+1;
				}
				if(alltrim(objPar.value) != "")
				{
					chkFlagPar = chkFlagPar+1;
				}
			}
		//------------------------------------------------------------------------------------------
		objCheckfind = document.getElementById("checkFind");
		objCheckNotfind = document.getElementById("checkNotFind");
		if(objCheckfind.checked == true && objCheckNotfind.checked == true &&  chkFlag == true)
		{
			if(chkFlagCar > 0 && chkFlagPar > 0 )
			{
				chkFlag=true;
			}else{
				alert("Please check CAR/PAR");
				chkFlag=false;
			}
		}
		if(objCheckfind.checked == true && objCheckNotfind.checked == false &&  chkFlag == true)
		{
			if(chkFlagCar > 0 )
			{
				chkFlag = true;
			}else{
				alert("Please check CAR");
				chkFlag=false;
			}
		}
		if(objCheckfind.checked == false && objCheckNotfind.checked == true &&  chkFlag == true)
		{
			if(chkFlagPar > 0 )
			{
				chkFlag = true;
			}else{
				alert("Please check PAR");
				chkFlag=false;
			}
		}
		if(objCheckfind.checked == false && objCheckNotfind.checked == false &&  chkFlag == true)
		{
			chkFlag = false;
			alert("Please select check box ! ");
		}
		if(chkFlag==true)
		{
			alert("save");
			objchkSave = document.getElementById("hidSave");
			objchkSave.value="Save";
			document.frmInAudit.action="InternalAudit.asp";
			document.frmInAudit.submit();
		}
	  }else{
		alert("Please insert audit name!");  
	  }
	}
}
function IAuditPrint()
{
		document.frmInAudit.action="PrintInternal.asp";
		document.frmInAudit.method="POST";
		document.frmInAudit.submit();
}
function changeSubDepart(subval)
{
		 
	 if(subval == 0)
	 {
			txtElse = document.getElementById("txtSubDepartElse");
			txtElse.readOnly = false;
			document.getElementById("spSubDepart").style.display="";
	 }else{
	 	 	txtElse = document.getElementById("txtSubDepartElse");
			txtElse.value="";
			txtElse.readOnly = true;
			document.getElementById("spSubDepart").style.display="none";
	}
	 	
}

function chkInternalAuditSource(gval)
{
	if(gval==5)
	{
		txtElse = document.getElementById("txtSourceElse5");
		txtElse.readOnly = false;
		document.getElementById("radio5").style.display="";
		
		txtElse = document.getElementById("txtSourceElse6");
		txtElse.readOnly = true;
		txtElse.value="";
		document.getElementById("radio6").style.display="none";
	}
	else if(gval == 6)
	{
		txtElse = document.getElementById("txtSourceElse6");
		txtElse.readOnly = false;
		document.getElementById("radio6").style.display="";
		
		txtElse = document.getElementById("txtSourceElse5");
		txtElse.readOnly = true;
		txtElse.value="";
		document.getElementById("radio5").style.display="none";
	}
	else
	{
		txtElse5 = document.getElementById("txtSourceElse5");
		txtElse5.value="";
		txtElse5.readOnly = true;
		txtElse6 = document.getElementById("txtSourceElse6");
		txtElse6.value="";
		txtElse6.readOnly = true;
		document.getElementById("radio5").style.display="none";
		document.getElementById("radio6").style.display="none";
	}
}

function AnalisCheckSave()
{
	var countGI,countGY,flagman;
	countGI=0;
	countGY=0;
	flagman= false;
	var gi1 = document.getElementById("chkStrategic1");
	var gi2 = document.getElementById("chkStrategic2");
	var gi3 = document.getElementById("chkStrategic3");
	
	var gy11 = document.getElementById("chkStrategy11");
	var gy12 = document.getElementById("chkStrategy12");
	var gy13 = document.getElementById("chkStrategy13");
	var gy14 = document.getElementById("chkStrategy14");
	var gy15 = document.getElementById("chkStrategy15");
	var gy16 = document.getElementById("chkStrategy16");
	var gy17 = document.getElementById("chkStrategy17");
	
	var gy21 = document.getElementById("chkStrategy21");
	var gy22 = document.getElementById("chkStrategy22");
	var gy23 = document.getElementById("chkStrategy23");
	var gy24 = document.getElementById("chkStrategy24");
	
	var gy31 = document.getElementById("chkStrategy31");
	var gy32 = document.getElementById("chkStrategy32");
	var gy33 = document.getElementById("chkStrategy33");
	var gy34 = document.getElementById("chkStrategy34");
	var gy35 = document.getElementById("chkStrategy35");
	
	var man = document.getElementById("Manual");
	if(man.value != "" && man.value != null)
	{
		flagman = true;
		
	}else{
		flagman = false;	
	}
	
	
	for(var i = 1;i <= 3;i++ )
	{
		var gi = eval("document.getElementById('chkStrategic"+i+"')")
		if(gi.checked == true)
		{
			countGI++;
		}
	}
	for(var i = 1;i <= 7;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy1"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	for(var i = 1;i <= 4;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy2"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	for(var i = 1;i <= 5;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy3"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	
	if(countGI > 0 && countGY > 0 && flagman == true)
	{
		alert("save data");
		document.frmAnalaysis.action="analaysis.asp";
		document.frmAnalaysis.method="POST";
		document.frmAnalaysis.submit();
		
	}else{
		//alert("ttttt");
		if(countGI == 0 || countGY == 0 && flagman == true)
		{
			alert("Please check Strategic and Strategy");
		}
		else if(countGI > 0 && countGY > 0 && flagman == false)
		{
			alert("Please check SOP! ");
		}
		else if(countGI == 0 && countGY == 0 && flagman == false)
		{
			alert("Please check SOP or Strategic or Strategy! ");
		}
	}
}
function AnalisCheckSaveUpdate()
{
	var countGI,countGY,flagman;
	countGI=0;
	countGY=0;
	flagman= false;
	var gi1 = document.getElementById("chkStrategic1");
	var gi2 = document.getElementById("chkStrategic2");
	var gi3 = document.getElementById("chkStrategic3");
	
	var gy11 = document.getElementById("chkStrategy11");
	var gy12 = document.getElementById("chkStrategy12");
	var gy13 = document.getElementById("chkStrategy13");
	var gy14 = document.getElementById("chkStrategy14");
	var gy15 = document.getElementById("chkStrategy15");
	var gy16 = document.getElementById("chkStrategy16");
	var gy17 = document.getElementById("chkStrategy17");
	
	var gy21 = document.getElementById("chkStrategy21");
	var gy22 = document.getElementById("chkStrategy22");
	var gy23 = document.getElementById("chkStrategy23");
	var gy24 = document.getElementById("chkStrategy24");
	
	var gy31 = document.getElementById("chkStrategy31");
	var gy32 = document.getElementById("chkStrategy32");
	var gy33 = document.getElementById("chkStrategy33");
	var gy34 = document.getElementById("chkStrategy34");
	var gy35 = document.getElementById("chkStrategy35");
	
	var man = document.getElementById("Manual");
	if(man.value != "" && man.value != null)
	{
		flagman = true;
		
	}else{
		flagman = false;	
	}
	
	
	for(var i = 1;i <= 3;i++ )
	{
		var gi = eval("document.getElementById('chkStrategic"+i+"')")
		if(gi.checked == true)
		{
			countGI++;
		}
	}
	for(var i = 1;i <= 7;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy1"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	for(var i = 1;i <= 4;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy2"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	for(var i = 1;i <= 5;i++ )
	{
		var gy = eval("document.getElementById('chkStrategy3"+i+"')")
		if(gy.checked == true)
		{
			countGY++;
		}
	}
	
	if(countGI > 0 && countGY > 0 && flagman == true)
	{
		alert("save data");
		document.frmAnalaysis.action="analaysis_update.asp";
		document.frmAnalaysis.method="POST";
		document.frmAnalaysis.submit();
		
	}else{
		//alert("ttttt");
		if(countGI == 0 || countGY == 0 && flagman == true)
		{
			alert("Please check Strategic and Strategy");
		}
		else if(countGI > 0 && countGY > 0 && flagman == false)
		{
			alert("Please check SOP! ");
		}
		else if(countGI == 0 && countGY == 0 && flagman == false)
		{
			alert("Please check SOP or Strategic or Strategy! ");
		}
	}
}
function goAnalisEdit()
{
	if(alltrim(document.getElementById("txtEditSOP").value) == "" )
	{
		alert("Please check SOP");
	}else{
		var s = document.getElementById("txtEditSOP").value;
		window.open("http://filing.fda.moph.go.th/kmfda/_block/qos/analaysis_update.asp?MID="+s,"_self");
	}
}
function goReviewEdit()
{
	if(alltrim(document.getElementById("txtREditSOP").value) == "" )
	{
		alert("Please check SOP");
	}else{
		var s = document.getElementById("txtREditSOP").value;
		window.open("http://filing.fda.moph.go.th/kmfda/_block/qos/FReview_update.asp?MC="+s,"_self");
	}
}
function goReviewReport()
{
		if(alltrim(document.getElementById("txtREditSOP").value) == "" )
		{
			alert("Please check SOP");
		}else{
		var s = document.getElementById("txtREditSOP").value; 
		window.open("http://filing.fda.moph.go.th/kmfda/_block/qos/showSingleReportReview1.asp?MID="+s,"_self"); 
		}
}
function goInternalAuditReport()
{
		if(alltrim(document.getElementById("txtREditSOP").value) == "" )
		{
			alert("Please check SOP");
		}else{
		var s = document.getElementById("txtREditSOP").value;
		window.open("http://filing.fda.moph.go.th/kmfda/_block/qos/PrintInternal1.asp?hidMC="+s,"_self"); 
		}
}
function goEditCar()
{
	document.getElementById("hidprint").value="Print";
	document.frmEdit.submit();
	
}
function goEditPar()
{	
	document.getElementById("hidprint").value="Print";
	document.frmEdit.submit();
	
}
function goFollowupPrint()
{
	document.getElementById("hidprint").value="Print";
	document.frmEdit.submit();
	
}
//-----------------------------------------------Block for ManagementReview.asp-----------------------------------------------
function ChangeJobresultGroupManagementReview(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			window.location.href="ManagementReview.asp?id="+val+"&oid="+val1;
		}else{
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="ManagementReview.asp?id="+strUser+"&oid="+val1;
		}
		
}
function ManagementReview_goSave()
{
		document.frmManagementReview.action="ManagementReview.asp";
		document.frmManagementReview.method="POST";
		document.frmManagementReview.hidSave.value="Save";
		document.frmManagementReview.submit();
}
function ManagementReview_goViewDoc(ID,DID)
{
	window.location.href="View_ManagementReview.asp?id="+ID+"&DID="+DID;
}
function ManagementReview_goEditDoc(ID,DID)
{
		window.location.href="Edit_ManagementReview.asp?id="+ID+"&DID="+DID;
}
function ManagementReview_goCancelDoc(ID,DID)
{
		var r = confirm("This document is being canceled! \n Do you want to continue or not.");
		if (r == true) {
			document.frmManagementReview.action="ManagementReview.asp";
			document.frmManagementReview.method="POST";
			document.frmManagementReview.hidSave.value="Cancel";
			document.getElementById("hidMRID").value = ID;
			document.getElementById("hidDid").value = DID;
			document.frmManagementReview.submit();
		} else {
			
		}
		
}
//-------------------------------------------End block for ManagementReview.asp-----------------------------------------------
//-------------------------------------------Block for Edit_ManagementReview.asp----------------------------------------------
function Edit_ManagementReview_goSave()
{
		document.frmManagementReview.action="Edit_ManagementReview.asp";
		document.frmManagementReview.method="POST";
		document.frmManagementReview.hidSave.value="Save";
		document.frmManagementReview.submit();
}
//------------------------------------------End block for Edit_ManagementReview.asp-------------------------------------------