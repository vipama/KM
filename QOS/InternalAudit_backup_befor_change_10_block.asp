<!--#include file="../../Config.inc.asp"-->
<!--#include file="../../Functions.lib.asp"-->

<%
if isEmpty(session("member")) = True then
	Response.Redirect("http://filing.fda.moph.go.th/kmfda/_block/qos/")
end if
dim Dateddmmyyyy
Dateddmmyyyy=Now()
Datemmddyyyy=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)
Datemmddyyyy1=month(Dateddmmyyyy)&"/"&day(Dateddmmyyyy)&"/"&year(Dateddmmyyyy)&" "&FormatDateTime(Dateddmmyyyy,3)
'----------------------------------------------------------------------Start block save data to DB---------------------------------------------------
dim getSave,chkError
dim getNameAdd
chkError = 0
getSave = Request.Form("hidSave")
if isEmpty(getSave)  = False then
	if getSave = "Save" then
		getTypeDepart = Request.Form("radioTypeDepart")
		getDepartID = Request.Form("DepartID")
		getSubDepartID = Request.Form("SubDepartID")
		gettxtSubDepartElse = Request.Form("txtSubDepartElse")
		getManual = Request.Form("Manual")
		gettxtchif = Request.Form("txtchif")
		gettxtFollower1 = Request.Form("txtFollower1")
		gettxtFollower2 = Request.Form("txtFollower2")
		gettxtFollower3 = Request.Form("txtFollower3")
		gettxtFollower4 = Request.Form("txtFollower4")
		gettxtFollower5 = Request.Form("txtFollower5")
		getradioAuditINOUT = Request.Form("radioAuditINOUT")
		gettxtSourceElse5 = Request.Form("txtSourceElse5")
		gettxtSourceElse6 = Request.Form("txtSourceElse6")
		getcheckComplete = Request.Form("checkComplete")
		getcheckFind = Request.Form("checkFind")
		getcheckNotFind = Request.Form("checkNotFind")
		gettxtNumCAR = Request.Form("txtNumCAR")
		gettxtNumPAR = Request.Form("txtNumPAR")
		gettxtCARDescript1 = Request.Form("txtCARDescript1")
		gettxtCARDescript2 = Request.Form("txtCARDescript2")
		gettxtCARDescript3 = Request.Form("txtCARDescript3")
		gettxtCARDescript4 = Request.Form("txtCARDescript4")
		gettxtCARDescript5 = Request.Form("txtCARDescript5")
		gettxtPARDescript1 = Request.Form("txtPARDescript1")
		gettxtPARDescript2 = Request.Form("txtPARDescript2")
		gettxtPARDescript3 = Request.Form("txtPARDescript3")
		gettxtPARDescript4 = Request.Form("txtPARDescript4")
		gettxtPARDescript5 = Request.Form("txtPARDescript5")
		getselctCARModerator1 = Request.Form("selctCARModerator1")
		getselctCARModerator2 = Request.Form("selctCARModerator2")
		getselctCARModerator3 = Request.Form("selctCARModerator3")
		getselctCARModerator4 = Request.Form("selctCARModerator4")
		getselctCARModerator5 = Request.Form("selctCARModerator5")
		getselctPARModerator1 = Request.Form("selctPARModerator1")
		getselctPARModerator2 = Request.Form("selctPARModerator2")
		getselctPARModerator3 = Request.Form("selctPARModerator3")
		getselctPARModerator4 = Request.Form("selctPARModerator4")
		getselctPARModerator5 = Request.Form("selctPARModerator5")
		gettxtDes = Request.Form("txtDes")
		gettxtGoodDes = Request.Form("txtGoodDes")
		gettxtBadDes = Request.Form("txtBadDes")
		gettxtShowQMR = Request.Form("txtShowQMR")
		
		getInternalAuditDay = Request.Form("InternalAuditDay")
		getInternalAuditMonth = Request.Form("InternalAuditMonth")
		getInternalAuditYear = Request.Form("InternalAuditYear")
		
		if getInternalAuditDay <> "0" and getInternalAuditMonth <> "0" and getInternalAuditYear <> "0" then
			Datemmddyyyy = getInternalAuditMonth&"/"&getInternalAuditDay&"/"&getInternalAuditYear
		end if
		
		dim showSource_Name
		if  getradioAuditINOUT  = "1" then
				showSource_Name = "การตรวจติดตามคุณภาพภายใน"
		elseif getradioAuditINOUT = "2" then
			 	showSource_Name = "การตรวจติดตามคุณภาพภายนอก"
		elseif getradioAuditINOUT = "3" then
				showSource_Name = "การประชุมทบทวนโดยฝ่ายบริหาร"
		elseif getradioAuditINOUT = "4" then
				showSource_Name = "การปฏิบัติงาน"
		elseif getradioAuditINOUT = "5" then
				showSource_Name = gettxtSourceElse5
		elseif getradioAuditINOUT = "6" then
				showSource_Name = gettxtSourceElse6
		end if
		if getcheckComplete = "on" then
				dim getCountIDInternal 
			'	if checkInternalAuditData(GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual),getTypeDepart,(year(Dateddmmyyyy)+543)) = 0 then
						ConQS.BeginTrans
						if GetSingleFieldQS("Tb_Internalaudit","top 1 ID","") = 0 then
							getCountIDInternal =1
						else
							getCountIDInternal = (GetSingleFieldQS("Tb_Internalaudit","top 1 ID"," order by ID DESC")+1)
						end if 
						'response.write getradioAuditINOUT&""
						SQL = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_Year,Audit_Flag_Complete,Audit_QMR_P1) values ('C','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-15-"&getCountIDInternal&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&(year(Dateddmmyyyy)+543)&"','0','"&gettxtShowQMR&"')"
						'response.write SQL
						ConQS.execute(SQL)
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-15-"&getCountIDInternal&"','C')"
						ConQS.execute(sql_log)
						If Err.Number = 0 Then
						  ConQS.CommitTrans
						  response.write "<script language=""javascript"">"
						  response.write "alert(""บันทึกข้อมูลเรียบร้อย"");"
						  response.write "</script>"
						Else
						   ConQS.RollbackTrans
						End If
			'	else
			'			  response.write "<script language=""javascript"">"
			'			  response.write "alert(""มีข้อมูลอยู่แล้วไม่สามารถบันทึกซ้ำได้"");"
			'			  response.write "window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos/Internalaudit.asp';"
						 
			'	end if
		elseif getcheckComplete = "" then
		'	if checkInternalAuditData(GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual),getTypeDepart,(year(Dateddmmyyyy)+543)) = 0 then
			   ConQS.BeginTrans
				if GetSingleFieldQS("Tb_Internalaudit","top 1 ID","") = 0 then
					getCountIDInternal =1
				else
					getCountIDInternal = (GetSingleFieldQS("Tb_Internalaudit","top 1 ID"," order by ID DESC")+1)
				end if 
				'response.write getradioAuditINOUT&""
				SQL = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_Year,Audit_Flag_Complete,Audit_QMR_P1) values ('C','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-15-"&getCountIDInternal&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&(year(Dateddmmyyyy)+543)&"','1','"&gettxtShowQMR&"')"
				'response.write SQL
				ConQS.execute(SQL)
				sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-15-"&getCountIDInternal&"','C')"
				ConQS.execute(sql_log)
				If Err.Number = 0 Then
				  ConQS.CommitTrans
				  response.write "<script language=""javascript"">"
				  response.write "alert(""บันทึกข้อมูลเรียบร้อย"");"
				  response.write "</script>"
				Else
				   ConQS.RollbackTrans
				End If
		 
			'=================================เริ่มบล๊อค     นี่เป็นบล๊อคถ้ามีการติ๊กที่ปุ่มพบหลักฐานแสดงว่าเกิดข้อบกพร่อง===========================================	
				if getcheckFind = "on" then
				 ConQS.BeginTrans
				'----------------------------------------------------------------car 1-----------------------------------------------------------------
					if trim(gettxtCARDescript1) <> "" then
						dim getNumCar1
						
						Select  case getselctCARModerator1
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumCar1 = (GetSingleFieldQS("Tb_RunNumCar","top 1 ID"," order by ID DESC")+1)
						SQL_CAR1 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('NC','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar1&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtCARDescript1&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						'response.write "<br> 1 ="&SQL_CAR1
						ConQS.execute(SQL_CAR1) 
						ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-16-"&getNumCar1)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar1&"','NC')"
						ConQS.execute(sql_log)	
					end if
				 '---------------------------------------------------------------------------------------------------------------------------------------
				 '---------------------------------------------------------------car 2------------------------------------------------------------------		
					if trim(gettxtCARDescript2) <> "" then
						dim getNumCar2
						Select  case getselctCARModerator2
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumCar2 = (GetSingleFieldQS("Tb_RunNumCar","top 1 ID"," order by ID DESC")+1)
						SQL_CAR2 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('NC','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar2&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtCARDescript2&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						'response.write "<br> 2 ="&SQL_CAR2
						ConQS.execute(SQL_CAR2)
						ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-16-"&getNumCar2)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')") 
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar2&"','NC')"
						ConQS.execute(sql_log) 
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				'---------------------------------------------------------------car 3-----------------------------------------------------------------------		
					if trim(gettxtCARDescript3) <> "" then
						dim getNumCar3
						
						Select  case getselctCARModerator3
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumCar3 = (GetSingleFieldQS("Tb_RunNumCar","top 1 ID"," order by ID DESC")+1)
						SQL_CAR3 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('NC','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar3&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtCARDescript3&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						'response.write "<br> 3 ="&SQL_CAR3
						ConQS.execute(SQL_CAR3)
						ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-16-"&getNumCar3)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')") 
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar3&"','NC')"
						ConQS.execute(sql_log) 
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				 '---------------------------------------------------------------car 4------------------------------------------------------------------		
					if trim(gettxtCARDescript4) <> "" then
						dim getNumCar4
						
						Select  case getselctCARModerator4
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumCar4 = (GetSingleFieldQS("Tb_RunNumCar","top 1 ID"," order by ID DESC")+1)
						SQL_CAR4 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('NC','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar4&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtCARDescript4&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						'response.write "<br> 4 ="&SQL_CAR4 
						ConQS.execute(SQL_CAR4)
						ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-16-"&getNumCar4)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar4&"','NC')"
						ConQS.execute(sql_log) 
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				'---------------------------------------------------------------car 5------------------------------------------------------------------		
					if trim(gettxtCARDescript5) <> "" then
						dim getNumCar5
						
						Select  case getselctCARModerator5
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumCar5 = (GetSingleFieldQS("Tb_RunNumCar","top 1 ID"," order by ID DESC")+1)
						SQL_CAR5 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('NC','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar5&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtCARDescript5&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						'response.write "<br> 5 ="&SQL_CAR5
						ConQS.execute(SQL_CAR5)
						ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-16-"&getNumCar5)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-16-"&getNumCar5&"','NC')"
						ConQS.execute(sql_log)  
					end if
					 If Err.Number = 0 Then
						  ConQS.CommitTrans
						  response.write "<script language=""javascript"">"
						    response.write "alert(""บันทึกข้อมูลใบ CAR เรียบร้อย"");"
						  response.write "</script>"
					Else
					      ConQS.RollbackTrans
					End If
				'-------------------------------------------------------------------------------------------------------------------------------------------
				end if
			'=================================จบบล๊อค     นี่เป็นบล๊อคถ้ามีการติ๊กที่ปุ่มพบหลักฐานแสดงว่าเกิดข้อบกพร่อง===========================================	
			'=================================เริ่มบล๊อค     นี่เป็นบล๊อคถ้ามีการติ๊กที่ปุ่มพบความมีแนวโน้มที่จะเกิดข้อบำพร่อง===========================================	
				if getcheckNotFind = "on" then
					ConQS.BeginTrans
					'----------------------------------------------------------------par 1-----------------------------------------------------------------
					if trim(gettxtPARDescript1) <> "" then
						dim getNumPar1
						
						Select  case getselctPARModerator1
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumPar1 = (GetSingleFieldQS("Tb_RunNumPar","top 1 ID"," order by ID DESC")+1)
						SQL_PAR1 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('OBS','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar1&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtPARDescript1&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						
						'response.write "<br> 1 ="&SQL_PAR1
						ConQS.execute(SQL_PAR1)
						ConQS.Execute("insert into  Tb_RunNumPAR (PAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-17-"&getNumPar1)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar1&"','OBS')"
						ConQS.execute(sql_log)  
					end if
				 '---------------------------------------------------------------------------------------------------------------------------------------
				 '---------------------------------------------------------------par 2------------------------------------------------------------------		
					if trim(gettxtPARDescript2) <> "" then
						dim getNumPar2
					
						Select  case getselctPARModerator2
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumPar2 = (GetSingleFieldQS("Tb_RunNumPar","top 1 ID"," order by ID DESC")+1)
						SQL_PAR2 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('OBS','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar2&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtPARDescript2&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						
						'response.write "<br> 2 ="&SQL_PAR2
						ConQS.execute(SQL_PAR2)
						ConQS.Execute("insert into  Tb_RunNumPAR (PAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-17-"&getNumPar2)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar2&"','OBS')"
						ConQS.execute(sql_log)  
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				'---------------------------------------------------------------par 3-----------------------------------------------------------------------		
					if trim(gettxtPARDescript3) <> "" then
						dim getNumPar3
					
						Select  case getselctPARModerator3
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumPar3 = (GetSingleFieldQS("Tb_RunNumPar","top 1 ID"," order by ID DESC")+1)
						SQL_PAR3 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('OBS','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar3&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtPARDescript3&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						
						'response.write "<br> 3 ="&SQL_PAR3
						ConQS.execute(SQL_PAR3)
						ConQS.Execute("insert into  Tb_RunNumPAR (PAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-17-"&getNumPar3)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar3&"','OBS')"
						ConQS.execute(sql_log)  
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				 '---------------------------------------------------------------par 4------------------------------------------------------------------		
					if trim(gettxtPARDescript4) <> "" then
						dim getNumPar4
					
						Select  case getselctPARModerator4
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumPar4 = (GetSingleFieldQS("Tb_RunNumPar","top 1 ID"," order by ID DESC")+1)
						SQL_PAR4 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('OBS','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar4&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtPARDescript4&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						
						'response.write "<br> 4 ="&SQL_PAR4
						ConQS.execute(SQL_PAR4)
						ConQS.Execute("insert into  Tb_RunNumPAR (PAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-17-"&getNumPar5)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar4&"','OBS')"
						ConQS.execute(sql_log)  
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
				'---------------------------------------------------------------par 5------------------------------------------------------------------		
					if trim(gettxtPARDescript5) <> "" then
						dim getNumPar5
					
						Select  case getselctPARModerator5
							case "1"
								getNameAdd = gettxtchif
							case "2"
								 getNameAdd = gettxtFollower1
							case "3"
								getNameAdd = gettxtFollower2
							case "4"
								getNameAdd = gettxtFollower3
							case "5"
								getNameAdd = gettxtFollower4
							case "6"
								getNameAdd = gettxtFollower5
						End select
						getNumPar5 = (GetSingleFieldQS("Tb_RunNumPar","top 1 ID"," order by ID DESC")+1)
						SQL_PAR5 = "insert into Tb_Internalaudit  (Audit_DocType,Audit_Level,No_Car_Par,Audit_Date,Audit_Source,Audit_Source_Details,Audit_Depart,Audit_Subdepart,Audit_SubDepartElseName,M_Code,M_Name,Audit_Defect,Audit_Name1,Audit_Name2,Audit_Name3,Audit_Name4,Audit_Name5,Audit_Name6,Audit_Descript,Audit_Advantages,Audit_Disadvantages,Audit_License_P1,Audit_Year,Audit_QMR_P1) values ('OBS','"&getTypeDepart&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar5&"','"&Datemmddyyyy&"','"&getradioAuditINOUT&"','"&showSource_Name&"','"&getDepartID&"','"&getSubDepartID&"','"&gettxtSubDepartElse&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&GetSingleFieldQS("Tb_Manual","M_Name","where  M_Id="&getManual)&"','"&gettxtPARDescript5&"','"&gettxtchif&"','"&gettxtFollower1&"','"&gettxtFollower2&"','"&gettxtFollower3&"','"&gettxtFollower4&"','"&gettxtFollower5&"','"&gettxtDes&"','"&gettxtGoodDes&"','"&gettxtBadDes&"','"&getNameAdd&"','"&(year(Dateddmmyyyy)+543)&"','"&gettxtShowQMR&"')"
						
						'response.write "<br> 5 ="&SQL_PAR5
						ConQS.execute(SQL_PAR5)
						ConQS.Execute("insert into  Tb_RunNumPAR (PAR_ID,M_Code) values ('"&(year(Dateddmmyyyy)+543&"-17-"&getNumPar5)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"')")
						sql_log = "insert into Tb_LogInternalAudit (User_Id,Method_Access,Date_Access,Department_Name,M_Code,No_Car_Par,Audit_DocType) values ('"&session("member")&"','Add','"&Datemmddyyyy1&"','"&getDepartmentname(getDepartID)&"','"&GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual)&"','"&year(Dateddmmyyyy)+543&"-17-"&getNumPar5&"','OBS')"
						ConQS.execute(sql_log)  
					end if
				'-------------------------------------------------------------------------------------------------------------------------------------------
					If Err.Number = 0 Then
						  ConQS.CommitTrans
						  response.write "<script language=""javascript"">"
						  response.write "alert(""บันทึกข้อมูลใบ PAR เรียบร้อย"");"
						  response.write "</script>"
					Else
					      ConQS.RollbackTrans
					End If
			end if
			'=================================จบบล๊อค     นี่เป็นบล๊อคถ้ามีการติ๊กที่ปุ่มพบความมีแนวโน้มที่จะเกิดข้อบำพร่อง===========================================	
	'   else
	 '   				  response.write "<script language=""javascript"">"
	'					  response.write "alert(""มีข้อมูลอยู่แล้วไม่สามารถบันทึกซ้ำได้"");"
	'					  response.write "window.location.href='http://filing.fda.moph.go.th/kmfda/_block/qos/Internalaudit.asp';"
						 
	'   end if	
		
	end if
		
	end if
end if
'----------------------------------------------------------------------End block save data to DB-----------------------------------------------------
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
	if Request.QueryString("tid") = "1" then
		getTid = "main"
		FlagMain_Reserve = " and M_Main=1"
		chkMain = "checked=""checked"""
	elseif Request.QueryString("tid") = "2" then
		getTid = "submain"
		FlagMain_Reserve = ""
		chkSubmain= "checked=""checked"""
	end if
else
	getTid = "main"
	FlagMain_Reserve = "and M_Main=1"
	chkMain = "checked=""checked"""
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874" />
<title>แบบการตรวจติดตามคุณภาพภายใน</title>
<style type="text/css">
<!--
.style1 {
font-size:13px;
font-family:Arial, Helvetica, sans-serif;


}
-->
</style>
<script language="javascript">
function ChangeJobresultGroupInternal(val,val1)
{
		
		if ((val != "" ) || (val1 != ""))
		{ 
			var typeID = getRadioValue("radioTypeDepart");
			window.location.href="InternalAudit.asp?id="+val+"&oid="+val1+"&tid="+typeID;
		}else{
			var typeID = getRadioValue("radioTypeDepart");
			var e = document.getElementById("DepartID");    
			var strUser = e.options[e.selectedIndex].value;
			window.location.href="InternalAudit.asp?id="+strUser+"&oid="+val1+"&tid="+typeID;
		}
		
}
</script>
<script type="text/javascript" src="JScript/JS.js"></script>
</head>

<body>
<div align="center" style="font-size:28px; font-style:oblique">แบบการตรวจติดตามคุณภาพภายใน</div>
<form name="frmInAudit" method="post" enctype="application/x-www-form-urlencoded">
<input type="hidden"  name="hidCountNumCAR1" value="0" id="hidCountNumCAR1"/>
<input type="hidden"  name="hidCountNumCAR2" value="0" id="hidCountNumCAR2"/>
<input type="hidden"  name="hidCountNumCAR3" value="0" id="hidCountNumCAR3"/>
<input type="hidden"  name="hidCountNumCAR4" value="0" id="hidCountNumCAR4"/>
<input type="hidden"  name="hidCountNumCAR5" value="0" id="hidCountNumCAR5"/>
<input type="hidden"  name="hidCountNumPAR1" value="0" id="hidCountNumPAR1"/>
<input type="hidden"  name="hidCountNumPAR2" value="0" id="hidCountNumPAR2"/>
<input type="hidden"  name="hidCountNumPAR3" value="0" id="hidCountNumPAR3"/>
<input type="hidden"  name="hidCountNumPAR4" value="0" id="hidCountNumPAR4"/>
<input type="hidden"  name="hidCountNumPAR5" value="0" id="hidCountNumPAR5"/>
<input type="hidden" name="hidMQR" id="hidQMR" value="<%=GetSingleFieldQS("Tb_Qmr","Q_Id","where  D_Id='"&getDid&"'")%>" />
<input type="hidden" name="hidSave" id="hidSave" value="" />
<% if isEmpty(getSave)  = False then %>
<input type="hidden" id="hidMC" name="hidMC" value="<%=GetSingleFieldQS("Tb_Manual","M_Code","where  M_Id="&getManual) %>" >
<% end if %>
<table width="85%" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td width="20%">&nbsp;</td>
    <td width="80%" class="style1"><label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart1" value="1" <%=chkMain%> onClick="ChangeJobresultGroupInternal('','')" />
    ระดับกรม</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <label>
      <input type="radio" name="radioTypeDepart" id="radioTypeDepart2" value="2" <%=chkSubmain%> onClick="ChangeJobresultGroupInternal('','')" />
    ระดับหน่วยงาน</label>    </td>
  </tr>
  <tr>
    <td class="style1">หน่วยงาน :</td>
    <td>
    <%
			  Set   recDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlDepart = "select  *  from  Tb_Department order by D_Numberlist  asc"
			  recDepart.open sqlDepart,ConQS,1,3
			  %>
			  <select name="DepartID" id="DepartID" onChange="ChangeJobresultGroupInternal(this.value,1)" style="font-size:14px"   >
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
			  </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			  <span class="style11">รายชื่อ QMR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			  <input name="txtShowQMR" type="text" class="style1" id="txtShowQMR" size="60" value="<%=GetSingleFieldQS("Tb_Qmr","Q_Name","where  D_Id='"&getDid&"'")%>" readonly />
			  </span></td>
  </tr>
  <tr>
    <td class="style1">หน่วยงานย่อย :</td>
    <td>
   		 <%
			  Set   recSubDepart = Server.CreateObject("ADODB.RECORDSET")
			  sqlSubDepart = "select  *  from  Tb_SubDepart where  D_Id='"&getDid&"' order by Subdepart_ID  asc"
			  'response.write sqlSubDepart 
			  recSubDepart.open sqlSubDepart,ConQS,1,3
			  %>
			  <select name="SubDepartID" id="SubDepartID" style="font-size:14px"  onchange="changeSubDepart(this.value)"   >
			  <%
			  while not recSubDepart.EOF
			  if recSubDepart("Subdepart_ID") = getDid then
			  selected = "selected=""selected"""
			  else
			  selected = ""
			  end if
			  %>
			  <option value="<%=recSubDepart("Subdepart_ID")%>" <%=selected%> ><%=recSubDepart("Name_Subdepart")%></option>
			  <%
			  recSubDepart.MoveNext
			  wend
			  recSubDepart.Close
			  Set recSubDepart = Nothing
			  %>
               <option value="0" >อื่นๆ (ระบุ)</option>
			  </select>&nbsp;&nbsp;
			  <span id="spSubDepart" style="display:none"><input type="text" name="txtSubDepartElse" id="txtSubDepartElse" readonly  /></span>			  </td>
  </tr>
  <tr>
    <td class="style1">กระบวนงาน :</td>
    <td>
    
    <%
	  Set   recSOP = Server.CreateObject("ADODB.RECORDSET")
	  	sqlSOP = "select  *  from  Tb_Manual where  D_Id='"&getDid&"' "&FlagMain_Reserve&" or M_Public = True order by M_Id  asc"
	  'response.write sqlSOP 
	  recSOP.open sqlSOP,ConQS,1,3
	  %>
	  <select name="Manual" id="Manual" style="font-size:14px"  >
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
      </select>    </td>
  </tr>
  <tr>
    <td class="style1">ผู้ตรวจติดตาม :</td>
    <td class="style1">(1) 
      <input name="txtchif" type="text" id="txtchif" size="60" />
      หัวหน้าผู้ตรวจติดตาม</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="style1"><label>
      (2) <input name="txtFollower1" type="text" id="txtFollower1" size="60" />
      ผู้ตรวจติดตามคนที่ 1
    </label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="style1"><label>
      (3) <input name="txtFollower2" type="text" id="txtFollower2" size="60" />
    </label>
     ผู้ตรวจติดตามคนที่ 2</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="style1"><label>
      (4) <input name="txtFollower3" type="text" id="txtFollower3" size="60" />
      ผู้ตรวจติดตามคนที่ 3
    </label></td>
  </tr>
  <tr>
    <td class="style1">&nbsp;</td>
    <td class="style1"><label>(5)
      <input name="txtFollower4" type="text" id="txtFollower4" size="60" />
ผู้ตรวจติดตามคนที่ 4 </label></td>
  </tr>
  <tr>
    <td class="style1">&nbsp;</td>
    <td class="style1"><label>(6)
      <input name="txtFollower5" type="text" id="txtFollower5" size="60" />
ผู้ตรวจติดตามคนที่ 5 </label></td>
  </tr>
  <tr>
  <td class="style1">วันที่ตรวจติดตาม :</td>
  <td>&nbsp;&nbsp;&nbsp;&nbsp;
  <select name="InternalAuditDay" id="InternalAuditDay">
   	   <option  value="0" selected="selected">วัน</option>
   <% for i=1 to 31 %>
      <option value="<%=i%>"><%=i%></option>
   <% next %>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;
  <select name="InternalAuditMonth" id="InternalAuditMonth" >
        <option  value="0" selected="selected">เดือน</option>
        <option value="1">มกราคม</option>
        <option value="2">กุมภาพันธ์</option>
        <option value="3">มีนาคม</option>
        <option value="4">เมษายน</option>
        <option value="5">พฤษภาคม</option>
        <option value="6">มิถุนายน</option>
        <option value="7">กรกฎาคม</option>
        <option value="8">สิงหาคม</option>
        <option value="9">กันยายน</option>
        <option value="10">ตุลาคม</option>
        <option value="11">พฤศจิกายน</option>
        <option value="12">ธันวาคม</option>
  </select>&nbsp;&nbsp;&nbsp;&nbsp;
  <select name="InternalAuditYear" id="InternalAuditYear" >
        <option value="2020">2563</option>
        <option value="2019">2562</option>
        <option value="2018">2561</option>
        <option value="2017">2560</option>
        <option value="2016">2559</option>
        <option value="2015">2558</option>
        <option value="2014">2557</option>
        <option value="0" selected="selected">ปี</option>
  </select>  </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="style1">ที่มา :</td>
    <td class="style1"><label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditIN" value="1" checked="checked" onClick="chkInternalAuditSource(this.value)" />
      การตรวจติดตามคุณภาพภายใน</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditOUT" value="2" onClick="chkInternalAuditSource(this.value)" />
      การตรวจติดตามคุณภาพภายนอก</label>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditOUT" value="3"  onclick="chkInternalAuditSource(this.value)" />
      การประชุมทบทวนโดยฝ่ายบริหาร</label><br />
      <label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditOUT" value="4"  onclick="chkInternalAuditSource(this.value)" />
      การปฏิบัติงาน</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditOUT" value="5" onClick="chkInternalAuditSource(this.value)" />
      ข้อร้องเรียนจาก</label>&nbsp;
<span id="radio5" style="display:none"><input type="text" name="txtSourceElse5" id="txtSourceElse5" readonly  /></span>      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <label>
      <input type="radio" name="radioAuditINOUT" id="radioAuditOUT" value="6"  onclick="chkInternalAuditSource(this.value)" />
      อื่นๆ</label>&nbsp;&nbsp;<span id="radio6" style="display:none"><input type="text" name="txtSourceElse6" id="txtSourceElse6" readonly  /></span>
      </td>
  </tr>
  <tr>
    <td class="style1">ผลการตรวจติดตาม </td>
    <td class="style1"><label>
      <input type="checkbox" name="checkComplete" id="checkComplete" onClick="chkAllowCARPAR()"  />
      ไม่พบข้อบกพร่อง</label></td>
  </tr>
  <tr><td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="2">
    <tr>
    <td colspan="2" class="style1">
     <label>
      <input name="checkFind" type="checkbox" id="checkFind"  onclick="chkAllowCAR()"   />
      พบหลักฐานที่แสดงว่าเกิดข้อบกพร่องหรือความไม่สอดคล้องขึ้นในคุณภาพ</label>
      &nbsp;&nbsp;<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;จึงออกใบ CAR จำนวน 
      <label>
      <input name="txtNumCAR" type="text" id="txtNumCAR" size="5" readonly value="0" align="middle" />
      </label>
      ใบ    </td>
    <td colspan="2" class="style1">
    <label>
      <input name="checkNotFind" type="checkbox" id="checkNotFind"   onclick="chkAllowPAR()" />
      พบความมีแนวโน้มที่จะเกิดข้อบกพร่องหรือความไม่สอดคล้องขึ้นในระบบคุณภาพ</label>
      &nbsp;&nbsp;<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      จึงออกใบ PAR จำนวน 
      <label>
      <input name="txtNumCAR" type="text" id="txtNumPAR" size="5" readonly value="0" align="middle" />
      </label>
      ใบ    </td>
    </tr>
    <tr>
      <td width="15%" class="style1"><p>ข้อบกพร่องที่พบ</p>
          <p>CAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle" class="style1">1.
        <label>
          <textarea name="txtCARDescript1" cols="80" rows="5" id="txtCARDescript1"   onchange="checkCountNum('CAR','1')"   readonly="readonly"    ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctCARModerator1" id="selctCARModerator1" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br /></td>
      <td width="15%" class="style1"><p>แนวโน้มข้อบกพร่องที่พบ</p>
        <p>PAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle" class="style1">1.
        <label>
          <textarea name="txtPARDescript1" cols="80" rows="5" id="txtPARDescript1"   onchange="checkCountNum('PAR','1')" readonly     ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctPARModerator1" id="selctPARModerator1" disabled="disabled" >
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
    </tr>
    <tr>
      <td class="style1"><p>ข้อบกพร่องที่พบ</p>
        <p>CAR No.(AutoNumber) </p></td>
      <td class="style1">2.
        <label>
          <textarea name="txtCARDescript2" cols="80" rows="5" id="txtCARDescript2" onChange="checkCountNum('CAR','2')" readonly></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctCARModerator2" id="selctCARModerator2" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
      <td width="15%" class="style1"><p>แนวโน้มข้อบกพร่องที่พบ</p>
          <p>PAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle">2.
        <label>
        <textarea name="txtPARDescript2" cols="80" rows="5" id="txtPARDescript2"   onchange="checkCountNum('PAR','2')" readonly    ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctPARModerator2" id="selctPARModerator2" disabled="disabled" >
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
    </tr>
    <tr>
      <td class="style1"><p>ข้อบกพร่องที่พบ</p>
        <p>CAR No.(AutoNumber) </p></td>
      <td class="style1">3.
        <label>
          <textarea name="txtCARDescript3" cols="80" rows="5" id="txtCARDescript3" onChange="checkCountNum('CAR','3')" readonly></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctCARModerator3" id="selctCARModerator3" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
      <td width="15%" class="style1"><p>แนวโน้มข้อบกพร่องที่พบ</p>
          <p>PAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle" class="style1">3.
        <label>
          <textarea name="txtPARDescript3" cols="80" rows="5" id="txtPARDescript3"   onchange="checkCountNum('PAR','3')"  readonly="readonly"   ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctPARModerator3" id="selctPARModerator3" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
    </tr>
    <tr>
      <td class="style1"><p>ข้อบกพร่องที่พบ</p>
          <p>CAR No.(AutoNumber) </p></td>
      <td class="style1">4.
        <label>
          <textarea name="txtCARDescript4" cols="80" rows="5" id="txtCARDescript4" onChange="checkCountNum('CAR','4')" readonly></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctCARModerator4" id="selctCARModerator4" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
      <td width="15%" class="style1"><p>แนวโน้มข้อบกพร่องที่พบ</p>
          <p>PAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle" class="style1">4.
        <label>
          <textarea name="txtPARDescript4" cols="80" rows="5" id="txtPARDescript4"   onchange="checkCountNum('PAR','4')"  readonly="readonly"    ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctPARModerator4" id="selctPARModerator4" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
    </tr>
    <tr>
      <td class="style1"><p>ข้อบกพร่องที่พบ</p>
        <p>CAR No.(AutoNumber) </p></td>
      <td class="style1">5.
        <label>
          <textarea name="txtCARDescript5" cols="80" rows="5" id="txtCARDescript5" onChange="checkCountNum('CAR','5')" readonly></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctCARModerator5" id="selctCARModerator5" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
      <td width="15%" class="style1"><p>แนวโน้มข้อบกพร่องที่พบ</p>
          <p>PAR No.(AutoNumber) </p></td>
      <td width="35%" valign="middle" class="style1">5.
        <label>
          <textarea name="txtPARDescript5" cols="80" rows="5" id="txtPARDescript5"   onchange="checkCountNum('PAR','5')"  readonly="readonly"   ></textarea>
        </label>
        <br /><br />&nbsp;&nbsp;&nbsp;&nbsp;<label>ลำดับชื่อผู้ทำการตรวจ
        <select name="selctPARModerator5" id="selctPARModerator5" disabled="disabled">
          <option value="1">หัวหน้าผู้ตรวจติดตาม</option>
          <option value="2">ผู้ตรวจติดตามคนที่ 1</option>
          <option value="3">ผู้ตรวจติดตามคนที่ 2</option>
          <option value="4">ผู้ตรวจติดตามคนที่ 3</option>
          <option value="5">ผู้ตรวจติดตามคนที่ 4</option>
          <option value="6">ผู้ตรวจติดตามคนที่ 5</option>
        </select>
        </label><br /><br />        </td>
    </tr>
  </table></td>
  </tr>
  <tr>
    <td class="style1">รายละเอียดเพิ่มเติม /<br />
      ข้อคิดเห็นของผู้ตรวจติดตาม</td>
    <td class="style1"><label>
      <textarea name="txtDes" cols="80" rows="5" id="txtDes" class="style1"></textarea>
    </label></td>
  </tr>
  <tr>
    <td>ข้อดี :</td>
    <td valign="middle"><label>
      <textarea name="txtGoodDes" cols="80" rows="5" class="style1" id="txtGoodDes"></textarea>
    </label></td>
  </tr>
  <tr>
    <td>ข้อเสีย :</td>
    <td><label>
      <textarea name="txtBadDes" cols="80" rows="5" class="style1" id="txtBadDes"></textarea>
    </label></td>
  </tr>
  <tr>
    <td><p>&nbsp;</p>      </td>
    <td><label>
      <input name="butSave" type="button" id="butSave" value="บันทึก" onClick="IAuditSave()" style="width:100px; height:35px; font-size:16px; font-style:bold" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button"  value="กลับหน้าแรก" style="width:100px; height:35px; font-size:16px; font-style:bold"  onclick="javascript:{window.open('http://filing.fda.moph.go.th/kmfda/_block/qos','_self');}"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <% if isEmpty(getSave)  = False then %>
      <input name="butPrint" type="button" id="butPrint" value="พิมพ์รายงาน" onClick="IAuditPrint()" style="width:100px; height:35px; font-size:16px; font-style:bold" />
      <%else%>
	  <input type="button" value="พิมพ์รายงาน" onClick="goInternalAuditReport()" style="width:100px; height:35px; font-size:16px; font-style:bold"  />&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="txtREditSOP" id="txtREditSOP" />&nbsp;&nbsp;&nbsp;หมายเหตุ กรุณาใส่รหัสเอกสารคุณภาพที่ต้องการพิมพ์รายงาน
	  <% end if %>
    </label></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><p>&nbsp;</p>      </td>
    <td>&nbsp;</td>
  </tr>
    <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
