<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
iddupProduct=request.form("iddup")
iddupAssignment=""
if instr(iddupProduct,"|") then
	dupArray=Split(iddupProduct,"|")
	iddupProduct=dupArray(0)
	iddupAssignment=dupArray(1)
end if

idProduct=request.form("idProduct")
pricingdup=request.form("pricingdup")
updatedup=request.form("updatedup")

dim strSQL, conntemp, rstemp
strSQL="SELECT DISTINCT options_optionsGroups.idProduct, options_optionsGroups.idOptionGroup, options_optionsGroups.idOption, options_optionsGroups.price, options_optionsGroups.Wprice,options_optionsGroups.sortOrder,options_optionsGroups.InActive FROM options_optionsGroups "
strSQL = strSQL & "WHERE (( (options_optionsGroups.idProduct)="&iddupProduct&" "
if len(iddupAssignment)>0 then
	strSQL = strSQL & "AND (options_optionsGroups.idOptionGroup)="&iddupAssignment&" " 
end if
strSQL = strSQL & ")) Order By idOptionGroup;"
'response.write strSQL
'response.end
call openDB()
set rs=Server.CreateObject("ADODB.Recordset")
set rs=conntemp.execute(strSQL)

repeatcnt = 0
Drepeatcnt = 0
contgo=0
xCounter = 0
cntG=0
pcv_strMsg = ""
cnt=0

'//////////////////////////////////////////////////////////////
'// START: LOOP Through every single Attribute
'//////////////////////////////////////////////////////////////
do until rs.eof

	'// Get all the Attribute Data
	intidOptionGroup=rs("idOptionGroup")
	intidOption=rs("idOption")
	intprice=rs("price")
	intWprice=rs("Wprice")
	intSortOrder=rs("sortOrder")
	intInActive=rs("InActive")
	
	if xCounter = 0 then
		OintidOptionGroup = intidOptionGroup
	end if
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  START: Reporting Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
	query="SELECT * FROM OptionsGroups WHERE idOptionGroup="&intidOptionGroup
	set rstemp=conntemp.execute(query)
	OptionGroupDesc=rstemp("OptionGroupDesc")
	set rstemp=nothing
	If Session("pcAdmin" & trim(OptionGroupDesc)) = "" AND xCounter>0 Then
		Session("pcAdmin" & trim(OptionGroupDesc)) = trim(OptionGroupDesc)
		if OintidOptionGroup<>intidOptionGroup then
			Drepeatcnt=0
			repeatcnt=0
			cnt=0
			pcv_strCleanUpSessions = pcv_strCleanUpSessions & Session("pcAdmin" & OptionGroupDesc) & ","
			pcv_strMsgMaster = pcv_strMsgMaster & pcv_strMsg
		end if	
	End if	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  END: Reporting Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  START: Attribute Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// Check if this Attribute ALREADY exists in database before adding				
	strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&" AND idOption="&intidOption&";"
	set rstemp=conntemp.execute(strSQL)
	if rstemp.eof then	
		'// ADD (Replicate)
		if pricingdup="1" then
			strSQL="INSERT INTO options_optionsGroups (idProduct, idOptionGroup, idOption, price, Wprice,sortOrder,InActive) VALUES ("&idProduct&","&intidOptionGroup&","&intidOption&","&intprice&","&intWprice&"," & intSortOrder & "," & intInActive & ");"
		else
			strSQL="INSERT INTO options_optionsGroups (idProduct, idOptionGroup, idOption, price, Wprice,sortOrder,InActive) VALUES ("&idProduct&","&intidOptionGroup&","&intidOption&",0,0"&"," & intSortOrder & "," & intInActive & ");"
		end if	
		set rsSetDup=Server.CreateObject("ADODB.Recordset")
		set rsSetDup=conntemp.execute(strSQL)
		'// Set the update flag
		contgo=1
		cnt=cnt+1		
	else
		'// UPDATE (if over-write was selected)
		if updatedup="1" then
			strSQL="UPDATE options_optionsGroups SET price="&intprice&",Wprice="&intWprice&",sortOrder="&intSortOrder&",InActive="&intInActive&"  WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&" AND idOption="&intidOption&";"
			set rsSetDup=Server.CreateObject("ADODB.Recordset")
			set rsSetDup=conntemp.execute(strSQL)
		end if						
	end if
	set rsSetDup=nothing
	set rstemp=nothing
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  END: Attribute Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  START: Required Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	intdupRequired=0
	'// Get Required Flag from Dup Product
	strSQL="SELECT pcProdOpt_Required,pcProdOpt_Order FROM pcProductsOptions WHERE idproduct="& iddupProduct &" AND idoptionGroup="&intidOptionGroup&";"
	set rstemp=conntemp.execute(strSQL)
	if not rstemp.eof then	
		intdupRequired=rstemp("pcProdOpt_Required")
		intSortOrder=rstemp("pcProdOpt_Order")
	end if

	'// Check if this Attribute ALREADY exists in database before adding				
	strSQL="SELECT * FROM pcProductsOptions WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&";"
	set rstemp=conntemp.execute(strSQL)
	if rstemp.eof then	
		'// ADD (Replicate)
		strSQL="INSERT INTO pcProductsOptions (idProduct, idOptionGroup, pcProdOpt_Required, pcProdOpt_Order) VALUES ("&idProduct&","&intidOptionGroup&","&intdupRequired&","&intSortOrder&");"
		set rsSetDup=Server.CreateObject("ADODB.Recordset")
		set rsSetDup=conntemp.execute(strSQL)
		'// Set the update flag
		'contgo=1
		'cnt=cnt+1		
	else
		'// UPDATE (if over-write was selected)
		if updatedup="1" then
			strSQL="UPDATE pcProductsOptions  SET pcProdOpt_Required="&intdupRequired&",pcProdOpt_Order="&intSortOrder&"  WHERE idproduct="& idProduct &" AND idoptionGroup="&intidOptionGroup&";"
			set rsSetDup=Server.CreateObject("ADODB.Recordset")
			set rsSetDup=conntemp.execute(strSQL)
		end if						
	end if
	set rsSetDup=nothing
	set rstemp=nothing
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  END: Required Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  START: Product Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'// If at least one Attribute was added for this product check that there is a relationship for the Group
	if contgo=1 then		
		'// If this is a new option group, then we need to add the relation
		strSQL="SELECT idOptionGroup, idproduct FROM pcProductsOptions WHERE idproduct="& idProduct &" AND idOptionGroup="& intidOptionGroup &" "
		set rsOptionCheck=conntemp.execute(strSQL)	
		if rsOptionCheck.eof then
			strSQL="INSERT INTO pcProductsOptions (idproduct, idOptionGroup, pcProdOpt_Required, pcProdOpt_Order) VALUES (" & idProduct &", " & intidOptionGroup & ", 0, 0)"
			set rstemp=conntemp.execute(strSQL)
			'// if the option group is new keep count
			cntG=cntG+1
		end if
		set rsOptionCheck = nothing		
	end if
	
	'// If and Attribute was NOT added for this product keep count
	if contgo=0 and updatedup="0" then
		repeatcnt=repeatcnt+1
	end if
	if contgo=0 and updatedup="1" then
		Drepeatcnt=Drepeatcnt+1
	end if
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'  END: Product Level Tasks
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	pcv_strMsg = ""
	
	If cnt>0 then
	pcv_strMsg = cnt &" option attributes were copied from the Option Group: <b>"& OptionGroupDesc &"</b>.  <br />"
	end if
	
	If repeatcnt>0 then 
		pcv_strMsg = pcv_strMsg & repeatcnt &" attributes were skipped from the Option Group: <b>"& OptionGroupDesc &"</b>.  <br />"
	end if
	
	If Drepeatcnt>0 then 
		pcv_strMsg = pcv_strMsg & Drepeatcnt &" attributes were over-written in the Option Group: <b>"& OptionGroupDesc &"</b>.  <br />"
	end if		
	
xCounter = xCounter + 1
rs.movenext
loop	
set rs=nothing

pcv_strMsgMaster = pcv_strMsgMaster & pcv_strMsg
'response.write "<hr>" & pcv_strMsgMaster & "<hr>" & xCounter

' Clean Up the Sessions
pcv_strCleanUpSessions = split(pcv_strCleanUpSessions, ",")
for x = lbound(pcv_strCleanUpSessions) to ubound(pcv_strCleanUpSessions)
	Session("pcAdmin" & pcv_strCleanUpSessions(x)) = ""
next

If cntG>0 then 
	'// If we have added a new group display that info
	strMsg = cntG &" Option Groups were copied. <br/><br/> " & pcv_strMsgMaster
else
	'// If we only updated existing groups
	strMsg = pcv_strMsgMaster
end if

call updPrdEditedDate(idProduct)

call closedb()
response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode(strMsg)&"&idproduct="&idProduct
%>