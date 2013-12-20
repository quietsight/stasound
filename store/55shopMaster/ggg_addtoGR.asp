<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=7%>
<% pageTitle="Add Products to the Registry" %>
<% section="mngAcc" %>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%

pcv_IdCustomer=getUserInput(request("idCustomer"),10)
gIDEvent=getUserInput(request("IDEvent"),10)

if not validNum(pcv_IdCustomer) or not validNum(gIDEvent) then
	response.redirect "menu.asp"
end if

dim conntemp,query,rs
call openDb()

if gIDEvent<>"" then
	query="select pcEv_IDEvent from pcEvents where pcEv_IDCustomer=" & pcv_IdCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "ggg_manageGRs.asp?idcustomer=" & pcv_IdCustomer
	end if
	set rstemp=nothing
end if


'*****************************************************************************************************
' START: Save Cart to Registry
'*****************************************************************************************************
if request("action")="add" then
	pcListArray=request("prdlist")
	pcA=split(pcListArray,",")
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Loop Through the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	for f=lbound(pcA) to ubound(pcA)
		if trim(pcA(f))<>"" then
			gIDProduct=pcA(f)
			gQty=1
			pcv_strSelectedOptions=""
			gxdetails=""
			gIDConfig="0"
			
			query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_Qty, pcEP_OptionsArray, pcEP_xdetails,pcEP_IDConfig) values (" & gIDEvent & "," & gIDProduct & "," & gQty & ",'" & pcv_strSelectedOptions & "','" & gxdetails & "'," & gIDConfig & ")"
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rstemp=nothing
		end if
	next
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Loop Through the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	set rstemp=nothing
	call closeDb()
	response.redirect "ggg_GRDetails.asp?IDEvent=" & gIDEvent & "&idcustomer=" & pcv_IdCustomer

end if
'*****************************************************************************************************
' END: Save Cart to Registry
'*****************************************************************************************************

call closedb()	

%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Add Products to the Registry"
				src_FormTips1="Use the following filters to look for products in your store that you would like to add to the selected Gift Registry."
				src_FormTips2="Select one or more products that you would like to add to the selected Gift Registry."
				src_IncNormal=1
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ggg_addtoGR.asp?idcustomer=" & pcv_IDCustomer & "&IDEvent=" & gIDEvent
				src_ToPage="ggg_addtoGR.asp?idcustomer=" & pcv_IDCustomer & "&IDEvent=" & gIDEvent & "&action=add"
				src_Button1=" Search "
				src_Button2=" Add to the Registry "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=0
				'session("srcprd_from")=""
				'session("srcprd_where")=" AND (products.idproduct NOT IN (SELECT pcEP_IDProduct FROM pcEvProducts WHERE pcEP_IDEvent=" & gIDEvent & ")) "
			%>
			<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->