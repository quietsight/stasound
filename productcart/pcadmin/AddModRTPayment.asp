<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pageTitle="Modify Existing Payment Option"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 
%>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<% Section="paymntOpt" %>
<!--#include file="AdminHeader.asp"-->
<!--#include file="RTGatewayConstants.asp"-->
<% pageTitle="Add a Real-Time Payment Option" %>
<% 
Dim query, connTemp, rs
Dim i, strErr, isComErr, strComErr
	
sMode=Request.Form("submitMode")
eMode=Request.Form("mode")
iMode=Request.QueryString("mode")
'Delete
If iMode="Del" Then
	pcv_processOrder=request.Form("pcv_processOrder")
	if pcv_processOrder="" then
		pcv_processOrder="0"
	end if
	pcv_setPayStatus=request.Form("pcv_setPayStatus")
	if pcv_setPayStatus="" then
		pcv_setPayStatus="3"
	end if
	idPayment=Request.QueryString("id")
	gwCode= Request.QueryString("gwChoice")
	call openDb()
	If request.QueryString("TYPE")="CC" then
		CCcode=request.queryString("CCCode")
		query="UPDATE CCTypes SET active=0 WHERE CCcode='" & CCcode & "'"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		query="SELECT * FROM CCTypes WHERE active<>0"
		set rs=Server.CreateObject("ADODB.Recordset")  
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		if rs.eof then
			query= "UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",active=0, paymentNickName='' WHERE gwCode=6"
			set rs=Server.CreateObject("ADODB.Recordset")  
			set rs=conntemp.execute(query)
		end if
	Else
		If gwCode="6" then
			query= "UPDATE CCTypes SET active=0 WHERE active<>0"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				strErrorDescription=err.description
				set rs=nothing
				call closeDb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
			end If
			query= "UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",active=0, paymentNickName='' WHERE gwCode=6;"
			set rs=Server.CreateObject("ADODB.Recordset")  
			set rs=conntemp.execute(query)
		End If
		
		If gwCode="7" then
			query= "DELETE FROM payTypes WHERE idPayment="& idPayment
		End If
		
		If gwCode<>"6" AND gwCode<>"7" then
			query= "DELETE FROM payTypes WHERE gwCode="& gwCode
			
			if gwCode="1" then 'Delete authorize check
				querytemp="DELETE FROM payTypes WHERE gwCode=16"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="24" then
				querytemp="DELETE FROM payTypes WHERE gwCode=25" 'Delete TCLink check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="27" then
				querytemp="DELETE FROM payTypes WHERE gwCode=28" 'Delete Netbilling check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="32" then 'Delete CyberSource Echeck
				querytemp="DELETE FROM payTypes WHERE gwCode=62"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="35" then
				querytemp="DELETE FROM payTypes WHERE gwCode=36" 'Delete USAePay check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="37" then
				querytemp="DELETE FROM payTypes WHERE gwCode=38" 'Delete FastCharge check
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="60" then 'Delete DowCom check
				querytemp="DELETE FROM payTypes WHERE gwCode=61"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
			if gwCode="11" then 'Delete authorize check
				querytemp="DELETE FROM payTypes WHERE gwCode=66"
				set rsDelObj=Server.CreateObject("ADODB.Recordset")     
				set rsDelObj=conntemp.execute(querytemp)
				set rsDelObj=nothing
			end if
			
		End if
	End If
	
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	call Closedb()
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	response.redirect "paymentOptions.asp"
End If

if eMode="Edit" Then
	pcv_processOrder=request.Form("pcv_processOrder")
	if pcv_processOrder="" then
		pcv_processOrder="0"
	end if
	pcv_setPayStatus=request.Form("pcv_setPayStatus")
	if pcv_setPayStatus="" then
		pcv_setPayStatus="3"
	end if

	PaymentDesc=request.Form("PaymentDesc")
	idPayment=request.Form("idPayment")
	priceToAddType=request.Form("priceToAddType")
	gwCode=request.form("gwCode")
	If priceToAddType="price" Then
		priceToAdd=replacecomma(Request("priceToAdd"))
		percentageToAdd="0"
		if priceToAdd="" then
			priceToAdd="0"
		end if
	Else
		percentageToAdd=request.Form("percentageToAdd")
		priceToAdd="0"
		if priceToAddType="null" then
			percentageToAdd="0"
		else
			if percentageToAdd="" then
				percentageToAdd="0"
			end if
		end if
	End If
	
	sslUrl=request.Form("sslUrl")
	paymentNickName=replace(request.Form("paymentNickName"),"'","''")

	'//Gateway to edit
	pcv_EditGW=request("addGW")
	call gwCallEdit()
	response.redirect "PaymentOptions.asp"
end if

If sMode <> "" Then
	If sMode="Add Gateway" Then
		dim varCheck
		varCheck=0
		'Start SDBA
		pcv_processOrder=request.Form("pcv_processOrder")
		if pcv_processOrder="" then
			pcv_processOrder="0"
		end if
		pcv_setPayStatus=request.Form("pcv_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA
		'//Gateway to activate
		pcv_AddGW=request("addGW")
		call gwCallAdd()
	end if
	
	If sMode="Add Gateway" and varCheck=1 then
		response.redirect "PaymentOptions.asp"
	else
		response.redirect "AddModRTPayment.asp?msg="&Server.URLEncode("You did not specify a payment option to add. Make sure that you check the box next to the payment option that you wish to add.")
	end if
end if

'check if Centinel has previously been activated.
dim intCentActive
intCentActive=0

err.clear
err.number=0
call openDb()  

query="SELECT pcPay_Cent_Active FROM pcPay_Centinel WHERE pcPay_Cent_ID=1;"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=connTemp.execute(query)
pcPay_Cent_Active=rs("pcPay_Cent_Active")
if pcPay_Cent_Active=1 then
	intCentActive=1
end if

set rs=nothing

if request("mode")="Edit" then
	gwCode=request("gwChoice")
	idPayment=Request.QueryString("id")
	query= "SELECT paymentDesc, priceToAdd, cvv, percentageToAdd, sslUrl, terms, CReq, Cprompt, Cbtob, paymentNickName, pcPayTypes_processOrder, pcPayTypes_setPayStatus FROM payTypes WHERE gwCode= "& gwCode &" AND idPayment= "& idPayment
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	
	if err.number <> 0 then
		strErrorDescription=err.description
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
	end If
	
	if NOT rs.eof then
		paymentDesc=rs("paymentDesc")
		priceToAdd=rs("priceToAdd")
		cvv=0
		percentageToAdd=rs("percentageToAdd")
		sslUrl=rs("sslUrl")
		terms=rs("terms")
		CReq=rs("CReq")
		Cprompt=rs("Cprompt")
		Cbtob=rs("Cbtob")
		paymentNickName=rs("paymentNickName")
		'Start SDBA
		pcv_processOrder=rs("pcPayTypes_processOrder")
		pcv_setPayStatus=rs("pcPayTypes_setPayStatus")
		if pcv_setPayStatus="" then
			pcv_setPayStatus="3"
		end if
		'End SDBA
	End If
	
	set rs=nothing
	if percentageToAdd<>"0" then
		priceToAddType="percentage"
	end if
	if priceToAdd<>"0" then
		priceToAddType="price"
	end if
else
	paymentNickName="Credit Card"
	paymentNickName2="Check"
	percentageToAdd="0"
	priceToAdd="0"
end if


call closedb()
%>
<form name="form1" method="post" action="AddModRTPayment.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr> 
			<td valign="top"> 
				<table border="0" width="100%" cellpadding="2" cellspacing="0">
					<% if not request("gwChoice")="" then 'If gwChoice is empty, hide the rest of the page %>
						<!--#include file="RTGatewayIncludes.asp"-->
						<%'Start SDBA%>
                        <tr>
                            <td class="pcCPspacer"></td>
                        </tr>
                        <% '//Do not display
						if intDoNotApply = 0 Then %>
                        <tr>
                            <th>Order Processing: Order Status and Payment Status</th>
                        </tr>
                        <tr>
                            <td class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td>Process orders when they are placed: <input type="checkbox" class="clearBorder" name="pcv_processOrder" value="1" <%if pcv_processOrder="1" then%>checked<%end if%>>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=301')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                        </tr>
                        <tr> 
                            <td>When orders are placed, set the payment status to:
                                <select name="pcv_setPayStatus">
                                    <option value="3" selected="selected">Default</option>
                                    <option value="0" <%if pcv_setPayStatus="0" then%>selected<%end if%>>Pending</option>
                                    <option value="1" <%if pcv_setPayStatus="1" then%>selected<%end if%>>Authorized</option>
                                    <option value="2" <%if pcv_setPayStatus="2" then%>selected<%end if%>>Paid</option>
                                </select>
                                &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=302')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                        </tr>
						<%'End SDBA%>			
                        <tr>
                            <td class="pcCPspacer"></td>
                        </tr>
                        <tr> 
                            <td align="center">
								<% if request("mode")="Edit" then
                                    strButtonValue="Save Changes"
								%>
                                <input type="hidden" name="submitMode" value="Edit">
                                <%
                                else
                                    strButtonValue="Add New Payment Method"
								%>
                                <input type="hidden" name="submitMode" value="Add Gateway">
                                <%
                                end if
                                %>
                                <input type="submit" value="<%=strButtonValue%>" name="Submit" class="submit2"> 
                                &nbsp;
                                <input type="button" value="Back" onclick="javascript:history.back()">
                            </td>
                        </tr>
                        <% end if %>
                        <tr>
                            <td class="pcCPspacer"></td>
                        </tr>
					<% end if 'gwChoice is empty %>
                </table>
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->
<% Function IsObjInstalled(intClassNum)
	On Error Resume Next
	CONST CLASSBOUND = 2
	Dim objTest, j
	Dim strError
	strError = ""
	for j = 0 to CLASSBOUND
		If Not IsEmpty(strClass(intClassNum, j)) Then 		
			Set objTest = Server.CreateObject(strClass(intClassNum, j))	
			If Err.Number = 0 Then
				Set objTest = Nothing
				If Not isEmpty(strClass(intClassNum, 3)) Then
				strError = ""
				End if
			Else
				If IsObject(objTest) Then Set objTest = Nothing
				If strError = "" Then
					strError = strClass(intClassNum, j)
				Else
					strError = strError & ",<br>" & strClass(intClassNum, j)
				End If
			End If
		Else
			If Not isEmpty(strClass(intClassNum, 3)) Then
				errArray = split(strError, ",<br>", -1)
				if ubound(errArray) = 1 then
					strError = errArray(0)
					strError = strError & "<BR> or " & errArray(1)
				end if 
			End if  				
		End If	
	Next
	IsObjInstalled = strError
End Function
%>
