<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/secureadminfolder.asp"--> 
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="CustLIv.asp"-->
<!--#include file="DBsv.asp"-->
<!--#include file="header.asp"-->

<%
Dim pIdOrder, rs, pcIntTempCustID
'==================================================
'= START Check successfull request and show thank you
'==================================================

pShowThankYou=getUserInput(request("thankYou"),0)
pIdOrder=getUserInput(request("idOrder"),0)

if not validNum(pIdOrder) then
   response.redirect "msg.asp?message=35" 
end if

	'// SECURITY CHECK
	'// Check that order belongs to correct customer
		call openDb()
		query="SELECT orders.idcustomer FROM orders WHERE orders.idOrder=" &pIdOrder
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
		if rs.EOF then
			set rs=nothing
			call closedb()
			response.redirect "msg.asp?message=35" 
		end if
		
		pcIntTempCustID=rs("idcustomer")
		set rs=nothing
		call closedb()
		
		if int(pcIntTempCustID)<>int(session("idCustomer")) then
			response.redirect "msg.asp?message=11" 
		end if
	'// END SECURITY CHECK

IF pShowThankYou <> "" THEN

	' Prepare notification email for store administrator
	rmaSubject="Return Authorization Request for order #"&(int(pIdOrder)+scpre)
	rmaBody=""
	rmaBody=rmaBody&"Return Authorization Request Notification"&VBcrlf&VBcrlf
	rmaBody=rmaBody&"Order #: "&(int(pIdOrder)+scpre)&VBcrlf&VBcrlf
	rmaBody=rmaBody&"A customer submitted a request for a Return Manufacturer Authorization (RMA). You may approve or deny the customer's request to return one or more of the products ordered."&VBcrlf&VBcrlf
	rmaBody=rmaBody&"To view the RMA request and decide whether it should be approved or not, log into the Control Panel and view order details for order # "&(int(pIdOrder)+scpre)&". Click on the link below to load that page:"&VBcrlf&VBcrlf
	dim tempURL
	tempURL=scStoreURL&"/"&scPcFolder&"/"&scAdminFolderName&"/ordDetails.asp?"
	tempURL=replace(tempURL,"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	rmaBody=rmaBody&tempURL&"id="&(int(pIdOrder))&VBcrlf&VBcrlf
	rmaBody=rmaBody&"Please refer to the ProductCart User Guide for more information about processing Return Authorizations."&VBcrlf&VBcrlf
	call sendmail (scCompanyName, scEmail, scFrmEmail, rmaSubject, rmaBody)
	
	'Show success message
	%>
	<div id="pcMain">
		<table class="pcMainTable">
			<tr>
				<td><%response.write dictLanguage.Item(Session("language")&"_rma_8")%></td>
			</tr>
		</table>
	</div>
<%
END IF
'==================================================
'= END Check successfull request and show thank you
'==================================================

IF pShowThankYou = "" THEN ' Don't show the page if the thank you message has been shown
'==================================================
'= Check RMA form submission and process request
'==================================================
	IF request.form("action")<>"" THEN ' Start form submission statement
	
			call openDb()
	
			pRmaReturnReason=getUserInput(request("rmaReturnReason"),0)						
			pRmaReturnReason=replace(pRmaReturnReason,"'","''")			
			pIdProduct=getUserInput(request("idProduct"),0)			
			pRMADate=Now()
			if SQL_Format="1" then
				pRMADate=Day(pRMADate)&"/"&Month(pRMADate)&"/"&Year(pRMADate)
			else
				pRMADate=Month(pRMADate)&"/"&Day(pRMADate)&"/"&Year(pRMADate)
			end if	
			pRMADate=pRMADate & " " & Time()
			if scDB="SQL" then
				query="INSERT INTO PCReturns (rmaNumber,rmaReturnReason,rmaDateTime,idOrder,rmaIdProducts,rmaApproved) VALUES ('" &pRmaNumber& "','" &pRmaReturnReason& "','"&pRMADate&"',"&pIdOrder&",'"&pIdProduct&"',0)"
			else
				query="INSERT INTO PCReturns (rmaNumber,rmaReturnReason,rmaDateTime,idOrder,rmaIdProducts,rmaApproved) VALUES ('" &pRmaNumber& "','" &pRmaReturnReason& "',#"&pRMADate&"#,"&pIdOrder&",'"&pIdProduct&"',0)"
			end if
			set rsTemp=Server.CreateObject("ADODB.Recordset")
			set rsTemp=connTemp.execute(query) 
			
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsTemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
			call closeDb()
	
			response.redirect "rmaIndex.asp?thankyou=1&idOrder="&pIdOrder
			
	ELSE
	
	'===================================================
	'= Form has NOT been submitted: display it
	'===================================================
	
	
	call openDb()
	
	query="SELECT ProductsOrdered.idProduct, ProductsOrdered.idOrder, products.description, products.sku, products.idProduct, orders.idOrder FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder & " AND orders.idcustomer=" & session("idCustomer")
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
			
		if rstemp.EOF then
			set rs=nothing
			call closedb()
			response.redirect "msg.asp?message=35" 
		end if
	%>
	
	<script language="JavaScript">
	<!--
	function Form1_Validator(theForm)
	{
			// require that at least one checkbox be checked
			if (typeof theForm.idProduct.length != 'undefined') {
				var checkSelected = false;
				for (i = 0;  i < theForm.idProduct.length;  i++)
				{
				if (theForm.idProduct[i].checked)
				checkSelected = true;
				}
				if (!checkSelected)
				{
				alert("<%response.write dictLanguage.Item(Session("language")&"_rma_20")%>");
				return (false);
				}
			} else {
				if (!theForm.idProduct.checked)
				{
				alert("<%response.write dictLanguage.Item(Session("language")&"_rma_20")%>");
				return (false);
				}
		}
		if (theForm.rmaReturnReason.value == "")
			{
				 alert("<%response.write dictLanguage.Item(Session("language")&"_rma_21")%>");
					theForm.rmaReturnReason.focus();
					return (false);
		}
	
	return (true);
	}
	//-->
	</script>
	<%
		if request.form("Submit")<>"" then
			rmaReturnReason=request.form("rmaReturnReason")
			Session("rmaReturnReason")=rmaReturnReason
		end if
	%>
	<div id="pcMain">
		<form method="POST" action="rmaindex.asp" name="orderform" onsubmit="return Form1_Validator(this)" class="pcForms">
			<input type="hidden" name="idCustomer" value="<%session("idCustomer")%>">
			<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
			<input type="hidden" name="action" value="1">
			<table class="pcMainTable">
				<tr>
					<td><%response.write dictLanguage.Item(Session("language")&"_rma_1")%></td>
				</tr>
				<tr>
					<td class="pcSpacer"></td>
				</tr>
				<tr>
					<td><p><%response.write dictLanguage.Item(Session("language")&"_rma_11")%></p></td>
				</tr>
				<tr>
				<td>
					<table class="pcShowProducts">
						<% 
						While Not rsTemp.EOF
						pIdProduct=rstemp("idProduct") 
						pSku=rstemp("sku")
						pDescription=rstemp("description")
						%>
							<tr>
								<td width="5%"><input name="idProduct" type="checkbox" id="idProduct" value="<% =pIdProduct %>" class="clearBorder"></td>
								<td width="95%" align="left"><% =psku %> - <% =pDescription %></td>
							</tr>
						<%
						rsTemp.MoveNext
						Wend
						%>
					</table>
				</td>
			</tr>
			<tr>
				<td align="left" valign="top">		
					<table class="pcShowContent">
						<tr>
							<td width="29%" align="right">
								<%response.write dictLanguage.Item(Session("language")&"_rma_2")%>
							</td>
								<td width="71%"><%=(int(pIdOrder)+scpre)%>
							</td>
						</tr>
						<tr>    
							<td align="right" valign="top">
								<%response.write dictLanguage.Item(Session("language")&"_rma_7")%>
							</td>
							<td>
								<textarea rows="5" cols="30" name="rmaReturnReason" value="<%session("rmaReturnReason")%>"><%session("rmaReturnReason")%></textarea>
							</td>
						</tr>
						<tr>
							<td align="right">&nbsp;</td>
							<td>
							<a href="javascript:history.go(-1)"><img src="<%=rslayout("back")%>" border=0></a>&nbsp;
							<input type="image" src="<%=rslayout("submit")%>" border="0" name="Submit" id="submit"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>
	</div>
	<%	
		rsTemp.close
		Set rsTemp = nothing    
		Session("rmaReturnReason")=""
		
	END IF ' End form submission statement
END IF 'Don't show the page if the thank you message has been shown
%>
<!--#include file="footer.asp"-->