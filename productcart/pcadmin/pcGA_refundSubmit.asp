<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Adjust Google Analytics Statistics" %>
<% Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Session.LCID = 1033 %>

<%
'// GOOGLE ANALYTICS
'// LOAD Google Analytics code

'// COPY and PASTE your tracking code 'as is' from your Google Analytics account
'// You can find the code on: Analytics Settings > Profile Settings > Tracking Status 
%>

<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-XXXXX-X']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>

<%
'// DO NOT edit the code after this line

'// STEP 3 - Prepare data and send to Google Analytics
'// Step 1 (pick order) and 2 (specify what to refund) are performed on pcGA_refund.asp

		dim query, conntemp, rstemp, paffiliateName, paffiliateCompany, ptaxAmount, ptaxDetails, pshipmentDetails, ptotal, pGetItems, iCount, pReturnQuantity, pcGAtransaction, pcGAtransactionLog, pcGAtransactionItems, pcGAtransactionItemsLog

		'// A - START - Prepare ORDER info
		'// Get order information from form

				pOID=request.Form("idOrder")
					if not validNum(pOID) then
						response.Redirect("pcGA_refund.asp")
					end if

				'// Amount to refund
				ptotal=trim(request.Form("total"))
				
				'// Tax amount to refund
				ptaxAmount=trim(request.Form("taxes"))
				
				'// Shipping amount to refund
				pTotalShipping=trim(request.Form("shipping"))
				
				'// Validate entries
				'// We are not validating for negative numbers to allow the admin to use this feature to
				'// increase the total for that order (e.g. order was edited adding on to the existing total)
				if not isNumeric(pTotal) or not isNumeric(ptaxAmount) or not isNumeric(pTotalShipping) then
					msg="Check the amounts entered under 'General Information' and make sure that they are all valid numbers."
					response.Redirect "pcGA_refund.asp?idOrder="&pOID&"&msg="&msg
				end if	
				
		'// Get the rest of the order information from the database.
				
				call openDb()
				query="SELECT city, state, stateCode, CountryCode, idAffiliate FROM orders WHERE idOrder=" & pOID
					set rs=Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)
												
						'// Gather affiliate information
						pidAffiliate=rs("idaffiliate")
							If pidaffiliate>"1" then
								query="SELECT affiliateName, affiliateCompany FROM affiliates WHERE idAffiliate =" & pidAffiliate
								Set rsTemp=Server.CreateObject("ADODB.Recordset")
								Set rsTemp=connTemp.execute(query)
								paffiliateName = rsTemp("affiliateName")
								paffiliateCompany = rsTemp("affiliateCompany")
									if trim(paffiliateCompany)<>"" then
										paffiliateName = paffiliateName & "(" & paffiliateCompany & ")"
									end if
								paffiliateName = replace(paffiliateName,"|","-")
								Set rsTemp = nothing
							else
								paffiliateName = "N/A"
							end if
								
						'// Gather order location information
						pcity=rs("city")
							pcity = replace(pcity,"|","-")
						pstate=rs("state")
						pstateCode=rs("stateCode")
							if trim(pstateCode)="" then
								pstateCode=pstate
							end if
						pCountryCode=rs("CountryCode")
						
						set rs = nothing
						call closeDb()
						
						'// Transaction line example per Google Analytics documentation
						pcGAtransaction = 					"_gaq.push(['_addTrans', " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & scpre+int(pOID) & "',           	// order ID - required " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & paffiliateName & "',  				// affiliation or store name " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & ptotal & "',          				// total - required " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & ptaxAmount & "',           			// tax " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & pTotalShipping & "',              	// shipping " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & pcity & "',       					// city " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & pstateCode & "',     					// state or province " & VbCrLf
						pcGAtransaction = pcGAtransaction & "'" & pCountryCode & "'             		// country " & VbCrLf
						pcGAtransaction = pcGAtransaction & "]); " & VbCrLf
						pcGAtransaction = pcGAtransaction & VbCrLf
						
						pcGAtransactionLog = 					"_gaq.push(['_addTrans', " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & scpre+int(pOID) & "',           	// order ID - required " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & paffiliateName & "',  				// affiliation or store name " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & ptotal & "',          				// total - required " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & ptaxAmount & "',           			// tax " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & pTotalShipping & "',              	// shipping " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & pcity & "',       					// city " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & pstateCode & "',     					// state or province " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "'" & pCountryCode & "'             		// country " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & "]); " & VbCrLf
						pcGAtransactionLog = pcGAtransactionLog & VbCrLf
																	
		'// A - END - Prepare ORDER info

		'// B - START - Prepare ITEM info
												
					'// Gather item information from database & form
					call openDB()
					query="SELECT ProductsOrdered.idProduct, ProductsOrdered.unitPrice, products.description, products.sku FROM ProductsOrdered, products WHERE ProductsOrdered.idProduct=products.idProduct AND ProductsOrdered.idOrder=" & pOID
					set rs=Server.CreateObject("ADODB.Recordset")
					set rs=conntemp.execute(query)												
						
						pcGAtransactionItems = ""
						pcGAtransactionItemsLog = ""
		
						Do While Not rs.eof
							
							pIdProduct = rs("idProduct")			
							pSKU = rs("sku")
								pSKU = replace(pSKU,"|","-")
							pName = rs("description")
								pName = replace(pName,"|","-")
							pUnitPrice = rs("unitPrice")
							
							'// Get the quantity to be returned from the form
							pReturnQuantity = trim(request.form("quantity"&pIdProduct))
								if pReturnQuantity="" or pReturnQuantity="0" then
									pReturnQuantity=0
								end if
							
								'// Validate entry
								'// We are not validating for a negative number to allow the admin to use this feature to
								'// increase the products purchased (e.g. added 1 unit via Edit Order feature)
								if not isNumeric(pReturnQuantity) then
									msg="Enter a valid number of units for each of the products listed below."
									response.Redirect "pcGA_refund.asp?idOrder="&pOID&"&msg="&msg
								end if	

							
								'// Find category information
								query="SELECT idCategory FROM categories_products WHERE idProduct ="& pIdProduct
								set rsTemp=server.CreateObject("ADODB.RecordSet")
								set rsTemp=conntemp.execute(query)
								if not rsTemp.eof then
									idCategory=rsTemp("idCategory")
									query="SELECT categoryDesc FROM categories WHERE idCategory =" & idCategory
									set rsTemp=conntemp.execute(query)
									if err.number <> 0 then
										set rsTemp=nothing
										pCategory = "NA"
									end If
									if rsTemp.eof then
										set rsTemp=nothing
										pCategory = "NA"
									end if
									pCategory = rsTemp("categoryDesc")
								else
									pCategory = "NA"
								end if
								pCategory = replace(pCategory,"|","-")
								set rsTemp=nothing
						
							'// Item line example per Google Analytics documentation
							if pReturnQuantity<>0 then

								pcGAtransactionItems = pcGAtransactionItems & "_gaq.push(['_addItem', " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & scpre+int(pOID) & "', 	// order ID - required " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & pSKU & "',           			// SKU/code " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & pName & "',        			// product name " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & pCategory & "',   			// category or variation " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & pUnitPrice & "',          	// unit price - required " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "  '" & pQuantity & "'               	// quantity - required " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & "]); " & VbCrLf
								pcGAtransactionItems = pcGAtransactionItems & VbCrLf							

								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "_gaq.push(['_addItem', " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & scpre+int(pOID) & "', 	// order ID - required " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & pSKU & "',           			// SKU/code " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & pName & "',        			// product name " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & pCategory & "',   			// category or variation " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & pUnitPrice & "',          	// unit price - required " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "  '" & pQuantity & "'               	// quantity - required " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & "]); " & VbCrLf
								pcGAtransactionItemsLog = pcGAtransactionItemsLog & VbCrLf									

							end if
						
						rs.movenext
						loop
						set rs = nothing
						call closeDb()						
				
		'// B - END - Prepare ITEM info
				
				
				
		'// Write the hidden form
		%>
		<script type="text/javascript">
		<%=pcGAtransaction%>
		<%=pcGAtransactionItems%>
		_gaq.push(['_trackTrans']); 
		</script> 
		

		
<%
'// *****************************************
'// Write to LOG file
'// *****************************************

FileName="GAlogs\gaLog.txt"
Contents=(scpre+int(pOID)) & "," & date & ",""" & pcGAtransactionLog & """,""" & pcGAtransactionItemsLog & """" & VbCrLf
set oFs = server.createobject("Scripting.FileSystemObject")
set oTextFile = oFs.OpenTextFile(Server.mappath(FileName), 8, True)
oTextFile.Write Contents
oTextFile.Close
set oTextFile = nothing
set oFS = nothing


'// *****************************************
'// Show information sent to Google Analytics
'// *****************************************
%>

<table class="pcCPcontent">
	<tr>
		<td>
			<div>The following information was sent to Google Analytics:</div>
			<div class="pcCPsectionTitle">Order Information</div>
			<div style="padding: 15px;"><textarea cols="100" rows="2"><%=pcGAtransaction%></textarea></div>
			<div class="pcCPsectionTitle">Item Information</div>
			<div style="padding: 15px;"><textarea cols="100" rows="6"><%=pcGAtransactionItems%></textarea></div>
			<div style="padding: 15px;">The information was formatted according to <a href="http://code.google.com/apis/analytics/docs/tracking/gaTrackingEcommerce.html" target="_blank">these requirements</a>.</div>
			<div style="padding-bottom: 15px;"><strong>Reporting Delays</strong>: please note that ecommerce transactions (orders and adjustments) typically do not appear in your Google Analytics reports until the following day. Therefore, <u>you should not expect your reports to immediately reflect the adjustments</u> that you just posted.</div>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="center"><a href="pcGA_refund.asp">Post another adjustment</a> | <a href="start.asp">Return to the Start page</a></td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->