<%
if pcv_WPPSpecialFeatures=1 AND (pcgwTransId<>"" AND isNULL(pcgwTransId)=False) then
%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">PayPal</th>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>	
	<%
	'// Query our PayPal Settings
	objPayPalClass.pcs_SetAllVariables()						
	
	'---------------------------------------------------------------------------
	' Construct the parameter string that describes the PayPal payment the varialbes 
	' were set in the web form, and the resulting string is stored in nvpstr
	'
	' Note: Make sure you set the class obj "objPayPalClass" at the top of this page.
	'---------------------------------------------------------------------------
	nvpstr="" '// clear 
	objPayPalClass.AddNVP "TRANSACTIONID", pcgwTransId
	
	'--------------------------------------------------------------------------- 
	' Make the call to PayPal to set the Express Checkout token
	' If the API call succeded, then redirect the buyer to PayPal
	' to begin to authorize payment.  If an error occurred, show the
	' resulting errors
	'---------------------------------------------------------------------------
	Set resArray = objPayPalClass.hash_call("gettransactionDetails",nvpstr)
	Set Session("nvpResArray")=resArray
	ack = UCase(resArray("ACK"))	
	
	pcv_strHideAllPayPal = 0 '// if there is an error hide all the PayPal information
	
	if err.number <> 0 then									
		'// PayPal Error Handler Include: Returns an User Friendly Error Message as the string "pcv_PayPalErrMessage"
		Dim pcv_PayPalErrMessage
		%><!--#include file="../includes/pcPayPalErrors.asp"--><%	
		pcv_strHideAllPayPal = 1														
	end if
	
	If ack="SUCCESS" Then
		pcv_strPayerEmail = resArray("EMAIL")
		pcv_strPayerID = resArray("PAYERID")
		pcv_strTransactionID = resArray("TRANSACTIONID")
		pcv_strAmount = resArray("AMT")
		pcv_strPayerStatus = resArray("PAYERSTATUS")
		pcv_strAddressStatus = resArray("ADDRESSSTATUS")
		'// pcv_strCurrentPPStatus = resArray("PAYMENTSTATUS") '// Broken PayPal Bug
		pcv_strPendingReason = resArray("PENDINGREASON")
	Else								
		objPayPalClass.GenerateErrorReport()
		%>
		<tr>
			<td colspan="2" class="pcCPnotes">
			<div><strong>Alert: We could not display your PayPal Details Section due to the following errors:  </strong></div>
			<% if len(pcPay_PayPal_Signature)>0 then %>
				<%=pcv_PayPalErrMessage%>
                <br />
                <div>Please check your API Signatures were entered correctly, and that you are not using live credentials in test mode (or visa versa)</div>
            <% else %>
            	API Credentials are required for direct payments and post-checkout operations.  When Express Checkout is activated with an email address post-checkout operations are disabled.  To enable post-checkout operations you must obtain API credentials and save them on the payment details page.  Please <a href="http://wiki.earlyimpact.com/productcart/activating_ec_ab#direct_payment_post_processing" target="_blank">review the documentation</a> for help obtaining API Credentials.
            <% end if %>            
			</td>
		</tr>
		<%
		pcv_strHideAllPayPal = 1
	End If
	
	
	If pcv_strHideAllPayPal = 0 Then '// Hide all PayPal Details and display the error.
		
		pcv_strCurrentPPStatus = GetLivePaymentStatus(pcgwTransId) '// Get the "LIVE" Status
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: PayPal Payment Status Conflict Management Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Select Case pcv_strCurrentPPStatus
			Case "Canceled-Reversal"										
				If pcv_PaymentStatus<>"0" Then
					'// Update Payment Status - Pending
					query="UPDATE Orders SET pcOrd_PaymentStatus=0 WHERE idorder=" & qry_ID & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					Set rs=conntemp.execute(query)
					set rs=nothing
					if err.number<>0 then
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails PayPal Payment Status CM: "& err.description) 
					end if
					response.Redirect("Orddetails.asp?id="&qry_ID)
				End if	
				
			Case "Completed"
				'// We can not make assumption regarding the order and payment status for "Completed".
			
			Case "Denied"
				'// will not be recorded in ProductCart
				
			Case "Expired"										
				'// will not be recorded in ProductCart	
				
			Case "Failed"
				'// will not be recorded in ProductCart
			
			Case "In-Progress"
				'// We can not make assumption regarding the order and payment status for "In-Progress".
			
			Case "Partially-Refunded"
				If pcv_PaymentStatus<>"2" Then
					'// Update Payment Status - Paid
					query="UPDATE Orders SET pcOrd_PaymentStatus=2 WHERE idorder=" & qry_ID & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					Set rs=conntemp.execute(query)
					set rs=nothing
					if err.number<>0 then
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails PayPal Payment Status CM: "& err.description) 
					end if
					response.Redirect("Orddetails.asp?id="&qry_ID)
				End if	
			
			Case "Pending"
				'// We can not make assumption regarding the order and payment status for "Pending".
				
			Case "Processed"
				'// We can not make assumption regarding the order and payment status for "Processed".
				
			Case "Refunded"
				If pcv_PaymentStatus<>"6" Then
					'// Update Payment Status - Refunded
					query="UPDATE Orders SET pcOrd_PaymentStatus=6 WHERE idorder=" & qry_ID & ";"
					set rs=server.CreateObject("ADODB.RecordSet")
					Set rs=conntemp.execute(query)
					set rs=nothing
					if err.number<>0 then
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in OrdDetails PayPal Payment Status CM: "& err.description) 
					end if
					response.Redirect("Orddetails.asp?id="&qry_ID)
				End if
				
			Case "Reversed"		
				'// complete		
									
			Case "..." '// "Voided"	

						
		End Select
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: PayPal Payment Status Conflict Management Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: PayPal Payment Customer Notes
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Set Transaction Specific Statuses
		Select Case pcv_PaymentStatus
			Case "6": pcv_PPStatusNote="You have submitted a Refund transaction. "
			Case "8": pcv_PPStatusNote="You have submitted a Void transaction. "									
		End Select
		'// Set "User Friendly" notes to that are specific to that status.
		Select Case pcv_strCurrentPPStatus
			Case "Canceled-Reversal": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Completed": pcv_PPStatusNote=pcv_PPStatusNote&"The transaction has completed successfully."
			Case "Denied": pcv_PPStatusNote=pcv_PPStatusNote&"This transaction was denied.  " '// This case should never show in ProducatCart.
			Case "Expired": pcv_PPStatusNote=pcv_PPStatusNote&"" '// This message is defined below in the Reauthorization Checks.
			Case "Failed": pcv_PPStatusNote=pcv_PPStatusNote&"This transaction failed. " '// This case should never show in ProducatCart.
			Case "In-Progress": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Partially-Refunded": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Pending"					
				if pcv_strPendingReason="echeck" then
						pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: eCheck - The payment is pending because it was made by an eCheck, which has not yet cleared"
				elseif pcv_strPendingReason="authorization" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Authorization - You can Capture funds with the action button below."
				elseif pcv_strPendingReason="multi_currency" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Multi Currency - You do not have a balance in the currency sent, and you do not have your Payment Receiving Preferences set to automatically convert and accept this payment. You must manually accept or deny this payment from your Account Overview"
				elseif pcv_strPendingReason="intl" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: International - The payment is pending because you, the merchant, hold a non-U.S. account and do not have a withdrawal mechanism. You must manually accept or deny this payment from your Account Overview"
				elseif pcv_strPendingReason="verify" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Not Verified - The payment is pending because you, the merchant, are not yet Verified. You must Verify your account before you can accept this payment"
				elseif pcv_strPendingReason="address" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Address Not Confirmed - The payment is pending because your customer did not include a confirmed shipping address and you, the merchant, have your Payment Receiving Preferences set such that you want to manually accept or deny each of these payments. To change your preference, go to the 'Selling Preferences' section of your Profile"
				elseif pcv_strPendingReason="upgrade" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Account Upgrade Required - The payment is pending because it was made via credit card and you, the merchant, must upgrade your account to Premier or Business status in order to receive the funds. It may also mean that you have reached the monthly limit for transactions on your account"
				elseif pcv_strPendingReason="unilateral" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: E-mail address not registered - The payment is pending because it was made to an email address that is not yet registered or confirmed with PayPal "
				elseif pcv_strPendingReason="other" then
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Unknown - The payment is pending for a unknown reason. For more information, please contact Customer Service"
				else
					pcv_PPStatusNote=pcv_PPStatusNote&"Pending Reason: Unknown - The payment is pending for a unknown reason. For more information, please contact Customer Service"
				end if										
			Case "Processed": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Refunded": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Reversed": pcv_PPStatusNote=pcv_PPStatusNote&""
			Case "Voided": pcv_PPStatusNote=pcv_PPStatusNote&""	
			Case "Unclaimed": pcv_PPStatusNote=pcv_PPStatusNote&"You have refunded a transaction and the recipient may not have confirmed their email address. Or you have Captured Funds that need to be reviewed."						
		End Select
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: PayPal Payment Customer Notes
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		
		'// Check if "Authorize Pending".
		pcv_isAuthorizePending=0						
		query="SELECT pcPay_PayPal_Authorize.idPayPal_Authorize, pcPay_PayPal_Authorize.idOrder, orders.orderDate, orders.orderStatus, orders.gwTransId, pcPay_PayPal_Authorize.amount, pcPay_PayPal_Authorize.paymentmethod, pcPay_PayPal_Authorize.transtype, pcPay_PayPal_Authorize.authcode, pcPay_PayPal_Authorize.idCustomer  FROM pcPay_PayPal_Authorize INNER JOIN orders ON pcPay_PayPal_Authorize.idOrder = orders.idOrder WHERE ((pcPay_PayPal_Authorize.transtype='Authorization') AND (pcPay_PayPal_Authorize.captured=0) AND (pcPay_PayPal_Authorize.idOrder="& pidOrder &"));"
		set rsStatus=server.CreateObject("ADODB.RecordSet")
		set rsStatus=conntemp.execute(query)
		if NOT rsStatus.eof then
			pcv_isAuthorizePending=-1	
		else
			pcv_isAuthorizePending=0	
		end if
		
		'// Check if Payment Captured at PayPal
		if pcv_strCurrentPPStatus="Pending" AND pcv_isAuthorizePending=-1 then
			pcv_isAuthorizePending=-1
		else
			'// This payment was captured at PayPal. Update our records now.
			pcv_isAuthorizePending=0
		end if
		
		'// Check if Payment is "Paid"
		pcv_isPaid=0
		if pcv_strCurrentPPStatus = "Canceled-Reversal" OR pcv_strCurrentPPStatus = "Completed" OR pcv_strCurrentPPStatus = "Partially-Refunded" then
			pcv_isPaid=-1
		end if
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Start: Check Reauthorization Status
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Dim pcv_strReAuthorizeFlag					
		pcv_strReAuthorizeFlag = 0
		
		'// Testing Mode
		'pPayPalAuthorizedDate=dateadd("d",-5,Date())
		'pPayPalOriginalAuthorizedDate=dateadd("d",-20,Date())
		
		'// Check Date is within the Honor Period
		if Date() > dateadd("d",3,pPayPalAuthorizedDate) then
			pcv_strReAuthorizeFlag = 1
		end if
		if Date() > dateadd("d",29,pPayPalOriginalAuthorizedDate) then
			pcv_strReAuthorizeFlag = 2
		end if
		
		amount=cCur(pcv_strAmount)
		curTotal=cCur(ptotal)
		
		'// Testing Mode
		'amount="-1076.13456"
		'curTotal=3002.25
		'response.write amount & "<br />"
		'response.write curTotal
		
		'// Check Current Amount
		if (curTotal > amount) AND (amount>0) then
			pcv_CurPriceDifference = abs(amount - curTotal)
			pcv_MaxAllowedPrice = (amount*1.15)
			'// no greater than $75.00 increase
			if pcv_CurPriceDifference > 75 then
				pcv_strReAuthorizeFlag = 4
			else
				pcv_strReAuthorizeFlag = 3
			end if
			'// no greater 115% increase
			if curTotal => pcv_MaxAllowedPrice then
				pcv_strReAuthorizeFlag = 4
			else
				pcv_strReAuthorizeFlag = 3
			end if
		end if	
		
	
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// End: Check Reauthorization Status
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		'// Display Capture Warning
		if (Date() > dateadd("d",3,pPayPalAuthorizedDate)) AND pcv_strCurrentPPStatus="Pending" then							
			pcv_PPStatusNote=pcv_PPStatusNote&"PayPal recommends that you capture funds within the honor period of three days because PayPal will honor the funds for a 3-day period after the authorization date. "
		end if
		
		'// Check Reauthorize
		pcv_isAuthorizeNeeded=0
		pcv_payPalIllegalTransaction=0
		pcv_isAuthorizeOptional=0
		
		if pcv_strReAuthorizeFlag = 1 AND pcv_isAuthorizePending=-1 then
			pcv_isAuthorizeOptional=-1 '// Reauthorize is required because the transaction has exceded 4 days.
			pcv_PPStatusNote="This order listed above is past the honor period. After the three-day day honor period, you can initiate a reauthorization, which will start a new three-day honor period. However, it will not extend the original authorization period past 29 days. You can reauthorize a transaction only once, up to 115% of the originally authorized amount (not to exceed an increase of $75 USD).  "
		end if
		
		if (pcv_strReAuthorizeFlag = 2 OR pcv_strCurrentPPStatus = "Expired") AND pcv_isAuthorizePending=-1 then
			pcv_isAuthorizeNeeded=-1 '// Reauthorize is required because the transaction has expired.
			pcv_PPStatusNote="This order listed above is over 29 days and must be reauthorized. When your buyer approves an authorization, the buyer’s balance can be placed on hold for a 29-day period to ensure the availability of the authorization amount for capture. You can reauthorize a transaction only once, up to 115% of the originally authorized amount (not to exceed an increase of $75 USD).  "
		end if
		
		if pcv_strReAuthorizeFlag = 3 then
			'// pcv_isAuthorizeNeeded=-1 '// Reauthorize not needed because the price is within valid parameters.
			pcv_PPStatusNote="This order's Total has increased over the original authorization amount, but has not exceded the capture max price (115% of the originally authorized amount or an increase of $75 USD).  "
		end if
		
		if pcv_strReAuthorizeFlag = 4 then
			pcv_isAuthorizeNeeded=-1 '// Reauthorize can not be performed, the price is outside valid parameters.
			pcv_payPalIllegalTransaction=-1 '// Disable the Transaction
			pcv_PPStatusNote="This order's Total has increased over the original authorization max. Unfortunately, this transaction can not be reauthorized because the price is over 115% of the originally authorized amount or exceeded an increase of $75 USD. "
		end if
		%>
		<% if pcv_strPayerEmail<>"" then %>
		<tr>
			<td colspan="2">
				Payer Email: <strong><%=pcv_strPayerEmail %></strong>
			</td>
		</tr>
		<% end if %>
		<% if pcv_strPayerID<>"" then %>
		<tr>
			<td colspan="2">
				Payer ID: <strong><%=pcv_strPayerID %></strong>								
			</td>
		</tr>
		<% end if %>
		<% if pcv_strPayerID<>"" then %>
		<tr>
			<td colspan="2">
				PayPal Transaction ID: <strong><%=pcv_strTransactionID %></strong>
			</td>
		</tr>
		<% end if %>
		<tr>
			<td colspan="2">
			<div>PayPal Transaction Status: <strong><%=pcv_strCurrentPPStatus%></strong></div>
			<% if pcv_PPStatusNote<>"" then %>
			<div style="padding-bottom:4px;" class="pcCPnotes">
				<strong>Note</strong>: <%=pcv_PPStatusNote %>
			</div>
			<% end if %>
			</td>
		</tr>
		<% if pcv_isAuthorizePending=-1 then %>										
		<tr>
			<td colspan="2">
			<div style="padding-bottom:4px;">						
				<a href="http://www.earlyimpact.com/faqs/afmviewfaq.asp?faqid=470" target="_blank">Read this KB article for more information about Authorization & Capture.</a>			</div>
			</td>
		</tr> 		
		<tr>
			<td colspan="2">
				Authorized Amount: <strong><%=scCurSign&money(pcv_strAmount) %></strong>
			</td>
		</tr> 
		<tr>
			<td colspan="2">
				Authorization Period: <strong><%=ShowDateFrmt(pPayPalOriginalAuthorizedDate) %></strong>&nbsp;-&nbsp;<strong><%=ShowDateFrmt(dateadd("d",29,pPayPalOriginalAuthorizedDate)) %></strong> (Valid for 29 days.)  
			</td>
		</tr>
		<tr>
			<td colspan="2">
				Honor Period: <strong><%=ShowDateFrmt(pPayPalAuthorizedDate) %></strong>&nbsp;-&nbsp;<strong><%=ShowDateFrmt(dateadd("d",3,pPayPalAuthorizedDate)) %></strong>  (Full amount guaranteed for 3 days.) 
			</td>
		</tr>
		<% end if %>
	
		<%
		PayPalbtns=0
		if pcv_isPaid=-1 OR pcv_isAuthorizePending=-1 OR pcv_isAuthorizeNeeded=-1 then %>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td colspan="2">
				<span class="pcCPsectionTitle">PayPal Actions</span>
			</td>
		</tr> 
		<tr>
			<td colspan="2">
			<div style="padding-bottom:4px;">
				<input type="hidden" name="PayPalTransID" value="<%=pcgwTransId%>">
				<input type="hidden" name="PayPalTransParentID" value="<%=pcgwTransParentId%>">
				<% if pcv_isAuthorizePending=-1 AND pcv_isAuthorizeNeeded=0 AND pcv_payPalIllegalTransaction=0 then
				PayPalbtns=1 %>
					<input type="submit" name="SubmitPayPal1" value=" Capture "  class="submit2">&nbsp;&nbsp;
				<% end if %>
				<% if (pcv_isAuthorizePending=-1 AND pcv_isAuthorizeNeeded=-1 AND pcv_payPalIllegalTransaction=0) OR (pcv_isAuthorizePending=-1 AND pcv_isAuthorizeOptional=-1 AND pcv_payPalIllegalTransaction=0) then
				PayPalbtns=1 %>
					<input type="submit" name="SubmitPayPal1" value=" Reauthorize "  class="submit2">&nbsp;&nbsp;
				<% end if %>
				<% if pcv_isPaid=-1 AND pcv_PaymentStatus<>6  AND pcv_PaymentStatus<>8 then
				PayPalbtns=1 %>					
					<input type="submit" name="SubmitPayPal1" value=" Refund "  onClick="javascript: if (confirm('This action will NOT cancel the order, it will refund the payment via PayPal. Are you sure you want to continue?')) return true ; else return false ;" class="submit2">&nbsp;&nbsp;					
				<% end if %>
				<% if pcv_isAuthorizePending=-1 then
				PayPalbtns=1 %>												
					<input type="submit" name="SubmitPayPal1" value="  Void  " onClick="javascript: if (confirm('This action will cancel the PayPal authorization and mark the order status as canceled.  Are you sure you want to continue?')) return true ; else return false ;" class="submit2">					
				<% end if %>												
			</div>
			<%if PayPalbtns=1 then%>							
			<div style="padding-bottom:4px;">						
				<a href="JavaScript:win('helpOnline.asp?ref=442')">Help with these buttons</a>
			</div>
			<%else%>
			<div style="padding-bottom:4px;">						
				no actions available
			</div>
			<%end if%>	
			</td>
		</tr>
		<% end if %>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td colspan="2">
				<span class="pcCPsectionTitle">PayPal Tools</span>
			</td>
		</tr> 
		<tr>
			<td colspan="2"><a href="javascript:openwin('popup_PayPalTransSearch.asp');">Search PayPal Transactions</a></td>
		</tr> 
	<% End If '// If pcv_strHideAllPayPal = 1 Then  %>


<% 
'// if pcv_WPPSpecialFeatures=1 AND (pcgwTransId<>"" AND isNULL(pcgwTransId)=False) then  
elseif pcv_WPPSpecialFeatures=1 AND (pcgwTransId="" OR isNULL(pcgwTransId)=True) then 
%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">PayPal</th>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPnotes">
		<div><strong>Alert: We could not display your PayPal Details Section:  </strong></div>
		<div>We can not locate a Transaction ID for this Order.  The Transaction ID is necessary to display PayPal details. This may occur when incomplete orders (orders in which the PayPal payment failed) are processed manually.  If the transaction is processed manually a PayPal Transaction ID will not exist.  Therefore, you will not be able to use any advanced PayPal features with this Order.</div>
		</td>
	</tr>
<%
end if 
%>


<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START: PAYPAL - Display Risk Managment if its available.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If isNULL(pcPay_PayPal_Version)=True Then pcPay_PayPal_Version=""
If isNULL(pcv_strAVSRespond)=True Then pcv_strAVSRespond=""
If isNULL(pcv_strCVNResponse)=True Then pcv_strCVNResponse=""
%>
<% if (pcv_strAVSRespond<>"" OR pcv_strCVNResponse<>"" OR pcv_strPayerStatus<>"" OR pcv_strAddressStatus<>"") AND pcPay_PayPal_Version<>"" then  %>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">PayPal Risk Management</th>
</tr>	
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<% end if %>
<% if pcv_strAVSRespond<>"" AND pcPay_PayPal_Version<>"" then %>
<tr>
	<td colspan="2">
		<%
		select case pcv_strAVSRespond
			case "A": pcv_strAVSRespond="Address only (no ZIP)"
			case "B": pcv_strAVSRespond="Address only (no ZIP)"
			case "D": pcv_strAVSRespond="Address and Postal Code"
			case "F": pcv_strAVSRespond="Address and Postal Code"
			case "P": pcv_strAVSRespond="Postal Code only (no Address)"
			case "W": pcv_strAVSRespond="Nine-digit ZIP code (no Address)"
			case "X": pcv_strAVSRespond="Exact match - Address and nine-digit ZIP code"
			case "Y": pcv_strAVSRespond="Yes - Address and five-digit ZIP"
			case "Z": pcv_strAVSRespond="Five-digit ZIP code (no Address)"
			case "0": pcv_strAVSRespond="All information matched"
			case "2": pcv_strAVSRespond="Part of the address information matched"
			case else: pcv_strAVSRespond="Not Available"						
		end select							
		%>
		AVS Response: <strong><%=pcv_strAVSRespond%></strong> 
	</td>
</tr>
<% end if %>

<% if pcv_strCVNResponse<>"" AND pcPay_PayPal_Version<>"" then %>
<tr>
	<td colspan="2">
		<%
		select case pcv_strCVNResponse
			case "M": pcv_strCVNResponse="CVV2 match"
			case "N": pcv_strCVNResponse="No CVV2 match"
			case "0": pcv_strAVSRespond="CVV2 match"
			case "1": pcv_strAVSRespond="No CVV2 match"
			case else: pcv_strCVNResponse="Not Available"							
		end select							
		%>
		CVV2 Response: <strong><%=pcv_strCVNResponse%></strong>
	</td>
</tr>
<% end if %>

<% if pcv_strPayerStatus<>"" AND pcPay_PayPal_Version<>"" then %>
<tr>
	<td colspan="2">
		Buyer's Status: <strong><%=pcv_strPayerStatus%></strong>
	</td>
</tr>
<% end if %>

<% if pcv_strAddressStatus<>"" AND pcPay_PayPal_Version<>"" then %>
<tr>
	<td colspan="2">
		Buyer's Address Status: <strong><%=pcv_strAddressStatus%></strong>
	</td>
</tr>
<% end if %>
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: PAYPAL - Display Risk Managment if its available.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>