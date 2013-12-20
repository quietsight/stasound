<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pageTitle="View &amp; Edit Active Payment Options"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 
%>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->

<% 
dim query, conntemp, rs

sMode=Request.Form("Submit")

If sMode <> "" Then
	
	If sMode="Add" Then
		iCnt=Request.Form("iCnt")
		call openDb()
		for i=1 to iCnt
			ck=Request("ck" & i)
			If ck="1" Then
				idPayment=Request("id" & i)
	   			query= "Update paytypes SET active=-1 WHERE idPayment="& idPayment
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				set rs=nothing
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
			End If
		next
		call closedb()
	End If
	
	If sMode="Delete" Then
		rCnt=Request.Form("rCnt")
		call openDb()
		for i=1 to rCnt
			ck=Request("ck" & i)
			If ck="1" Then
				idPayment=Request("id" & i)
				query= "Update paytypes SET active=0 WHERE idPayment="& idPayment
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				set rs=nothing
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
			End If
		next
		call closedb()
	End If
	
	response.redirect "PaymentOptions.asp"
	
End If

pcv_strShowSpashScreen=0
%>

<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"><!--#include file="inc_PayPalExpressCheck.asp"--></td>
	</tr>
	<tr> 
		<td>The following payment options are <strong>currently active</strong> on your store:</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 	
		<td> 
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">Real-time credit card processing, etc. - <a href="AddRTPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>

				<% ' get real-time payment types
				call opendb()
				query="SELECT active, idPayment, gwCode, paymentDesc, paymentNickName FROM paytypes WHERE ((gwCode<>2 AND gwCode<>3 AND gwCode<>6 AND gwCode<>7 AND gwCode<>9 AND gwCode<>46 AND gwCode<>80 AND gwCode<>53 AND gwCode<>999999 AND gwCode<>99 AND gwCode<>50) AND gwCode<100) ORDER BY paymentDesc"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				if err.number <> 0 then
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
				If rs.eof then 
					pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
					%>
					<tr>
						<td colspan="3">No real-time payment options found</td>
					</tr>
				<% Else %>
						<%
						do until rs.eof 
							active=rs("active")
							id=rs("idPayment")
							gwCode=rs("gwCode")
							Desc=rs("paymentDesc")
							NickName = rs("paymentNickName")
							If Desc = "LinkPoint" Then
								Desc = "LinkPoint Basic"
								query="SELECT lp_yourpay FROM linkpoint"
								set rsLPObj=Server.CreateObject("ADODB.Recordset")     
								set rsLPObj=conntemp.execute(query)
								LPTypeCheck = rsLPObj("lp_yourpay")
								If LPTypeCheck = "YES" Then
									Desc = "LinkPoint - YourPay"
								End If
								If LPTypeCheck = "API" Then
									Desc = "LinkPoint API "
								End If
								Set rsLPObj = Nothing
							End If								
							If active="-1" Then %>
                                <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td width="90%">  
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" then
											response.write Desc & " is enabled."
										else
											response.write Desc
											if len(NickName)>0 then
												response.write "&nbsp;&nbsp;<i>["&NickName&"]</i>"
											end if
										end if %>&nbsp;&nbsp; 
									</td>
									<td align="center">
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" or gwCode="62" or gwCode="66" then %>
											&nbsp;
										<% else %> 
											<a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwchoice=<%=gwCode%>"><img src="images/pcIconGo.jpg"></a>
										<% end if %>
									</td>
									<td align="center">
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" or gwCode="66" then %>
											&nbsp;
										<% else %> 
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id=<%=id%>&gwChoice=<%=gwCode%>'"><img src="images/pcIconDelete.jpg"></a>
										<% end if %> 
									</td>
								</tr>
							<% end if %>
							<% rs.movenext
						loop
						set rs=nothing
					End If %>
				</table>
			</form>
			
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>

			<% ' get custom payment options
			query="SELECT idcustomCardType, customcardTypes.customcardDesc, paymentDesc, gwcode, paytypes.active, paytypes.idpayment FROM customCardTypes,paytypes WHERE paytypes.paymentDesc=customcardTypes.customcardDesc AND gwcode <> 7;"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?error="& Server.Urlencode("Error in paymentOptions: "&Err.Description) 
			end If %>

			<form name="form3" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr>
						<th width="90%">Debit cards, store cards, and other custom options - <a href="AddCustomCardPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
                    
					<% 
					If rs.eof then 
					pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
					%>
					<tr>
						<td colspan="3">No custom payment options found.</td>
					</tr>
					<% Else
						
						do until rs.eof
							pidcustomCardType=rs("idcustomCardType")
							pcustomcardDesc=rs("customcardDesc")
							ppaymentDesc=rs("paymentDesc")
							pgwcode=rs("gwcode")
							pactive=rs("active")
							pidpayment=rs("idpayment")
								
							If pactive="-1" Then%>
								<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td><%=ppaymentDesc%></td>
									<td align="center"> 
										<a href="modCustomCardPaymentOpt.asp?mode=Edit&idc=<%=pidcustomCardType%>&id=<%=pidpayment%>&gwCode=<%=pgwCode%>"><img src="images/pcIconGo.jpg"></a>	
									</td>
									<td align="center"> 
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='modCustomCardPaymentOpt.asp?mode=Del&idc=<%=pidcustomCardType%>&id=<%=pidpayment%>&gwCode=<%=pgwCode%>'"><img src="images/pcIconDelete.jpg"></a>
									</td>
								</tr>
							<%
								end if
								rs.movenext
								loop
								set rs=nothing

							End If
							%>
				</table>
			</form>
			
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">Offline credit cards, check, Net 30, etc. - <a href="AddCCPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
							
					<%
					pcv_strIsOfflineOptions=0
					query="SELECT idPayment,paymentDesc,active FROM paytypes WHERE gwCode=6"
					set rs=Server.CreateObject("ADODB.Recordset")     
					set rs=conntemp.execute(query)
					if err.number <> 0 then
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
					end If
				
					if rs.eof then
						set rs=nothing
						query="INSERT INTO payTypes (gwCode, paymentDesc, priceToAdd, percentageToAdd, ssl, sslUrl,quantityFrom, quantityUntil, weightFrom, weightuntil, priceFrom, priceuntil, active, Cbtob, Type) VALUES (6,'Credit Card',0, 0, -1,'paymnta_o.asp',0,9999,0,9999,0,9999,0,0,'A');"
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
					else
						id=rs("idPayment")
						Desc=rs("paymentDesc")
						intActive=rs("active")
						set rs=nothing
						
						If intActive="-1" Then %>
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
							<td><%=Desc%></td>
							<td align="center">
								<a href="AddModRTPayment.asp?mode=Edit&id=<%=id%>&gwchoice=6"><img src="images/pcIconGo.jpg"></a>
							<td align="center">
								<a href="javascript:if (confirm('You are about to remove this payment option. This will also disable all credit card options associated with it. Are you sure you want to complete this action?')) location='AddModRTPayment.asp?mode=Del&id=<%=id%>&gwChoice=6'"><img src="images/pcIconDelete.jpg"></a>
							</td>
						</tr>
						<%
						query="SELECT CCType,CCCode FROM CCtypes WHERE active=-1 ORDER BY CCType ASC"
						set rs=Server.CreateObject("ADODB.Recordset")     
						rs.Open query, conntemp
						if err.number <> 0 then
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
						end If 
						%>
						<tr> 
							<td colspan="3"> 
								<table class="pcCPcontent">
									<% do until rs.eof
										pcv_strIsOfflineOptions=1
										strCCType=rs("CCType")
										strCCCode=rs("CCCode") %>
										<tr> 
											<td width="93%" style="padding-left: 15px;"><%=strCCType%> is enabled</td>
											<td width="7%" align="center">
											<a href="javascript:if (confirm('You are about to remove this credit card payment option. Are you sure you want to complete this action?')) location='AddModRTPayment.asp?mode=Del&TYPE=CC&id=<%=id%>&CCCode=<%=strCCCode%>'"><img src="images/pcIconDelete.jpg" border="0"></a></td>
										</tr>
										<% rs.moveNext
									loop
									set rs=nothing 
									%>
								</table>
							</td>
						</tr>
						<% 						
						end if 							
					end if 
					%>
				
					<% 
					query="SELECT idPayment, paymentDesc, active FROM paytypes WHERE gwCode=7"
					set rs=Server.CreateObject("ADODB.Recordset")     
					set rs=conntemp.execute(query)
					if err.number <> 0 then
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in PaymentOptions.asp: "&Err.Description) 
					end If
	
					If rs.eof then					
					Else						
						do until rs.eof 						
						btActive=rs("active")
						id=rs("idPayment")
						Desc=rs("paymentDesc")
			 
							If btActive="-1" Then 
							pcv_strIsOfflineOptions=1
							%>
								<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td><%=Desc%></td>
									<td width="14%" align="center">
										<a href="AddModRTPayment.asp?mode=Edit&id=<%=id%>&gwChoice=7"><img src="images/pcIconGo.jpg"></a>
									</td>
									<td width="13%" align="center">
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='AddModRTPayment.asp?mode=Del&id=<%=id%>&gwChoice=7'"><img src="images/pcIconDelete.jpg"></a></td>
								</tr>
							<% end if %>		
							<% 
							rs.movenext
						loop
						set rs=nothing 
						
					end if
					call closedb() 
					%>
					<% 
					if pcv_strIsOfflineOptions=0 then 
						pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
						%>
						<tr>
							<td colspan="3">No Offline payment options found</td>
						</tr>
					<% end if %>
				</table>
			</form>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 	
		<td> 
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">PayPal Payment Options - <a href="pcPaymentSelection.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>

						<% ' get real-time payment types
						call opendb()
						query="SELECT active,idPayment,gwCode,paymentDesc FROM paytypes WHERE ((gwCode=2 OR gwCode=3 OR gwCode=9 OR gwCode=46 OR gwCode=99 OR gwCode=53 OR gwCode=80 OR gwCode=999999) AND gwCode<>50) ORDER BY paymentDesc"
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
						if err.number <> 0 then
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
						end If
						If rs.eof then 
							pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
							%>
							<tr>
								<td colspan="3">No PayPal payment options found</td>
							</tr>
						<% Else %>
						<% 
						do until rs.eof 
							active=rs("active")
							id=rs("idPayment")
							gwCode=rs("gwCode")
							Desc=rs("paymentDesc")
								
							If active="-1" Then %>
							<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td>  
										<%=Desc%>&nbsp;&nbsp; 
									</td>
									<td width="14%" align="center">
										<a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwchoice=<%=gwCode%>"><img src="images/pcIconGo.jpg"></a>
									</td>
									<td width="13%" align="center">
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id=<%=id%>&gwChoice=<%=gwCode%>'"><img src="images/pcIconDelete.jpg"></a>
									</td>
								</tr>
							<% end if %>
							<% rs.movenext
						loop
						set rs=nothing
					End If %>
				</table>
			</form>
			
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>

			<% 
			'// Check if Google Checkout is active
			pcv_intGoogleActive=GOOGLEACTIVE
			%>
			<form name="form4" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr>
						<th width="90%">Google Checkout</th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
					<% If pcv_intGoogleActive=-1 Then %>							
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                        <td>Google Checkout</td>
                        <td align="center">
                            <a href="ConfigureGoogleCheckout2.asp"><img src="images/pcIconGo.jpg"></a>
                        </td>
                        <td align="center">
                            <a href="javascript:if (confirm('You are about to remove Google Checkout. Are you sure you want to complete this action?')) location='ConfigureGoogleCheckout2.asp?mode=Del'"><img src="images/pcIconDelete.jpg"></a></td>
                    </tr>
                    <%
                    else 
                        pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
                        %>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                        <td>Google Checkout is not Enabled - <a href="ConfigureGoogleCheckout.asp" title="Enable Google Checkout">Add it now</a></td>
                        <td align="center">&nbsp;</td>
                        <td align="center">&nbsp;</td>
                    </tr>
                    <% end if %>
				</table>
			</form>
		</td>
	</tr>
	<%
	If pcv_strShowSpashScreen = 5 Then
		response.Redirect("pcPaymentSelection.asp")
	End If
	%>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="center">
		<form class="pcForms">
			<input type="button" value="Add New" onClick="location.href='pcPaymentSelection.asp'">&nbsp;
			<input type="button" value="Set Display Order" onClick="location.href='OrderPaymentOptions.asp'">&nbsp;
			<input type="button" value="Back" onClick="javascript:history.back()">
		</form>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
<%
'// If session variable says to setup PayPal Express, redirect to it
if session("pcSetupPayPalExpress") <> "" AND pcv_strHideAlert=0 AND session("pcPayPalExpressCookie")="" then
	%>
	<div id="GlobalMsgDialog" title="PayPal Express Checkout" style="display:none">
		<div id="GlobalMsg" class="pcSuccessMessage" style="width: 100%;">
            <div id="showPayPalExpressModal">
            	<div id="showPayPalExpressImage"><img src="images/paypal_29794_screenshot2.gif"></div>
            	<div id="showPayPalExpressTitleModal">Would you like to add Express Checkout?</div>
                <div id="showPayPalExpressTextModal">According to Jupiter Research, 23% of online shoppers consider PayPal one of their favorite ways to pay online<sup>1</sup>. Accepting PayPal in addition to credit cards is proven to increase your sales<sup>2</sup>. <a href="https://www.paypal.com/us/cgi-bin/?&cmd=_additional-payment-overview-outside" target="_blank">See Quick Demo</a>.</div>
                <div id="showPayPalExpressTextSmallModal">(1) Payment Preferences Online, Jupiter Research, September 2007. <br />
                (2) Applies to online businesses doing up to $10 million/year in online sales. Based on a Q4 2007 survey of PayPal shoppers conducted by Northstar Research, and PayPal internal data on Express Checkout transactions.
                </div>
            </div>
		</div>
	</div>
	<style>
	
		#showPayPalExpressImage {
			float: right;
		}
        
		#showPayPalExpressTitleModal {
			font-size: 15px;
			font-weight: bold;
			margin-bottom: 10px;
		}
        
        #showPayPalExpressTextModal {
            color: #666;
        }
        
        #showPayPalExpressTextSmallModal {
            color: #999;
            font-size: 9px;
			margin-top: 6px;
        }
    </style>
	<script>
		$(document).ready(function()
		{
			$("#GlobalMsgDialog").dialog({
					bgiframe: true,
					autoOpen: false,
					resizable: false,
					width: 700,
					height: 270,
					modal: true,
					buttons: {
						' No Thanks ': function() {
								PayPalExpressCookie(1);
								$(this).dialog('close');						
						},
						' Maybe Later ': function() {
								PayPalExpressCookie(2);
								$(this).dialog('close');						
						},
						' Yes ': function() {
								location='pcConfigurePayment.asp?gwchoice=999999';
								$(this).dialog('close');						
						}
					}
			});
			$("#GlobalMsgDialog").dialog('open');
			
		function PayPalExpressCookie(duration) {
				var isChecked = 0;
				if ($("#PayPalExpressActive").is(':checked')) 
				{
					isChecked = 1;
				}
				$.ajax({
					type: "POST",
					url: "inc_PayPalExpressCookie.asp",
					data: "duration=" + duration,
					timeout: 5000,
					global: false,
					success: function(data, textStatus){
						if (data=="SECURITY")
						{
							window.location="login_1.asp";
							
						} else {
							
							if (data=="OK")
							{
								
								// no action
								
							} else {
								
								// no action
								
							}
						}
					}
				});
		}
			
		});
	</script>
    <%
end if
%>
<!--#include file="AdminFooter.asp"-->