<%
'// we can charge the order if
'if porderStatus = 2 AND pcv_PaymentStatus = 1 then

'// we can ship the order if
'if porderStatus = 3 AND pcv_PaymentStatus = 2 then

'// we can archive the order if
'if porderStatus = 4 AND pcv_PaymentStatus = 2 then
%>
<tr> 
	<td colspan="2">
	Current Order Status: <b><%=os%></b>
	<br /><br />
	This order was placed with Google Checkout. 
	As the order is processed, its status changes will appear on the buyer's purchase history page 
	and on your Orders page in the Merchant Center. 
	<a target="_blank" href="https://checkout.google.com/sell/">Go to the Merchant Center.</a>
	</td>
</tr>
<%
'// Start - Charge Order
if porderStatus="2" AND pcv_PaymentStatus="1"then
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/quick.gif" width="24" height="18"></td>
	<td><strong>Charge Order</strong>. 
	&nbsp;
	<br />
	After you have verified the accuracy and legitimacy of the order, you can initiate a &quot;Charge Request&quot; 
	by clicking on the button below. 
	</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td> 
		<% if pcv_PaymentStatus="7" then %>
			<div class="pcCPnotes">
				This order's status is charging. Refresh this page periodically to view an updated status. 
				An order should not remain &quot;Charging&quot; for more than a minute.
			</div>	
			<br />										
			<input type="submit" name="Submit4Google" value=" Charging " disabled="disabled" style="background-color:#F5F5F5; font-style:italic; color:#CCCCCC; border: solid 1px #CCCCCC;" />
		<% else %>
			<input name="GoogleMethod" type="hidden" value="charge" />
			<input type="submit" name="Submit4Google" value="Charge This Order" class="submit2">
		<% end if %>
	</td>
</tr>
<%
end if
'// End - Charge Order


'// Start - Ship Order
if (porderStatus="2" OR porderStatus="3") AND (pcv_PaymentStatus="2" OR pcv_PaymentStatus="6") AND pcv_PaymentStatus<>"7" then
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/quick.gif" width="24" height="18"></td>
	<td><strong>Ship and confirm the order</strong>. &nbsp;
	<br />
	The customer has been charged.  Your order is now available for shipping.  
	Ship this order using the <a href="#" onclick="change('tabs1', '');change('tabs2', '');change('tabs3', '');change('tabs4', 'current');change('tabs5', '');change('tabs6', '');change('tabs7', '');showTab('tab4');form2.ActiveTab.value = 4">Shipping Wizard &gt;&gt;</a>, or 
	<a target="_blank" href="https://checkout.google.com/sell/">Google's Merchant Center  &gt;&gt;</a>.
	If you ship via Google's Merchant Center your order's status will update in ProductCart. 
	When you ship with ProductCart, after the order is shipped, you will use the "Mark as Shipped" button in the 
	"Synchronize Shipping..." section below.</td>							
</tr>
<%
end if
'// End - Ship Order


'// Start - Mark Order Shipped
if (porderStatus="2" OR porderStatus="3" OR porderStatus="4") AND (pcv_PaymentStatus="2" OR pcv_PaymentStatus="6") then
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/quick.gif" width="24" height="18"></td>
	<td><strong>Synchronize Shipping with Google's Merchant Center</strong>. &nbsp;
	<br />
	Click the &quot;Mark as Shipped&quot; button to send your shipping information to Google's Merchant Center.  
	If your shipment has tracking information it will appear on the buyer's account page. 
	Please note Google does not support partial shipments at this time. ProductCart automatically performs the &quot;Mark as Shipped&quot; action on fully shipped orders. </td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td> 
		<input name="GoogleMethod" type="hidden" value="mark" />
		<input name="strShipper" type="hidden" value="<%=Shipper%>" />
		<% if porderStatus="2" OR porderStatus="3" then %>
			<div class="pcCPnotes">
			Note: This order has not yet been shipped. 
			To &quot;Mark as Shipped&quot; you must first ship and confirm this order above.</div>	
			<br />										
			<input type="submit" name="Submit5Google" value=" Mark as Shipped " disabled="disabled" style="background-color:#F5F5F5; font-style:italic; color:#CCCCCC; border: solid 1px #CCCCCC;" />
			
		<% elseif  porderstatus="7" then %>
			<div class="pcCPnotes">
                Note: This order is partially shipped. We recommend that you ship all packages before synchronizing with Google Checkout.  
                ProductCart automatically performs the &quot;Mark as Shipped&quot; action on fully shipped orders.
			</div>
			<input type="submit" name="Submit5Google" value=" Mark as Shipped " class="submit2">
		<% else %>
			<input type="submit" name="Submit5Google" value=" Mark as Shipped Again " onClick="javascript: if (confirm('ProductCart automatically performs the &quot;Mark as Shipped&quot; action on fully shipped orders. The button below should only be used to repeat this action. For example, in a rare case where Google Checkout does not receive the first request and you need to manually mark your order as shipped. Are you sure you want to continue?')) return true ; else return false ;" class="submit2">
		<% end if %>
	</td>
</tr>
<%
end if
'// End - Mark Order Shipped


'// Start - Archive Order
if (porderStatus="5" OR porderStatus="10" OR porderStatus="11") AND porderStatus<>"12" then
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/quick.gif" width="24" height="18"></td>
	<td><strong>Archive this Order with Google Merchant Center</strong>. &nbsp;
	<br />
	This order is marked as shipped and has the "Delivered" status. 
	Click the "Archive this Order" button when you're finished, this will remove it from the list of orders 
	that appear on the Orders page of the Merchant Center.
</tr>
<tr> 
	<td>&nbsp;</td>
	<td> 
		<input name="GoogleMethod" type="hidden" value="archive" />
		<input type="submit" name="Submit6Google" value="Archive this Order" class="submit2">
	</td>
</tr>
<%
end if
'// END - Archive Order



'// Start - Refund Order
if (porderStatus="3" OR porderStatus="4" OR porderStatus="10") AND pcv_PaymentStatus="2" then
%>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/delete2.gif" width="23" height="18"></td>
	<td valign="top">
	<strong>Refund Order</strong>. &nbsp;
	<br />
	You may refund this order before you cancel the order. You do not have to cancel an order after you issue the refund.
<!--									<div class="pcCPnotes">
		Note: Cancelling an order does issue a refund. To issue a refund visit you Google Merchant Center.
	</div> -->
	</td>
</tr>
<tr>
	<td></td>
	<td align="top">Reason:<br>
	<textarea name="strRefundReason" cols="60" rows="2" wrap="VIRTUAL"></textarea>
	</td>
</tr>
<tr>
	<td></td>
	<td align="top">Comments:<br>
	<textarea name="strRefundComment" cols="60" rows="4" wrap="VIRTUAL"></textarea>
	</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td> 
		<input name="GoogleMethod4" type="hidden" value="refund" />
		<input type="submit" name="Submit9Google" value="Refund Order" class="submit2">
	</td>
</tr>
<%
end if
'// END - Refund Order


								
'// Start - Cancel Order
if (porderStatus="2" OR porderStatus="3") AND (pcv_PaymentStatus="0" OR pcv_PaymentStatus="1" OR pcv_PaymentStatus="2" OR pcv_PaymentStatus="6") then
%>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/delete2.gif" width="23" height="18"></td>
	<td valign="top">
	<strong>Cancel Order</strong>. &nbsp;
	<br />
	You may cancel this order at anytime before the order is shipped. To cancel, click on the &quot;Cancel&quot; button below.
	<div class="pcCPnotes">
		Note: Canceling an order does NOT issue a refund. Refunds must be issued before you cancel the order.
	</div>
	</td>
</tr>
<tr>
	<td></td>
	<td align="top">Reason:<br>
	<textarea name="strReason" cols="60" rows="2" wrap="VIRTUAL"></textarea>
	</td>
</tr>
<tr>
	<td></td>
	<td align="top">Comments:<br>
	<textarea name="strComment" cols="60" rows="4" wrap="VIRTUAL"></textarea>
	</td>
</tr>
<!--								<tr>
	<td align="right">
	<input type="checkbox" name="sendEmailCanc" value="YES" checked class="clearBorder"></td>
	<td>Send <u>order cancelled</u> e-mail when the order is cancelled.</td>
</tr> -->
<tr> 
	<td>&nbsp;</td>
	<td> 
		<input name="GoogleMethod2" type="hidden" value="cancel" />
		<input type="submit" name="Submit7Google" value="Cancel Order" class="submit2">
	</td>
</tr>
<%
end if
'// END - Cancel Order


'// Start - Send Buyer a Message
'if (porderStatus="2" OR porderStatus="3") AND (pcv_PaymentStatus="0" OR pcv_PaymentStatus="1" OR pcv_PaymentStatus="2") then
%>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr> 
	<td valign="top" align="right"><img src="images/delete2.gif" width="23" height="18"></td>
	<td valign="top">
	<strong>Send Buyer a Message</strong>. &nbsp;
	<br />
	You may send the buyer a message at anytime. The message will appear in the buyer's account mailbox.
	</td>
</tr>

<tr>
	<td></td>
	<td align="top">Comments:<br>
	<textarea name="strBuyerMessage" cols="60" rows="4" wrap="VIRTUAL"></textarea>
	</td>
</tr>
<tr> 
	<td>&nbsp;</td>
	<td> 
		<input name="GoogleMethod3" type="hidden" value="message" />
		<input type="submit" name="Submit8Google" value="Send Message" class="submit2">										
	</td>
</tr>
<%
'// END - Send Buyer a Message



'// START - Standard Order Options
if (RMAVar=0 AND RMAStatus=0) AND (porderStatus>"3" AND porderStatus<"11") then
%>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Returns</th>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="2"><strong>Create Return Merchandise Authorization Number</strong> (&quot;RMA&quot; Number)</td>
	</tr>
	<tr>
		<td valign="top"><img src="images/quick.gif" width="24" height="18"></td>
		<td><input type="button" name="RMAInfo" value="Create RMA Number" class="submit2" onClick="location.href='genRma.asp?idOrder=<%=qry_ID%>'"></td>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
<% 
end if

if RMAVar=1 AND RMAStatus=0 then %>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="2" valign="top"><strong>RMA Information</strong></td>
	</tr>
	<tr>
		<td colspan="2">
		<% if trim(prmaNumber)="" then %>
		A customer requested a RMA (Return Merchandise Authorization). Click below to generate an RMA number and send it to the customer.
		<% else %>
		A Return Merchandise Authorization number was created for this order. Here are the details:
		<% end if %>
		</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
		<td> 
			<table class="pcCPcontent" style="width:auto;">
				<tr> 
					<td nowrap="nowrap">Date Submitted: </td>
					<td nowrap="nowrap"><%=ShowDateFrmt(pRMADate)%></td>
				</tr>
				<tr> 
					<td nowrap="nowrap">RMA Number:</td>
					<td nowrap="nowrap"><%=pRmaNumber%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
		<td><input type="button" name="RMAInfo" value="View RMA Information" onClick="location.href='modRmaa.asp?idOrder=<%=qry_ID%>&idRMA=<%=pIdRMA%>'"></td>
	</tr>
<% end if %>

<% if RMAVar=1 AND RMAStatus=1 then %>
	<tr> 
		<td colspan="2">&nbsp;</td>
	</tr>
<tr> 
<td colspan="2" class="pcCPspacer"></td>
</tr>							
	<tr> 
		<td colspan="2"><b>RMA Information</b></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td>
		<table class="pcCPcontent" style="width:auto;">
			<tr>
				<td nowrap="nowrap">RMA Number:</td>
				<td nowrap="nowrap"><b><%=pRmaNumber%></b></td>
			</tr>
			<tr>
				<td nowrap="nowrap">Status:</td>
				<td><b><%=pRmaStatus%></b></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td><input type="button" name="RMAInfo" value="View RMA Information" onClick="location.href='modRmaa.asp?idOrder=<%=qry_ID%>&idRMA=<%=pIdRMA%>'"></td>
	</tr>
<% end if
'// END - Standard Order Options



if porderStatus>2 then
'// START - Re Send Order Emails
%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Resend E-mail Messages</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<%
	'// Resend Confirmation email								
	if (porderStatus = "3" OR porderStatus = "4" OR porderStatus = "7" OR porderStatus = "8" OR porderStatus = "10") then%>
		<tr>
			<td colspan="2">
			<strong>Order &quot;Charged&quot; Confirmation E-mail</strong><br>
			Use the button below to <u>resend the Order Confirmation e-mail</u> to this 
			customer (e.g. the customer reports that they did not receive it).
			</td>
		</tr>
		<%
		query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=1;"
		set rsQ=connTemp.execute(query)
		pcv_AdmComments=""
		if not rsQ.eof then
			pcv_AdmComments=rsQ("pcACom_Comments")
		end if
		set rsQ=nothing%>
		<tr>
			<td colspan="2">
			<div>Comments:</div>
			<div style="padding: 5px 0 5px 0;">
				<textarea name="AdmComments1A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
			</div>
			<input type="submit" name="Submit4A" value="Resend Order Confirmation e-mail" class="submit2">
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
	<%
	end if
	%>

	<%						
	'// Resend order shipped e-mail
	if porderStatus="4" OR porderStatus="10" then
	
		query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=3;"
		set rsQ=connTemp.execute(query)
		pcv_AdmComments=""
		if not rsQ.eof then
			pcv_AdmComments=rsQ("pcACom_Comments")
		end if
		set rsQ=nothing
		%>
		<tr>
			<td colspan="2">
				<strong>Order Shipped E-mail</strong><br>
				Use the button below to <u>resend the Order Shipped e-mail</u> to this customer (e.g. the customer reports that they did not receive it).
			</td>
		</tr>
		<tr>
			<td colspan="2">
			<div>Comments:</div>
			<div style="padding: 5px 0 5px 0;">
				<textarea name="AdmComments3A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
				<input name="sendEmailShip" type="hidden" value="YES">
			</div>
			<input type="submit" name="Submit5A" value="Resend Order Shipped Email" class="submit2"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<%
	end if
	
	'// Resend order cancelled e-mail
	if porderStatus="5" then
		query="SELECT pcACom_Comments FROM pcAdminComments WHERE idOrder=" & qry_ID & " AND pcACom_ComType=5;"
		set rsQ=connTemp.execute(query)
		pcv_AdmComments=""
		if not rsQ.eof then
			pcv_AdmComments=rsQ("pcACom_Comments")
		end if
		set rsQ=nothing
		%>
		<tr>
			<td colspan="2">
			Use the button below to <u>resend the Order Cancelled email</u> to this 
			customer (e.g. the customer reports that they did not receive it).
			</td>
		</tr>
		<tr>
			<td colspan="2">
			<div>Comments:</div>
			<div style="padding: 5px 0 5px 0;">
			<textarea name="AdmComments5A" cols="60" rows="5" wrap="VIRTUAL"><%=pcv_AdmComments%></textarea>
			</div>
			<input type="submit" name="Submit10A" value="Resend Order Cancelled Email" class="submit2">
			</td>
		</tr>
		<%
	end if
'// END - Re Send Order Emails
end if '// if porderStatus>"2" then
%>