<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Checkout Options"
pageIcon="pcv4_icon_checkout.png"
Section="layout" 
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->

<% pcPageName="checkoutOptions.asp"
'/////////////////////////////////////////////////////
'// Retrieve current database data
'/////////////////////////////////////////////////////
dim query, rs, conntemp
%>
<!--#include file="pcAdminRetrieveSettings.asp"-->
<%
pcv_isReferLabelRequired=false
pcv_isViewReferRequired=false
pcv_isRefNewCheckoutRequired=false
pcv_isRefNewRegRequired=false
pcv_isAllowNewsRequired=false
pcv_isNewsCheckoutRequired=false
pcv_isNewsRegRequired=false
pcv_isTermsRequired=false
pcv_isGuestCheckoutOptRequired=false
pcv_isTermsLabelRequired=false
pcv_isTermsCopyRequired=false
pcv_isTermsShownRequired=false
pcv_isNewsLabelRequired=false
pcv_isDFLabelRequired=false
pcv_isDFShowRequired=false
pcv_isDFReqRequired=false
pcv_isTFLabelRequired=false
pcv_isTFShowRequired=false
pcv_isTFReqRequired=false
pcv_isDTCheckRequired=false
pcv_isDeliveryZipRequired=false
pcv_isCustomerIPAlertRequired=false

' Update referrer list
Dim pcv_refAction
pcv_refAction = getUserInput(request("action"),0)
If pcv_refAction <> "" and pcv_refAction <> "update" then
	IDrefer=getUserInput(request("IDrefer"),20)
	call opendb()
	Select Case pcv_refAction
		Case "inactive": query="update Referrer set removed=1 where IDrefer=" & IDrefer &";"
		Case "active": query="update Referrer set removed=0 where IDrefer=" & IDrefer &";"
		Case "del": query="delete from Referrer where IDrefer=" & IDrefer &";"
	End Select	
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	call closedb()
end if 

if pcv_refAction = "update" then
	newrefer=getUserInput(request("newrefer"),50)
	if newrefer<>"" then
		call opendb()
		query="insert into Referrer (Name,SortOrder) values ('" & newrefer & "',0);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
		call closedb()
	end if
	iCnt=getUserInput(request("iCnt"),20)
	if iCnt<>"" then
		call opendb()
		iCnt1=cint(iCnt)
		For i=1 to iCnt
			IDrefer=request("ID" & i)
			SortOrder=request("ORD" & i)
				if not isNumeric(SortOrder) then
					SortOrder = 0
				end if
			query="update Referrer Set SortOrder=" & SortOrder & " where IDrefer=" & IDRefer &";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
		Next
		call closedb()
	end if
end if

' End update referrer list
if request("Submit1")="Update" then
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = "One of more fields were not filled in correctly."
	
	'// validate all fields
	pcs_ValidateTextField	"ReferLabel", pcv_isReferLabelRequired, 250
	pcs_ValidateTextField	"ViewRefer", pcv_isViewReferRequired,2
	pcs_ValidateTextField	"RefNewCheckout", pcv_isRefNewCheckoutRequired, 2
	pcs_ValidateTextField	"RefNewReg", pcv_isRefNewRegRequired, 2
	pcs_ValidateTextField	"Terms", pcv_isTermsRequired, 2
	pcs_ValidateTextField	"GuestCheckoutOpt", pcv_isGuestCheckoutOptRequired, 2
	pcs_ValidateTextField	"TermsLabel", pcv_isTermsLabelRequired, 150
	pcs_ValidateTextField	"TermsCopy", pcv_isTermsCopyRequired, 0
	pcs_ValidateTextField	"TermsShown", pcv_isTermsShownRequired, 2
	pcs_ValidateTextField	"AllowNews", pcv_isRefAllowNews, 2
	pcs_ValidateTextField	"NewsCheckout", pcv_isNewsCheckoutRequired, 2
	pcs_ValidateTextField	"NewsReg", pcv_isNewsRegRequired, 2
	pcs_ValidateTextField	"NewsLabel", pcv_isNewsLabelRequired, 250
	pcs_ValidateTextField	"DFLabel", pcv_isDFLabelRequired, 250
	pcs_ValidateTextField	"DFShow", pcv_isDFShowRequired, 2
	pcs_ValidateTextField	"DFReq", pcv_isDFReqRequired, 2
	pcs_ValidateTextField	"TFLabel", pcv_isTFLabelRequired, 250
	pcs_ValidateTextField	"TFShow", pcv_isTFShowRequired, 2
	pcs_ValidateTextField	"TFReq", pcv_isTFReqRequired, 2
	pcs_ValidateTextField	"DTCheck", pcv_isDTCheckRequired, 2
	pcs_ValidateTextField	"DeliveryZip", pcv_isDeliveryZipRequired, 2
	pcs_ValidateTextField "CustomerIPAlert", pcv_isCustomerIPAlertRequired, 2

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError& "&lmode=" & pcLoginMode
	End If
	
	'/////////////////////////////////////////////////////
	'// Set Local Variables for Setting
	'/////////////////////////////////////////////////////
	if NOT validNum(Session("pcAdminRefNewCheckout")) OR Session("pcAdminRefNewCheckout")<>"1" then
		Session("pcAdminRefNewCheckout")="0"
	end if
	if NOT validNum(Session("pcAdminRefNewReg")) OR Session("pcAdminRefNewReg")<>"1" then
		Session("pcAdminRefNewReg")="0"
	end if
	if NOT validNum(Session("pcAdminNewsCheckout")) OR Session("pcAdminNewsCheckout")<>"1" then
		Session("pcAdminNewsCheckout")="0"
	end if
	if NOT validNum(Session("pcAdminNewsReg")) OR Session("pcAdminNewsReg")<>"1" then
		Session("pcAdminNewsReg")="0"
	end if
	
	pcStrReferLabel = Session("pcAdminReferLabel")
	pcIntViewRefer = Session("pcAdminViewRefer")
	pcIntRefNewCheckout = Session("pcAdminRefNewCheckout")
	pcIntRefNewReg = Session("pcAdminRefNewReg")
	pcIntGuestCheckoutOpt = Session("pcAdminGuestCheckoutOpt")
	pcIntTerms = Session("pcAdminTerms")
	pcStrTermsLabel = Session("pcAdminTermsLabel")
	pcStrTermsCopy = Session("pcAdminTermsCopy")
	pcIntTermsShown = Session("pcAdminTermsShown")
	pcIntAllowNews = Session("pcAdminAllowNews")
	pcIntNewsCheckout = Session("pcAdminNewsCheckout")
	pcIntNewsReg = Session("pcAdminNewsReg")
	pcStrNewsLabel = Session("pcAdminNewsLabel")
	pcStrDFLabel = Session("pcAdminDFLabel")
	pcStrDFShow = Session("pcAdminDFShow")
	pcStrDFReq = Session("pcAdminDFReq")
	pcStrTFLabel = Session("pcAdminTFLabel")
	pcStrTFShow = Session("pcAdminTFShow")
	pcStrTFReq = Session("pcAdminTFReq")
	pcStrDTCheck = Session("pcAdminDTCheck")
	pcStrDeliveryZip = Session("pcAdminDeliveryZip")
	pcStrCustomerIPAlert = Session("pcAdminCustomerIPAlert")

	'/////////////////////////////////////////////////////
	'// Update database with new Settings
	'/////////////////////////////////////////////////////
	%>
	<!--#include file="pcAdminSaveSettings.asp"-->
	<% response.redirect pcPageName
end if %>
		<table class="pcCPcontent">
			<tr> 
				<td colspan="2">
					<table class="pcCPcontent">
						<tr>
							<td colspan="2">This page contains several settings that affect the checkout process:</td>
						</tr>
						<tr>
							<td width="33%" valign="top">
								<ul>
									<li><a href="#referrer">Referrer drop down field</a></li>
									<li><a href="#guest">Guest Checkout Options</a></li>
								</ul>
							</td>
							<td width="33%" valign="top">
								<ul>
									<li><a href="#terms">Terms and conditions agreement</a></li>
									<li><a href="#newsletter">Newsletter settings</a></li>
									<li><a href="#datetime">Custom date/time fields</a></li>
								</ul>
							</td>
							<td width="33%" valign="top">
								<ul>
									<li><a href="#zipcodes">Limit delivery area by zip code</a></li>
									<li><a href="#CustomerIP">Customer's IP Address Alert</a></li>
									<li><a href="manageCustFields.asp">Special customer fields</a></li>
								</ul>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
		</table>
	
	<form name="form1" method="post" action="checkoutOptions.asp?action=update" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<th colspan="2"><a name="referrer"></a>Referrer Drop Down Field</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">You have the ability to require a customer to fill out a &quot;Referrer&quot; field when checking out or registering on the store for the first time. This can help you determine where your customers are coming from. Of course, you may also use this field for other purposes.</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="right" width="25%">Available Selections:</td>
			<td width="75%">
				<input name="newrefer" type="text" id="referrerAdd" size="20">	
				<input name="Submit2" type="submit" value="Add New" class="submit2"> 
			</td>
		</tr>
		<tr> 
			<td colspan="2" align="center">
				<table class="pcCPList">
				<% Call Opendb()
					query="select IDrefer, [name], sortOrder, removed from Referrer order by SortOrder;"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					if rs.eof then
					set rs=nothing
				%>
					<tr>
						<td colspan="3">No Selections found.</td>
					</tr>
				<%else%>
					<tr> 
						<td width="90%">Referrer</td>
						<td>Order</td>
						<td>Action</td>
					</tr>
					<% iCnt=0
					do while not rs.eof
						intIdRefer=rs("IDrefer")
						strName=rs("name")
						intSortOrder=rs("sortOrder")
						intRemoved=rs("removed")
						iCnt=iCnt+1	%>
						<tr> 
							<td>
								<input type="hidden" name="ID<%=iCnt%>" value="<%=intIdRefer%>"> 
								<%=strName%>&nbsp;<% if intRemoved <> 0 then %><span class="pcSmallText">[Inactive]</span><% end if %>
							</td>
							<td><input name="Ord<%=iCnt%>" type="text" value="<%=intSortOrder%>" size="3"></td>
							<td nowrap>
							<script language="JavaScript"><!--
									function newWindow(file,window) {
											catWindow=open(file,window,'resizable=no,width=380,height=160,scrollbars=1');
											if (catWindow.opener == null) catWindow.opener = self;
									}
								//--></script>
						  <a href="javascript:newWindow('refer_popup.asp?idrefer=<%=rs("idrefer")%>','window2')"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit" title="Edit" border="0"></a>&nbsp;<% if intRemoved = 0 then %><a href="javascript:if (confirm('Make this selection inactive if you no longer want to show it in the storefront. This will not affect your customer and order history.')) location='checkoutOptions.asp?action=inactive&idrefer=<%=rs("idrefer") %>'"><img src="images/pcIconMinus.jpg" width="12" height="12" alt="Make Inactive" title="Make Inactive" border="0"></a>
						  <% else %><a href="javascript:if (confirm('Make this selection active if you want to start showing it again in the storefront.')) location='checkoutOptions.asp?action=active&idrefer=<%=rs("idrefer") %>'"><img src="images/pcIconPlus.jpg" width="12" height="12" alt="Make Active" title="Make Active" border="0"></a>
						  <% end if %>&nbsp;<a href="javascript:if (confirm('You are about to remove this selection item from your database. Are you sure you want to complete this action? If you want to keep this piece of information in your customer and order history, make the selection inactive instead of deleting it.')) location='checkoutOptions.asp?action=del&idrefer=<%=rs("idrefer") %>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" border="0"></a>&nbsp;</td>
				  </tr>
            <%
							rs.MoveNext
							loop
							set rs=nothing
							call closedb() %>
						<tr> 
							<td></td>
							<td>
								<input type="hidden" name="iCnt" value="<%=iCnt%>">
								<input name="submit" type="submit" value="Update Order">
							</td>
							<td></td>
						</tr>
				<%END IF%>
			</table>
			</td>
		</tr>
	</table>
	</form>
	<form name="form2" method="post" action="<%=pcPageName%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="right" width="25%">Field Description:</td>
			<td align="left" width="75%"><input name="ReferLabel" type="text" value="<%=ReferLabel%>" size="35">
				(e.g. How did you hear about us?)</td>
		</tr>
		<tr> 
			<td align="right">Field Required?</td>
			<td>
				<input type="radio" name="ViewRefer" value="1" <%if ViewRefer=1 then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="ViewRefer" value="0" <%if not (ViewRefer=1) then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td rowspan="2" align="right">Show on:</td>
			<td>New customer checkout 
				<input type="checkbox" name="RefNewCheckout" value="1" <%if RefNewCheckout="1" then%>checked<%end if%> class="clearBorder">
            </td>
		</tr>
		<tr> 
			<td>New customer registration 
				<input type="checkbox" name="RefNewReg" value="1" <%if RefNewReg="1" then%>checked<%end if%> class="clearBorder">
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="guest"></a>Guest Checkout Options&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=466')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">These settings pertain to a customer's ability to checkout without entering a password. If - for any reason - you need e-mails to uniquely identify one record in the &quot;customers&quot; table, you must choose the third option.
            	<table cellpadding="4" cellspacing="0" width="100%" style="margin-top: 10px; margin-bottom: 15px;">
                    <tr>
                        <td align="right" valign="top" width="5%">
                                <input type="radio" name="GuestCheckoutOpt" value="0" <%if pcIntGuestCheckoutOpt="0" OR pcIntGuestCheckoutOpt="" then%>checked<%end if%> class="clearBorder">
                        </td>
                        <td valign="top" width="95%">Guest Checkout enabled<br>
                        <em>Customers can checkout as guests (default behavior in ProductCart v4.0).</em></td>
                    </tr>
                    <tr>
                        <td align="right" valign="top">
                                <input type="radio" name="GuestCheckoutOpt" value="1" <%if pcIntGuestCheckoutOpt="1" then%>checked<%end if%> class="clearBorder">
                        </td>
                        <td valign="top">Guest Checkout allowed<br>
                        <em>Customers can checkout as guests, but by default they are asked to register (and they can opt out).</em></td>
                    </tr>
                    <tr>
                        <td align="right" valign="top">
                                <input type="radio" name="GuestCheckoutOpt" value="2" <%if pcIntGuestCheckoutOpt="2" then%>checked<%end if%> class="clearBorder">
                        </td>
                        <td valign="top">Guest Checkout disabled<br>
                        <em>Customers cannot checkout as guests. They must enter a password and use a unique e-mail address.</em></td>  
                    </tr>
                </table>
             </td>
        </tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="terms"></a>Terms &amp; Conditions Agreement</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">You can ask your customers to agree to your &quot;Terms &amp; Conditions&quot; before they can checkout. They are asked to do so when placing the first order (or registering with the store), and not again. When shown, this field is always required.
          	</td>
		</tr>
		<tr> 
			<td align="right">Enable Agreement ?</td>
			<td>
					<input type="radio" name="Terms" value="1" <%if pcIntTerms="1" then%>checked<%end if%> class="clearBorder">Yes
					<input type="radio" name="Terms" value="0" <%if pcIntTerms<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td align="right">Field Description:</td>
			<td align="left">
				<input name="TermsLabel" type="text" value="<%=pcStrTermsLabel%>" size="35">
				(e.g. &quot;I agree to the Terms &amp; Conditions shown above&quot;)
            </td>
		</tr>
		<tr> 
			<td align="right" valign="top">Content of the Agreement:</td>
			<td><textarea name="TermsCopy" cols="60" rows="6"><%=pcStrTermsCopy%></textarea></td> 
		</tr>
		<tr> 
			<td align="right">Require customers to agree:</td>
			<td>
					<input type="radio" name="TermsShown" value="1" <%if pcIntTermsShown="1" then%>checked<%end if%> class="clearBorder">
					Everytime they purchase&nbsp;
					<input type="radio" name="TermsShown" value="0" <%if pcIntTermsShown<>"1" then%>checked<%end if%> class="clearBorder">
					First time only
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="newsletter"></a>Newsletter Settings</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">You can ask your customers whether or not they would like to receive information from the store. Use the Newsletter Wizard to send a message to selected customers. For more powerful and flexible e-mail marketing features, check out the <a href="mailup_home.asp">MailUp</a> integration.</td>
		</tr>
		<tr> 
			<td align="right">Field Description:</td>
			<td align="left">
				<input name="NewsLabel" type="text" value="<%=NewsLabel%>" size="35">
				(e.g. Receive information about new products)
            </td>
		</tr>
		<tr> 
			<td align="right">Show this field?</td>
			<td>
				<input type="radio" name="AllowNews" value="1" <%if AllowNews="1" then%>checked<%end if%> class="clearBorder">Yes
				<input type="radio" name="AllowNews" value="0" <%if AllowNews<>"1" then%>checked<%end if%> class="clearBorder">No</td>
		</tr>
		<tr> 
			<td rowspan="2" align="right">Show on:</td>
			<td>New customer checkout 
				<input type="checkbox" name="NewsCheckout" value="1" <%if NewsCheckout="1" then%>checked<%end if%> class="clearBorder">
            </td>
		</tr>
		<tr> 
			<td>New customer registration 
				<input type="checkbox" name="NewsReg" value="1" <%if NewsReg="1" then%>checked<%end if%> class="clearBorder">
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="datetime"></a>Custom Date Field</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">You can ask your customers when the order should be delivered (e.g. catering business). Customers will select a date using a pop-up calendar.</td>
		</tr>
		<tr> 
			<td align="right">Date Field description:</td>
			<td><input name="DFLabel" type="text" value="<%=DFLabel%>" size="35" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right">Show this field?</td>
			<td>
				<input type="radio" name="DFShow" value="1" <%if DFShow="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="DFShow" value="0" <%if DFShow<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td align="right">Field required?</td>
			<td>
				<input type="radio" name="DFReq" value="1" <%if DFReq="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="DFReq" value="0" <%if DFReq<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>If you wish to set dates when delivery would be unavailable (Blackout Dates), please <a href="blackout_main.asp">click here for the Blackout Date manager</a>.</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td align="right">Time Field description:</td>
			<td><input name="TFLabel" type="text" value="<%=TFLabel%>" size="35" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right">Show this field?</td>
			<td>
				<input type="radio" name="TFShow" value="1" <%if TFShow="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="TFShow" value="0" <%if TFShow<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td align="right">Field required?</td>
			<td>
				<input type="radio" name="TFReq" value="1" <%if TFReq="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="TFReq" value="0" <%if TFReq<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td align="right">Date must be 24 hours in the future?</td>
			<td>
				<input type="radio" name="DTCheck" value="1" <%if DTCheck="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="DTCheck" value="0" <%if DTCheck<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="zipcodes"></a>Limit Delivery Area By Zip Code</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">By turning this option on, customers will only have the option of shipping/delivering items to a list of accepted zip codes. You can create this list by using the <a href="DeliveryZipCodes_main.asp">Delivery Zip Codes</a> manager.</td>
		</tr>
		<tr> 
			<td align="right">Limit Delivery Area by Zip Code?</td>
			<td>
				<input type="radio" name="DeliveryZip" value="1" <%if DeliveryZip="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="DeliveryZip" value="0" <%if DeliveryZip<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2"><a name="CustomerIP"></a>Customer IP Alert</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2">ProductCart saves the Customer's IP address during checkout. You can alert the customer that you are saving their IP address on the payment page.</td>
		</tr>
		<tr> 
			<td align="right">Show customer alert message?</td>
			<td>
				<input type="radio" name="CustomerIPAlert" value="1" <%if CustomerIPAlert="1" then%>checked<%end if%> class="clearBorder">Yes 
				<input type="radio" name="CustomerIPAlert" value="0" <%if CustomerIPAlert<>"1" then%>checked<%end if%> class="clearBorder">No
            </td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center">
				<input name="Submit1" type="submit" value="Update" class="submit2"> 
				&nbsp;
				<input name="back" type="button" onClick="javascript:history.back()" value="Back">
            </td>
		</tr>
		</table>
	</form>
<!--#include file="AdminFooter.asp"-->