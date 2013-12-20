<%@ LANGUAGE="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "viewcart.asp"
' This page displays the items in the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2013. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------

'Clear One Page Checkout progress
session("OPCstep")=0
session("pcPay_PayPalExp_OrderTotal")=""

%>
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/bto_language.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/opendb.asp"--> 
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"-->  

<%
Response.Buffer = True
Dim query, conntemp, rs, rstemp

'SB S
Dim pcIsSubscription
'SB E

'*****************************************************************************************************
' START: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*****************************************************************************************************
' END: Check store on/off, start PC session, check affiliate ID
'*****************************************************************************************************

%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<!--#include file="inc_PrdCatTip.asp"-->
<!--#include file="pcCheckPricingCats.asp"-->
<%

'*******************************
' Display settings
'*******************************
' 1) The following variable controls the size of the small image shown in the cart
' Move to the Control Panel in a future version

Dim pcIntSmImgWidth
pcIntSmImgWidth = 35

' 2) The following varaible controls whether the SKU is shown in the cart or not.
' Move to the Control Panel in a future version

Dim pcIntShowSku
pcIntShowSku = 1 ' Change to 0 to hide the SKU

'*****************************************************************************************************
' START: PAGE ON LOAD
'*****************************************************************************************************

session("availableShipStr")=""
session("provider")=""

If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

'// Express Checkout
if Request("cmd")="_express-checkout" then
	session("ExpressCheckoutPayment")=""	
end if

dim f, total, totalDeliveringTime
call opendb()
%><!--#include file="inc_SaveShoppingCart.asp"--><%
call closedb()
'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************

total=Cint(0)
totalDeliveringTime=Cint(0)

if countCartRows(pcCartArray, pcCartIndex)=0 then
 	response.redirect "msg.asp?message=1"     
end if

call opendb()

Dim strCCSLCheck
strCCSLcheck = checkCartStockLevels(pcCartArray, pcCartIndex, aryBadItems)

If Len(Trim(strCCSLCheck))>0 Then
   response.write "<div class=""pcErrorMessage"">"
   response.write dictLanguage.Item(Session("language")&"_alert_19") & strCCSLcheck
   response.write "</div>"
End If


'// Duplicate Order Validation
If len(session("GWOrderID"))>0 Then
	Dim pcv_intOrderID
	pcv_intOrderID = session("GWOrderID")
	pcv_intOrderID = pcv_intOrderID-int(scPre)
	query="SELECT orderStatus FROM orders WHERE orderStatus>1 AND idOrder="&pcv_intOrderID&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	If NOT rs.eof Then
		set rs=nothing
		call closedb()
		session("SFClearCartURL")="msg.asp?message=308"
		response.Redirect("CustLOb.asp")
	End If
	set rs=nothing
End If


'**********************************************
'// START: Google Checkout and PayPal Express
'**********************************************
pcv_strShowCheckoutBtn=pcf_PaymentTypes("")
'**********************************************
'// END: Google Checkout and PayPal Express
'**********************************************

'see if there are any ship types setup for this store
dim iShipService
iShipService=0
query="SELECT * FROM shipService WHERE serviceActive=-1;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rs.eof then
	iShipService=1
end if
set rs=nothing

'Calculate Product Promotions - START
%>
<!--#include file="inc_CalPromotions.asp"-->
<%
'Calculate Product Promotions - END
'SB S
pcIsSubscription = False
'SB E

'//***************************************
'// START - Load Gift Wrapping settings
'//***************************************

query="select pcGWSet_Show,pcGWSet_OverviewCart,pcGWSet_HTMLCart from pcGWSettings"
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rstemp.eof then
	pcv_GW=rstemp("pcGWSet_Show")
	if pcv_GW="0" then
		pcv_GW=""
	end if
	session("Cust_GW")=pcv_GW
	pcv_Overview=rstemp("pcGWSet_OverviewCart")
	if pcv_Overview="0" then
		pcv_Overview=""
	end if
	session("Cust_GWText")=pcv_Overview
	pcv_GWDetails=rstemp("pcGWSet_HTMLCart")
	if trim(pcv_GWDetails)<>"" then pcv_GWDetails=replace(pcv_GWDetails,"&quot;",chr(34))
else
	session("Cust_GW")=""
	session("Cust_GWText")=""
end if
set rstemp=nothing

'//***************************************
'// END - Load Gift Wrapping settings
'//***************************************




'*****************************************************************************************************
' END: PAGE ON LOAD
'*****************************************************************************************************
%>

<script language="JavaScript">
	var RemainIssue="";
	var RemainIssue1="";
	
	function checkQtyChange()
	{
			var i=0;
			for (i=1;i<=<%=pcCartIndex%>;i++)
			{
				if (eval("document.recalculate.Cant" + i).value != eval("document.recalculate.SavQty" + i).value)
				{
					alert("<%response.write dictLanguage.Item(Session("language")&"_alert_recal")%>");
					return(false);
				}
			}
			return(true);
		
	}
	
	function win(fileName)
		{
			myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=300,height=250')
			myFloater.location.href=fileName;
		}

	function winShipPreview(fileName)
		{
		myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=520,height=400')
		myFloater.location.href=fileName;
		}
		
	function validateNumber(field)
	{
		var val=field.value;
		if(!/^\d*$/.test(val)||val==0)
		{
				alert("<%response.write dictLanguage.Item(Session("language")&"_showcart_2")%>");
				field.focus();
				field.select();
		}
	}

	function isDigit(s)
	{
	var test=""+s;
	if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
	{
	return(true) ;
	}
	return(false);
	}

	function allDigit(s)
	{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (isDigit(c)==false)
	{
	return (false);
	}
	}
	return (true);
	}

	<%'GGG Add-on start%>
	function checkproqty(fname,qty,ctype,remain,MultiQty)
	<%'GGG Add-on end%>
	{
	RemainIssue1="A";
	if (fname.value == "")
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
		fname.focus();
		return (false);
		}
	if (allDigit(fname.value) == false)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
		if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
		fname.focus();
		return (false);
	}
	if (fname.value == "0")
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
		if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
		fname.focus();
		return (false);
	}
	
		if (ctype > 0)
		{
		TempValue=eval(fname.value);
			TempV1=(TempValue/MultiQty);
		TempV1a=TempValue*TempV1;
			TempV2=Math.round(TempValue/MultiQty);
		TempV2a=TempValue*TempV2;
		if ((TempV1a != TempV2a) || (TempV1<1))
			{
				alert("<% Response.write(dictLanguage.Item(Session("language")&"_alert_3"))%>" + MultiQty);
			if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
			fname.focus();
			return (false);
			}
		}
		if (qty > 0)
		{
		TempValue=eval(fname.value);
		if (TempValue < qty)
			{
			alert("<% Response.write(dictLanguage.Item(Session("language")&"_alert_8"))%>"+qty+"<% Response.write(dictLanguage.Item(Session("language")&"_alert_9"))%>");
			if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
			fname.focus();
			return (false);
		}
	}
	
	<%'GGG Add-on start
	if Session("Cust_IDEvent")<>"" then%>
		if ((eval(fname.value) > remain) && (remain > 0))
		{
		    alert("Your entered a quantity greater than remaining quantity.");
		    if (RemainIssue.indexOf('||' + fname.name)==-1) RemainIssue=RemainIssue + '||' + fname.name;
		    fname.focus();
		    return (false);
		}
	<%end if
	'GGG Add-on end%>
	if (RemainIssue!="") RemainIssue=RemainIssue.replace('||' + fname.name,'');
	if (RemainIssue1=="A") RemainIssue1="";
	return (true);
	}

/***********************************************
* Disable "Enter" key in Form script- By Nurul Fadilah(nurul@REMOVETHISvolmedia.com)
* This notice must stay intact for use
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/
                
function handleEnter (field, event) {
		var keyCode = event.keyCode ? event.keyCode : event.which ? event.which : event.charCode;
		if (keyCode == 13) {
			var i;
			for (i = 0; i < field.form.elements.length; i++)
				if (field == field.form.elements[i])
					break;
			i = (i + 1) % field.form.elements.length;
			field.form.elements[i].focus();
			return false;
		} 
		else
		return true;
	}      
</script>

<% 
Dim strRedirectSSL
strRedirectSSL="onepagecheckout.asp"
if scSSL="1" AND scIntSSLPage="1" then
	strRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/pc/onepagecheckout.asp"),"//","/")
	strRedirectSSL=replace(strRedirectSSL,"https:/","https://")
	strRedirectSSL=replace(strRedirectSSL,"http:/","http://")
end if
%>

<div id="pcMain">
<% '// START main form %>
<form method="post" action="cRec.asp" name="recalculate" class="pcForms" onsubmit="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);}">
	<table class="pcMainTable">
		<tr> 
			<td>
				<h1><%=dictLanguage.Item(Session("language")&"_showcart_3")%></h1>
			</td>
		</tr>
		<tr>
			<td align="right">
                	
				<% if session("idcustomer")=0 then %>

                    <a href="javascript:location='checkout.asp?cmode=1';" id="save-cart" name="save-cart"><img src="<%=RSlayout("pcLO_Savecart")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_SaveCart_1")%>"></a>&nbsp;

                <% else %>

					<div id="dialog-form" title="<%=dictLanguage.Item(Session("language")&"_SaveCart_1")%>">
                        <div align="left">
                            <p class="validateTips"><%=dictLanguage.Item(Session("language")&"_SaveCart_2")%></p>
                            <div id="saved_cart_success"></div>
							<div>
                            <%=dictLanguage.Item(Session("language")&"_SaveCart_3")%>:  <input type="text" name="SavedCartName" id="SavedCartName" />
							</div>
                        </div>
                    </div>                        
                    <a href="javascript:;" id="save-cart" name="save-cart"><img src="<%=RSlayout("pcLO_Savecart")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_SaveCart_1")%>"></a>&nbsp;
                    <script language="JavaScript">
                        $(function() {
                                   
                            var SavedCartName = $( "#SavedCartName" ),
                                allFields = $( [] ).add( SavedCartName ),
                                tips = $( ".validateTips" );
                        
                            function updateTips( t ) {
                                tips
                                    .text( t )
                                    .addClass( "pcErrorMessage" );
                            }
                        
                            function checkLength( o, n, min, max ) {
                                if ( o.val().length > max || o.val().length < min ) {
                                    o.addClass( "ui-state-error" );
                                    updateTips( "Length of " + n + " must be between " +
                                        min + " and " + max + "." );
                                    return false;
                                } else {
                                    return true;
                                }
                            }
                        
                            function checkRegexp( o, regexp, n ) {
                                if ( !( regexp.test( o.val() ) ) ) {
                                    o.addClass( "ui-state-error" );
                                    updateTips( n );
                                    return false;
                                } else {
                                    return true;
                                }
                            }
                            
                            $( "#dialog-form" ).dialog({
                                autoOpen: false,
                                height: 300,
                                width: 350,
                                modal: true,
                                buttons: {
                                    "Save Cart": function() {
                                        var bValid = true;
                                        allFields.removeClass( "ui-state-error" );
                                        bValid = bValid && checkLength( SavedCartName, "<%=dictLanguage.Item(Session("language")&"_SaveCart_3")%>", 3, 100 );
                                        bValid = bValid && checkRegexp( SavedCartName, /^[a-z]([0-9a-z_ ])+$/i, "<%=dictLanguage.Item(Session("language")&"_SaveCart_4")%>" );
                                        if ( bValid ) {
                                            
                                            $.ajax(
                                                   {
                                                    type: "GET",
                                                    url: "viewcart.asp?SaveCart=1&SavedCartName=" + SavedCartName.val(),
                                                    data: "{}",
                                                    timeout: 45000,
                                                    success: function(data, textStatus){
														$("#saved_cart_success").html('<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_SaveCart_5")%></div>');
														setTimeout(function(){$( "#dialog-form" ).dialog( "close" )},2000);
                                                    }
                                            });
                                        }
                                    },
                                    Cancel: function() {
                                        $( this ).dialog( "close" );
                                    }
                                },
                                close: function() {
                                    allFields.val( "" ).removeClass( "ui-state-error" );
                                }
                            });
                        
                            $('#save-cart').click(function() {
                                    $( "#dialog-form" ).dialog( "open" );
                                });
                        });
                    </script>
                    
                <% end if %>

				<%'GGG Add-on start
				if Session("Cust_BuyGift")<>"" then
		  		query="select pcEv_Code from pcEvents where pcEv_IDEvent=" & session("Cust_IDEvent")
		  		set rsG=conntemp.execute(query)
		  		grCode=rsG("pcEv_Code")%>
		  			<a href="ggg_viewGR.asp?grcode=<%=grCode%>"><img src="<%=RSlayout("RetRegistry")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_10")%>"></a>
		  		<%else%>
					<a href="default.asp"><img src="<%=RSlayout("continueshop")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_11")%>"></a>
				<%end if
				'GGG Add-on end%>
				
				<%'GGG Add-on start
				HaveGcsTest=0
				for f=1 to pcCartIndex
					if pcCartArray(f,10)=0 then
						query="select pcprod_Gc from Products where idproduct=" & pcCartArray(f,0) & " AND pcprod_Gc=1"
						set rsGc=conntemp.execute(query)
						if not rsGc.eof then
							HaveGcsTest=1
							exit for
						end if
					end if
				next%>
				<%if HaveGcsTest=1 then %>
						<% if pcv_strShowCheckoutBtn=1 then %>
							&nbsp;<input type="image" id="submit" name="image1" value="Gift Certificates" src="<%=RSlayout("checkout")%>" border="0" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);} else {document.recalculate.actGCs.value='Hello'; return(true);}">
							<input type=hidden name="actGCs" value="">
						<% end if %>
               		<%else%>
						<% if pcv_strShowCheckoutBtn=1 then %>
							&nbsp;<a href="<%=strRedirectSSL%>" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);} return(checkQtyChange());"><img src="<%=RSlayout("checkout")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_12")%>"></a>
						<% end if %>
					<%end if
				'GGG Add-on end%>
			</td>
		</tr>
	
	<%
	' ------------------------------------------------------
	'Start SDBA - Notify Drop-Shipping
	' ------------------------------------------------------
		if scShipNotifySeparate="1" and pcCartIndex>1 then
			tmp_showmsg=0
			for f=1 to pcCartIndex
				tmp_idproduct=pcCartArray(f,0)
				query="SELECT pcProd_IsDropShipped FROM products WHERE idproduct=" & tmp_idproduct & " AND pcProd_IsDropShipped=1;"
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rs.eof then
					tmp_showmsg=1
					exit for
				end if
				set rs=nothing
			next
			if tmp_showmsg=1 then%>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
			<tr>
				<td>
					<div class="pcTextMessage"><img src="images/sds_boxes.gif" alt="<%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>" align="left" vspace="5" hspace="10"><%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%></div>
				</td>
			</tr>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
			<%end if
		end if
	' ------------------------------------------------------
	' End SDBA - Notify Drop-Shipping
	' ------------------------------------------------------
	%>
	
	<%
	' ------------------------------------------------------
	' Start Cross Selling - Notify Accessory Added
	' ------------------------------------------------------
		if Session("cs_Accessory") <> "" then 
			cs_Msg = replace(dictLanguage.Item(Session("language")&"_showcart_28"),"<main product name>", Session("cs_Accessory"))
			%>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
			<tr>
				<td>
					<div class="pcTextMessage">
						<img src="images/note.gif" alt="<%response.write cs_Msg %>" align="left" vspace="5" hspace="10">
						<%response.write cs_Msg %>
					</div>
				</td>
			</tr>
			<tr> 
				<td class="pcSpacer"></td>
			</tr>
			<%
			Session("cs_Accessory") = ""
		end if

	' ------------------------------------------------------
	' End Cross Selling - End Accessory Added
	' ------------------------------------------------------
	%>


		<tr> 
			<td>
			<table class="pcShowCart">
			<% '// START 1st Row - Table Headers %>
				<tr> 
					<th width="6%"> 
						<div style="text-align: left;"><%response.write dictLanguage.Item(Session("language")&"_showcart_4")%></div>
					</th>
					<th width="57%">
						<div style="text-align: left;"><%response.write dictLanguage.Item(Session("language")&"_showcart_6")%></div>
					</th>
					<th width="12%" nowrap>
						<div style="text-align: left;"><%response.write dictLanguage.Item(Session("language")&"_showcart_8b")%></div>
					</th>
					<th width="12%">
						<div style="text-align: left;"><%response.write dictLanguage.Item(Session("language")&"_showcart_8")%></div>
					</th>
					<th nowrap>
						<%if session("Cust_GW")="1" then%>
							<div style="text-align: center;"><%response.write dictLanguage.Item(Session("language")&"_showcart_24")%></div>
						<%end if%>
                    </th>
					<th width="13%">&nbsp;</th>
				</tr>
				<tr>
					<td class="pcSpacer" colspan="6"></td>
				</tr>
			<% '// END 1st Row - Table Headers %>
				<% dim totalRowWeight
				totalRowWeight=0
				
				pcv_SpecialServer=0
				
				if InStr(Cstr(10/3),",")>0 then
					pcv_SpecialServer=1
					for f=1 to pcCartIndex
						for n=0 to 34
							if Instr(pcCartArray(f,n),".")>0 then
								if IsNumeric(pcCartArray(f,n)) then
									pcCartArray(f,n)=replace(pcCartArray(f,n),".",",")
								end if
							end if
						next
					next
				else
					if scDecSign="," then
						pcv_SpecialServer=0
						for f=1 to pcCartIndex
							for n=0 to 34
								if Instr(pcCartArray(f,n),",")>0 then
									if IsNumeric(pcCartArray(f,n)) then
										pcCartArray(f,n)=replace(pcCartArray(f,n),",",".")
									end if
								end if
							next
						next
					end if
				end if
				
				Dim ProList(100,5)
				
				for f=1 to pcCartIndex
					ProList(f,0)=pcCartArray(f,0)
					ProList(f,1)=pcCartArray(f,10)
					ProList(f,3)=pcCartArray(f,2)
					ProList(f,4)=0
					if pcCartArray(f,10)=0 then%>
				<% '// START 2nd Row - Main Product Data %>
						<tr> 
							<td> 
								<% 'Validate for multiple of N
								query="select serviceSpec,pcprod_HideBTOPrice,pcprod_QtyValidate,pcprod_MinimumQty,pcProd_multiQty from products where idproduct=" & pcCartArray(f,0) 
								set rs=server.CreateObject("ADODB.RecordSet")									
								set rs=connTemp.execute(query)
								if err.number<>0 then
									'//Logs error to the database
									call LogErrorToDatabase()
									'//clear any objects
									set rs=nothing
									'//close any connections
									call closedb()
									'//redirect to error page
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								IsBTO=rs("serviceSpec")
								pcv_intHideBTOPrice=rs("pcprod_HideBTOPrice")
								if isNULL(pcv_intHideBTOPrice) OR pcv_intHideBTOPrice="" then
									pcv_intHideBTOPrice="0"
								end if
								pcv_intQtyValidate=rs("pcprod_QtyValidate")
								if isNULL(pcv_intQtyValidate) OR pcv_intQtyValidate="" then
									pcv_intQtyValidate="0"
								end if				
								pcv_lngMinimumQty=rs("pcprod_MinimumQty")
								if isNULL(pcv_lngMinimumQty) OR pcv_lngMinimumQty="" then
									pcv_lngMinimumQty="0"
								end if
								pcv_lngMultiQty=rs("pcProd_multiQty")
								if isNULL(pcv_lngMultiQty) OR pcv_lngMultiQty="" then
									pcv_lngMultiQty="0"
								end if 
								set rs=nothing %>
								<%'GGG Add-on start
								gRemain=0
								if (pcCartArray(f,33)<>"") and (pcCartArray(f,33)<>"0") then
									query="select pcEP_Qty,pcEP_HQty from pcEvProducts where pcEP_ID=" & pcCartArray(f,33)
									set rsG=connTemp.execute(query)
									if not rsG.eof then
										gRemain=cdbl(rsG("pcEP_Qty"))-cdbl(rsG("pcEP_HQty"))
									end if
									set rsG=nothing
								end if
								'GGG Add-on end%>
								
								<% '// Set Quantity field to transparent if it's a child (Cross Sell) or it's a Finalized Quote
								pcv_FinalizedQuote=0
								if Instr(session("sf_FQuotes"),"****" & pcCartArray(f,0) & "****")>0 then
									pcv_FinalizedQuote=1
								end if
								
								if trim(pcCartArray(f,27))="" then
									pcCartArray(f,27)=0
								end if
								if trim(pcCartArray(f,28))="" then
									pcCartArray(f,28)=0
								end if

								'SB S
								If (pcCartArray(f,38)>0) then
									pcIsSubscription = True
									'// Get the data 
								    pSubscriptionID = (pcCartArray(f,38))
									%>
									<!--#include file="../includes/pcSBDataInc.asp" --> 
									<%								
								End if 
								'SB E

								if (pcCartArray(f,27)>0) AND (pcCartArray(f,28)>0) OR (pcv_FinalizedQuote=1) OR (pcCartArray(f,38)>0) then %>
								    <input type="text" name="Cant<%=f%>" size="3" value="<% response.write pcCartArray(f,2) %>" class="transparentField" <% If CLng(aryBadItems(f-1))<>0 Then response.write " style=""background-color: #fcc;"">" %> readonly>
								<% else %>
								    <input type="text" name="Cant<%=f%>" size="3" value="<% response.write pcCartArray(f,2) %>" onBlur="checkproqty(this,<%=pcv_lngMinimumQty%>,<%if pcv_intQtyValidate<>"1" then%>0<%else%>1<%end if%>,<%if session("Cust_IDEvent")<>"" then%><%=gRemain%><%else%>0<%end if%>,<%=pcv_lngMultiQty%>)" onkeypress="return handleEnter(this, event)" <% If CLng(aryBadItems(f-1))<>0 Then response.write " style=""background-color: #fcc;""" %>>
								<% 
								end if 
								%>
								<input type="hidden" name="SavQty<%=f%>" value="<% response.write pcCartArray(f,2) %>" />
								
							</td>
							
							<% ' Get product image and sku
									query="SELECT sku,smallImageUrl FROM products WHERE idProduct=" & pcCartArray(f,0)
									set rsImg=Server.CreateObject("ADODB.Recordset")
									set rsImg=conntemp.execute(query)
									if err.number<>0 then
										'//Logs error to the database
										call LogErrorToDatabase()
										'//clear any objects
										set rsImg=nothing
										'//close any connections
										call closedb()
										'//redirect to error page
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									
									pcvStrSku = rsImg("sku")
									pcvStrSmallImage = rsImg("smallImageUrl")
									if pcvStrSmallImage = "" or pcvStrSmallImage = "no_image.gif" then
										pcvStrSmallImage = "hide"
									end if
									set rsImg = nothing
									' End get product image
							%>
							
							<td>
								
								
								<table width="100%" cellpadding="0" cellspacing="0" border="0">
									<tr>
										<td>
										<%
										'// Below we setup the image and product name links
										'// If this is a Gift Registry then we use different links
										%>	
										<% If Session("Cust_BuyGift")="" Then %>
											<% if pcvStrSmallImage = "hide" then %>
												&nbsp;
											<% else %>
												<a href="viewPrd.asp?idproduct=<%=pcCartArray(f,0)%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pcCartArray(f,0)%>'; sav_callxml='1'; return runXML1('prd_<%=pcCartArray(f,0)%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>>
												<img src="catalog/<%=pcvStrSmallImage%>" hspace="2" width="<%=pcIntSmImgWidth%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")& pcCartArray(f,1)%>">
												</a>
											<% end if %>
										<% Else %>
											<% if pcvStrSmallImage = "hide" then %>
												&nbsp;
											<% else %>
												<a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=pcCartArray(f,33)%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pcCartArray(f,0)%>'; sav_callxml='1'; return runXML1('prd_<%=pcCartArray(f,0)%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>>
												<img src="catalog/<%=pcvStrSmallImage%>" hspace="2" width="<%=pcIntSmImgWidth%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")& pcCartArray(f,1)%>">
												</a>
											<% end if %>
										<% End If %>	
										</td>
										<td width="100%" align="left">											
										<% If Session("Cust_BuyGift")="" Then %>											
											<a href="viewPrd.asp?idproduct=<%=pcCartArray(f,0)%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pcCartArray(f,0)%>'; sav_callxml='1'; return runXML1('prd_<%=pcCartArray(f,0)%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%response.write pcCartArray(f,1)%></a> <span class="pcSmallText">(<%=pcvStrSku%>)</span>											
										<% Else %>										
											<a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=pcCartArray(f,33)%>" <%if scStoreUseToolTip="1" or scStoreUseToolTip="2" then%>onmouseover="javascript:document.getPrd.idproduct.value='<%=pcCartArray(f,0)%>'; sav_callxml='1'; return runXML1('prd_<%=pcCartArray(f,0)%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%response.write pcCartArray(f,1)%></a> <span class="pcSmallText">(<%=pcvStrSku%>)</span>											
										<% End If %>											
										</td>
									</tr>
								</table>
								
							</td>
							
							<% 'BTO ADDON-S
							pBTOValues=0
							if trim(pcCartArray(f,16))<>"" then 

								query="SELECT stringProducts, stringValues, stringCategories, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=conntemp.execute(query)
								if err.number<>0 then
									'//Logs error to the database
									call LogErrorToDatabase()
									'//clear any objects
									set rs=nothing
									'//close any connections
									call closedb()
									'//redirect to error page
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								stringProducts=rs("stringProducts")
								stringValues=rs("stringValues")
								stringCategories=rs("stringCategories")
								ArrProduct=Split(stringProducts, ",")
								ArrValue=Split(stringValues, ",")
								ArrCategory=Split(stringCategories, ",")
								Qstring=rs("stringQuantity")
								ArrQuantity=Split(Qstring,",")
								Pstring=rs("stringPrice")
								ArrPrice=split(Pstring,",")
								set rs=nothing
								
								if ArrProduct(0)="na" then
								else
									for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
										query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
										set rsQ=connTemp.execute(query)
										tmpMinQty=1
										if not rsQ.eof then
											tmpMinQty=rsQ("pcprod_minimumqty")
											if IsNull(tmpMinQty) or tmpMinQty="" then
												tmpMinQty=1
											else
												if tmpMinQty="0" then
													tmpMinQty=1
												end if
											end if
										end if
										set rsQ=nothing
										tmpDefault=0
										query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
										set rsQ=connTemp.execute(query)
										if not rsQ.eof then
											tmpDefault=rsQ("cdefault")
											if IsNull(tmpDefault) or tmpDefault="" then
												tmpDefault=0
											else
												if tmpDefault<>"0" then
												 	tmpDefault=1
												end if
											end if
										end if
										set rsQ=nothing
											
										if pcv_SpecialServer=1 then
											ArrValue(i)=replace(ArrValue(i),".",",")
											ArrPrice(i)=replace(ArrPrice(i),".",",")
										end if
										if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
											if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
												if tmpDefault=1 then
													UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
												else
													UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
												end if
											else
												UPrice=0
											end if
											pBTOValues=pBTOValues+ccur((ArrValue(i)+UPrice)*pcCartArray(f,2))
										end if
										set rs=nothing
									next
								end if
														
							End if
							'BTO ADDON-E %>
							
							<% dim pRowPrice, pRowWeight, pExtRowPrice
							pRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))
							pExtRowPrice=ccur(pcCartArray(f,2) * pcCartArray(f,17))-ccur(pBTOvalues)
							if pcCartArray(f,20)=0 then
								pRowWeight=pcCartArray(f,2)*pcCartArray(f,6)
							else
								pRowWeight=0
							end if
							totalRowWeight=totalRowWeight+pRowWeight %>
	
									
							<td>
                            	<div class="pcShowCartPrices">
								<%if pcv_intHideBTOPrice<>"1" then%>
                                    <% if pcCartArray(f,17) > 0 then %>
                                    <%=scCurSign & money(pcCartArray(f,17)-ccur(ccur(pBTOvalues)/pcCartArray(f,2)))%>
                                    <% end if %>
                                <%end if%>
                                </div>
							</td>
							<td>
                                <div class="pcShowCartPrices">
								<% if pExtRowPrice > 0 then response.write(scCurSign &  money(pExtRowPrice)) end if %>
                                </div>
							</td>
							<td>
                                <div style="text-align:center;">
								<%
                                if session("Cust_GW")="1" then
									PrdCanGWchecks=0
									query="SELECT pcGC_EOnly FROM pcGC WHERE pcGC_idproduct=" & pcCartArray(f,0)
									set rs1=connTemp.execute(query)
									if rs1.eof then
										query="SELECT pcPE_IDProduct FROM pcProductsExc WHERE pcPE_IDProduct=" & pcCartArray(f,0)
										set rs1=connTemp.execute(query)
										if rs1.eof then
											PrdCanGWchecks=1
										else 
											PrdCanGWchecks=0
										end if
									end if
									if PrdCanGWchecks=1 then
                                        response.write dictLanguage.Item(Session("language")&"_showcart_30")
                                    else
                                        response.write dictLanguage.Item(Session("language")&"_showcart_25")
                                    end if
                                end if
                                %>
                                </div>
                            </td>
							<td>
                            	<div style="text-align:right;">
								<% '// Show Remove Button if it's NOT a Required child (Cross Sell) %>
               	<% if (pcCartArray(f,12)<>"-2") then %>
								    <a href="cRemv.asp?pcCartIndex=<%response.write f%>"><img src="<%=RSlayout("remove")%>" alt="<%=dictLanguage.Item(Session("language")&"_altTag_13")%>"></a>
								<% end if %>
							
								<%if (IsBTO=-1) and pcCartArray(f,16)="" and pcv_FinalizedQuote=0 then
									call opendb()
									queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & ";"
									set rsQ=connTemp.execute(queryQ)
									if err.number<>0 then
										'//Logs error to the database
										call LogErrorToDatabase()
										'//clear any objects
										set rsQ=nothing
										'//close any connections
										call closedb()
										'//redirect to error page
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									if not rsQ.eof then%>
										<br>
										<a href="Reconfigure.asp?pcCartIndex=<%=f%>"><img src="<%=rslayout("reconfigure")%>" alt="Reconfigure <%=pcCartArray(f,1)%>"></a>
									<%End if
									set rsQ=nothing%>
								<%end if%>
                                </div>
							</td>
						</tr>
						<% '// END 2nd Row - Main Product Data %>
					
						<% 'BTO ADDON-S
						if trim(pcCartArray(f,16))<>"" then 
							query="SELECT stringProducts, stringValues, stringCategories, stringQuantity, stringPrice FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							if err.number<>0 then
								'//Logs error to the database
								call LogErrorToDatabase()
								'//clear any objects
								set rs=nothing
								'//close any connections
								call closedb()
								'//redirect to error page
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							stringProducts=rs("stringProducts")
							stringValues=rs("stringValues")
							stringCategories=rs("stringCategories")
							ArrProduct=Split(stringProducts, ",")
							ArrValue=Split(stringValues, ",")
							ArrCategory=Split(stringCategories, ",")
							Qstring=rs("stringQuantity")
							ArrQuantity=Split(Qstring,",")
							Pstring=rs("stringPrice")
							ArrPrice=split(Pstring,",")
							set rs=nothing %>
							<% '// START 3nd Row - BTO Product Details %>
							<tr valign="top"> 
								<td>&nbsp;</td>
								<td colspan="3"> 
									<div class="pcShowBTOconfiguration">
										<% if ArrProduct(0)="na" then %>
										<p>
											<%response.write bto_dictLanguage.Item(Session("language")&"_viewcart_2")%>
										</p>
										<% else %>
										<p>
											<strong><%response.write bto_dictLanguage.Item(Session("language")&"_viewcart_1")%></strong>
										</p>
											<% for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
												if pcv_SpecialServer=1 then
													ArrValue(i)=replace(ArrValue(i),".",",")
													ArrPrice(i)=replace(ArrPrice(i),".",",")
												end if
												query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCategory(i)&") AND ((products.idProduct)="&ArrProduct(i)&"))" 
												set rs=server.CreateObject("ADODB.RecordSet")
												set rs=conntemp.execute(query)
												if err.number<>0 then
													'//Logs error to the database
													call LogErrorToDatabase()
													'//clear any objects
													set rs=nothing
													'//close any connections
													call closedb()
													'//redirect to error page
													response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												strCategoryDesc=rs("categoryDesc")
												strDescription=rs("description")
												set rs=nothing
												
												query="SELECT displayQF FROM configSpec_Products WHERE configProduct="&ArrProduct(i) & " and specProduct=" & pcCartArray(f,0) 
												set rs=server.CreateObject("ADODB.RecordSet")
												set rs=conntemp.execute(query)
												if err.number<>0 then
													'//Logs error to the database
													call LogErrorToDatabase()
													'//clear any objects
													set rs=nothing
													'//close any connections
													call closedb()
													'//redirect to error page
													response.redirect "techErr.asp?err="&pcStrCustRefID
												end if
												
												btDisplayQF=rs("displayQF")
												set rs=nothing
												
												query="SELECT pcprod_minimumqty FROM Products WHERE idproduct=" & ArrProduct(i) & ";"
												set rsQ=connTemp.execute(query)
												tmpMinQty=1
												if not rsQ.eof then
													tmpMinQty=rsQ("pcprod_minimumqty")
													if IsNull(tmpMinQty) or tmpMinQty="" then
														tmpMinQty=1
													else
														if tmpMinQty="0" then
															tmpMinQty=1
														end if
													end if
												end if
												set rsQ=nothing
												tmpDefault=0
												query="SELECT cdefault FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & " AND configProduct=" & ArrProduct(i) & " AND cdefault<>0;"
												set rsQ=connTemp.execute(query)
												if not rsQ.eof then
													tmpDefault=rsQ("cdefault")
													if IsNull(tmpDefault) or tmpDefault="" then
														tmpDefault=0
													else
														if tmpDefault<>"0" then
														 	tmpDefault=1
														end if
													end if
												end if
												set rsQ=nothing %>
												<table width="100%" border="0" cellspacing="0" cellpadding="0">
													<tr valign="top"> 
														<td width="85%">
															<p>
																<%=strCategoryDesc%>:&nbsp;
																<%if btDisplayQF=True AND clng(ArrQuantity(i))>1 then%>(<%=ArrQuantity(i)%>)&nbsp;<%end if%>
																<%=strDescription%>
															</p>
														</td>
														<td width="15%" align="right" nowrap="nowrap">
															<p>
																<%if (ccur(ArrValue(i))<>0) or ((((ArrQuantity(i)-clng(tmpMinQty)<>0) AND (tmpDefault=1)) OR ((ArrQuantity(i)-1<>0) AND (tmpDefault=0))) and (ArrPrice(i)<>0)) then
																	if (ArrQuantity(i)-clng(tmpMinQty))>=0 then
																		if tmpDefault=1 then
																			UPrice=(ArrQuantity(i)-clng(tmpMinQty))*ArrPrice(i)
																		else
																			UPrice=(ArrQuantity(i)-1)*ArrPrice(i)
																		end if
																	else
																		UPrice=0
																	end if %>
																	<%=scCurSign & money(ccur((ArrValue(i)+UPrice)*pcCartArray(f,2)))%>
																<%else
																	if tmpDefault=1 then%>
																		<%=dictLanguage.Item(Session("language")&"_defaultnotice_1")%>
																	<%end if
																end if%>
															</p>
														</td>
													</tr>
												</table>
											<% next %>
										<% end if %>
									</div>
								</td>
								<td>&nbsp;</td>
								<td align="right">
									<%if pcv_FinalizedQuote=0 then%>
									<%call opendb()
									queryQ="SELECT TOP 1 configProduct FROM configSpec_products WHERE specProduct=" & pcCartArray(f,0) & ";"
									set rsQ=connTemp.execute(queryQ)
									if err.number<>0 then
										'//Logs error to the database
										call LogErrorToDatabase()
										'//clear any objects
										set rsQ=nothing
										'//close any connections
										call closedb()
										'//redirect to error page
										response.redirect "techErr.asp?err="&pcStrCustRefID
									end if
									if not rsQ.eof then%>
									<a href="Reconfigure.asp?pcCartIndex=<%=f%>"><img src="<%=rslayout("reconfigure")%>" alt="Reconfigure <%=pcCartArray(f,1)%>"></a>
									<%End if
									set rsQ=nothing%>
									<%end if%>
								</td>
							</tr>
							<% '// END 3nd Row - BTO Product Details %>
						<% End if 
						'BTO ADDON-E %>
						
						<% '// START 4th Row - Product Options %>
						<%
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' START: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						Dim pcv_strOptionsArray, pcv_intOptionLoopSize, pcv_intOptionLoopCounter, tempPrice, tAprice
						Dim pcArray_strOptionsPrice, pcArray_strOptions, pcArray_strSelectedOptions
						
						pcv_strOptionsArray = trim(pcCartArray(f,4))
						
						if len(pcv_strOptionsArray)>0 then %>
							<tr valign="top">
								<td>&nbsp;</td>
								<td colspan="3">
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<%
									'#####################
									' START LOOP
									'#####################	
									
									'// Generate Our Local Arrays from our Stored Arrays  
									
									' Column 11) pcv_strSelectedOptions '// Array of Individual Selected Options Id Numbers	
									pcArray_strSelectedOptions = ""					
									pcArray_strSelectedOptions = Split(trim(pcCartArray(f,11)),chr(124))
									
									' Column 25) pcv_strOptionsPriceArray '// Array of Individual Options Prices
									pcArray_strOptionsPrice = ""
									pcArray_strOptionsPrice = Split(trim(pcCartArray(f,25)),chr(124))
									
									' Column 4) pcv_strOptionsArray '// Array of Product "option groups: options"
									pcArray_strOptions = ""
									pcArray_strOptions = Split(trim(pcv_strOptionsArray),chr(124))
									
									' Get Our Loop Size
									pcv_intOptionLoopSize = 0
									pcv_intOptionLoopSize = Ubound(pcArray_strSelectedOptions)
									
									' Start in Position One
									pcv_intOptionLoopCounter = 0
									
									' Display Our Options
									For pcv_intOptionLoopCounter = 0 to pcv_intOptionLoopSize %>
										<tr>
											<td width="67%"><p><%=pcArray_strOptions(pcv_intOptionLoopCounter) %></p></td>
											<td align="right" width="33%">									
												<% tempPrice = pcArray_strOptionsPrice(pcv_intOptionLoopCounter)
												if tempPrice="" or tempPrice=0 then
													response.write "&nbsp;"
												else %>
													<table width="100%" cellpadding="0" cellspacing="0" border="0">
														<tr>
															<td align="left" width="50%">
																<%=scCurSign&money(tempPrice)%>
															</td>
															<td align="right" width="50%">
																<%									
																tAprice=(tempPrice*ccur(pcCartArray(f,2)))
																response.write scCurSign&money(tAprice) 
																%>
															</td>
														</tr>
													</table>
												<% end if %>			
											</td>
										</tr>
									<% Next
									'#####################
									' END LOOP
									'#####################	
						
						
									'// If there are product options AND NOt GGG, show link to edit them
									if trim(pcCartArray(f,16))="" AND Session("Cust_BuyGift")="" then %>						
										<tr valign="top">
											<td colspan="3" class="pcSmallText">
												<p>
													<a href="viewPrd.asp?idproduct=<%=pcCartArray(f,0)%>&index=<%=f%>&imode=updOrd">
														<%=dictLanguage.Item(Session("language")&"_showcart_21")%>
													</a>
												</p>
											</td>
										</tr>							
									<% end if %>
								</table>
								
							</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
						<% 
						End if
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' END: SHOW PRODUCT OPTIONS
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						
						pRowPrice=pRowPrice + ccur(pcCartArray(f,2) * pcCartArray(f,5))	%>
						<% '// END 4th Row - Product Options %>
							
						<% '// START 5th Row - Custom Input Fields %>
						<% 'if there are custom input fields, show them here
						if trim(pcCartArray(f,21))<>"" then %>
							<tr>
								<td>&nbsp;</td>
								<td colspan="3">
									<p><% response.write replace(pcCartArray(f,21),"''","'") %></p>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% End if %>

						<% 'if there are custom input fields and NO product options, and not GGG, and NOT BTO, show EDIT here
						if len(pcv_strOptionsArray)=0 AND trim(pcCartArray(f,16))="" AND Session("Cust_BuyGift")="" then
							query= "SELECT xfield1,xfield2,xfield3 FROM products WHERE idproduct="&pcCartArray(f,0)
							set rs=server.createobject("adodb.recordset")
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								set rs=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							if not rs.eof then
								if rs("xfield1") <> 0 or rs("xfield2") <> 0 OR rs("xfield3") <> 0 then %>
									<tr valign="top">
										<td>&nbsp;</td>
										<td colspan="3">
											<table width="100%" border="0" cellspacing="0" cellpadding="0">
												<tr valign="top">
													<td colspan="3" class="pcSmallText">
														<p>
															<a href="viewPrd.asp?idproduct=<%=pcCartArray(f,0)%>&index=<%=f%>&imode=updOrd">
																<%=dictLanguage.Item(Session("language")&"_showcart_27")%>
															</a>
														</p>
													</td>
												</tr>							
											</table>
										</td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									</tr>
								<% End if %>
							<% End if 
							set rs=nothing %>							
						<% End if %>						            
						<% '// END 5th Row - Custom Input Fields %>

						<% 'if items quantities discounts apply to this product, show the total applied amount here
						if pcCartArray(f,16)<>"" then
							itemsDiscounts=0
							for i=lbound(ArrProduct) to (UBound(ArrProduct)-1)
								query="select quantityFrom, quantityUntil, discountperUnit, percentage, discountperWUnit from discountsPerQuantity where IDProduct=" & ArrProduct(i)
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								
								if err.number<>0 then
									call LogErrorToDatabase()
									set rs=nothing
									call closedb()
									response.redirect "techErr.asp?err="&pcStrCustRefID
								end if
								
								TempDiscount=0
								do while not rs.eof
									QFrom=rs("quantityFrom")
									QTo=rs("quantityUntil")
									DUnit=rs("discountperUnit")
									QPercent=rs("percentage")
									DWUnit=rs("discountperWUnit")
									if (DWUnit=0) and (DUnit>0) then
										DWUnit=DUnit
									end if
									

									TempD1=0
									if (clng(ArrQuantity(i)*pcCartArray(f,2))>=clng(QFrom)) and (clng(ArrQuantity(i)*pcCartArray(f,2))<=clng(QTo)) then
										if QPercent="-1" then
											if session("customerType")=1 then
												TempD1=ArrQuantity(i)*pcCartArray(f,2)*ArrPrice(i)*0.01*DWUnit
											else
												TempD1=ArrQuantity(i)*pcCartArray(f,2)*ArrPrice(i)*0.01*DUnit
											end if
										else
											if session("customerType")=1 then
												TempD1=ArrQuantity(i)*pcCartArray(f,2)*DWUnit
											else
												TempD1=ArrQuantity(i)*pcCartArray(f,2)*DUnit
											end if
										end if
									end if
									TempDiscount=TempDiscount+TempD1
									rs.movenext
								loop
								set rs=nothing
								itemsDiscounts=ItemsDiscounts+TempDiscount
							next			

							if ItemsDiscounts>0 then
								ItemsDiscounts=round(ItemsDiscounts+0.001,2)
								pcCartArray(f,30)=ItemsDiscounts
								pRowPrice=pRowPrice-ItemsDiscounts %>
								
								<% '// START 6th Row - Discounts %>
								<tr> 							
									<td>&nbsp;</td>
									<td colspan="2" align="right">
										<p><%=dictLanguage.Item(Session("language")&"_showcart_23")%></p>
									</td>
									<td align="right" nowrap>
										- <% response.write scCurSign &  money(ItemsDiscounts) %>
									</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								</tr>
							<% else
								pcCartArray(f,30)=0
							end if
						End if%>
						<% '// END 6th Row - Discounts %>

						<% 'BTO Additional Charges
						if trim(pcCartArray(f,16))<>"" then 
							query="SELECT stringCProducts, stringCValues, stringCCategories FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(f,16))
							set rs=server.CreateObject("ADODB.RecordSet")	
							set rs=conntemp.execute(query)
							
							if err.number<>0 then
								call LogErrorToDatabase()
								set rs=nothing
								call closedb()
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if

							stringCProducts=rs("stringCProducts")
							stringCValues=rs("stringCValues")
							stringCCategories=rs("stringCCategories")
							ArrCProduct=Split(stringCProducts, ",")
							ArrCValue=Split(stringCValues, ",")
							ArrCCategory=Split(stringCCategories, ",")
							set rs=nothing %>
								
							<% if ArrCProduct(0)<>"na" then%>
								<% '// START 7th Row - BTO Additional Charges %>				
								<tr valign="top"> 
									<td>&nbsp;</td>
									<td colspan="3"> 
										<div class="pcShowBTOconfiguration">
											<table width="100%" border="0" cellspacing="0" cellpadding="0">
												<% pRowPrice=pRowPrice+ccur(pcCartArray(f,31)) %>
												<tr> 
													<td>
													<p><strong><%=bto_dictLanguage.Item(Session("language")&"_viewcart_3")%></strong></p>
													</td>
													<td>&nbsp;</td>
												</tr>
												<% for i=lbound(ArrCProduct) to (UBound(ArrCProduct)-1)
													if pcv_SpecialServer=1 then
														ArrCValue(i)=replace(ArrCValue(i),".",",")
													end if
													query="SELECT categories.categoryDesc, products.description FROM categories, products WHERE (((categories.idCategory)="&ArrCCategory(i)&") AND ((products.idProduct)="&ArrCProduct(i)&"))" 
													set rs=server.CreateObject("ADODB.RecordSet")
													set rs=conntemp.execute(query)
													
													if err.number<>0 then
														call LogErrorToDatabase()
														set rs=nothing
														call closedb()
														response.redirect "techErr.asp?err="&pcStrCustRefID
													end if
													
													strCategoryDesc=rs("categoryDesc")
													strDescription=rs("description")
													set rs=nothing %>
	
													<tr valign="top"> 
														<td width="85%">
															<p><%=strCategoryDesc%>: <%=strDescription%></p>
														</td>
														<td width="15%" align="right" nowrap="nowrap">
														<p>
														<%if (ccur(ArrCValue(i))>0) then%>
															<%=scCurSign & money(ArrCValue(i))%>
														<%end if%>
														</p>
														</td>
													</tr>
												<% next %>
											</table>
										</div>
									</td>
									<td>&nbsp;</td>
									<td align="right"><%if pcv_FinalizedQuote=0 then%><a href="RePrdAddCharges.asp?pcCartIndex=<%=f%>"><img src="<%=rslayout("reconfigure")%>" alt="Reconfigure <%=pcCartArray(f,1)%>"></a><%end if%></td>
								</tr>
							<% end if %>							
						<% End if 
						'BTO Additional Charges
						'// END 7th Row - BTO Additional Charges
							
						'// START 8th Row - Quantity Discounts	
						'if quantity discounts apply to this product, show the total applied amount here
						if trim(pcCartArray(f,15))<>"" AND trim(pcCartArray(f,15))>0 then
							pRowPrice=pRowPrice-ccur(pcCartArray(f,15))
							%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
									<p><a href="javascript:win('priceBreaks.asp?idproduct=<%=pcCartArray(f,0)%>')"><%response.write dictLanguage.Item(Session("language")&"_showcart_20")%></a><%response.write dictLanguage.Item(Session("language")&"_showcart_20b")%></p>
								</td>
								<td align="right" nowrap>- 
								<% response.write scCurSign &  money(pcCartArray(f,15)) %>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% End if %>
						<% '// END 8th Row - Quantity Discounts
						 
						'// START 9th Row - Product Subtotal %>	
						<% if pExtRowPrice<>pRowPrice then %>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%response.write dictLanguage.Item(Session("language")&"_showcart_22")%>
								</td>
								<td align="right">
								<% response.write scCurSign &  money(pRowPrice) %>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% end if %>
						<% '// END 9th Row - Product Subtotal %>

						<% 
						'SB S
						if (pcCartArray(f,38))>0 then %>
							<tr>				
								<td></td>
								<td colspan="3">
									<!--#include file="inc_sb_widget.asp"-->
                          		</td>
                          		<td></td>
                          		<td></td>
                      		</tr>							
					 		<%
					   		'// If there's a trial set the line total to the trial price
							if pcv_intIsTrial = "1" Then
						  		pRowPrice = pcv_curTrialAmount
							end if 
							
						End if 
						'SB E  
						%>
						<% '// START 10th Row - Cross Sell Bundle Discount %>	
						<% if (pcCartArray(f,27)>"0") AND (pcCartArray(f,28)>"0") then 	%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%response.write dictLanguage.Item(Session("language")&"_showcart_26")%>
								</td>
								<td align="right">
								<% response.write scCurSign &  money( ((ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28))) * -1) * pcCartArray(f,2)) %>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% end if %>
						<%'// END 10th Row - Cross Sell Bundle Discount
						
						'// START 11th Row - Cross Sell Bundle Subtotal %>	
						<% if (pcCartArray(f,27)>"0") AND (pcCartArray(f,28)>"0") then 
						    pRowPrice = ( ccur(pRowPrice) + ccur(ProList(cint(pcCartArray(f,27)),2)) ) - ( ( ccur(pcCartArray(f,28)) + ccur(pcCartArray(cint(pcCartArray(f,27)),28) ) ) * pcCartArray(f,2) )%>
							<tr> 							
								<td>&nbsp;</td>
								<td colspan="2" align="right">
								<%response.write dictLanguage.Item(Session("language")&"_showcart_22")%>
								</td>
								<td align="right">
								<% response.write scCurSign &  money(pRowPrice) %>
								</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<% end if %>
						<% '// END 11th Row - Cross Sell Bundle Subtotal %>	

						<tr>
							<td colspan="6"><hr></td>
						</tr>

						<%
						ProList(f,2)=pRowPrice
						
						'// Don't Add to total if parent of a Bundle Cross Sell Product
						pcv_HaveBundles=0
						if pcCartArray(f,27)=-1 then
							for mc=1 to pcCartIndex
								if (pcCartArray(mc,27)<>"") AND (pcCartArray(mc,12)<>"") then
									if cint(pcCartArray(mc,27))=f AND cint(pcCartArray(mc,12))="0" then
									 	pcv_HaveBundles=1
										exit for
									end if
								end if
							next
						end if
						if (pcCartArray(f,27)>-1) OR (pcv_HaveBundles=0) then
						    total=total + pRowPrice
						end if

						if Cint(pcCartArray(f,9))>totalDeliveringTime then
							totalDeliveringTime=Cint(pcCartArray(f,9))
						end if
					end if ' item deleted						
				next
				
				session("CartProList")=ProList
				
			' ------------------------------------------------------
			' START - Calculate category-based quantity discounts
			' ------------------------------------------------------
				CatDiscTotal=0
	
				query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					'//Logs error to the database
					call LogErrorToDatabase()
					'//clear any objects
					set rs=nothing
					'//close any connections
					call closedb()
					'//redirect to error page
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				Do While not rs.eof
					CatSubQty=0
					CatSubTotal=0
					CatSubDiscount=0
					CanNotRun=0
					IDCat=rs("IDCat")
					query="SELECT categories_products.idcategory FROM categories_products INNER JOIN pcPrdPromotions ON categories_products.idproduct=pcPrdPromotions.idproduct WHERE categories_products.idcategory=" & IDCat & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						CanNotRun=1
					end if
					set rsQ=nothing
					
					IF CanNotRun=0 THEN
				
					For f=1 to pcCartIndex
						if (ProList(f,1)=0) and (ProList(f,4)=0) then 
							query="select idproduct from categories_products where idcategory=" & IDCat & " and idproduct=" & ProList(f,0)
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=connTemp.execute(query)
							if err.number<>0 then
								'//Logs error to the database
								call LogErrorToDatabase()
								'//clear any objects
								set rstemp=nothing
								'//close any connections
								call closedb()
								'//redirect to error page
								response.redirect "techErr.asp?err="&pcStrCustRefID
							end if
							
							if not rstemp.eof then
								CatSubQty=CatSubQty+ProList(f,3)
								CatSubTotal=CatSubTotal+ProList(f,2)
								ProList(f,4)=1
							end if
							set rstemp=nothing
						end if
					Next
			
					if CatSubQty>0 then
						query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & IDCat & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
						set rstemp=server.CreateObject("ADODB.RecordSet")
						set rstemp=conntemp.execute(query)
						if err.number<>0 then
							'//Logs error to the database
							call LogErrorToDatabase()
							'//clear any objects
							set rstemp=nothing
							'//close any connections
							call closedb()
							'//redirect to error page
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
	
						if not rstemp.eof then
							' there are quantity discounts defined for that quantity 
							pDiscountPerUnit=rstemp("pcCD_discountPerUnit")
							pDiscountPerWUnit=rstemp("pcCD_discountPerWUnit")
							pPercentage=rstemp("pcCD_percentage")
							pbaseproductonly=rstemp("pcCD_baseproductonly")
							set rstemp=nothing
							
							if session("customerType")<>1 then  'customer is a normal user
								if pPercentage="0" then 
									CatSubDiscount=pDiscountPerUnit*CatSubQty
								else
									CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
								end if
							else  'customer is a wholesale customer
								if pPercentage="0" then 
									CatSubDiscount=pDiscountPerWUnit*CatSubQty
								else
									CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
								end if
							end if
						end if
					end if
	
					CatDiscTotal=CatDiscTotal+CatSubDiscount
					
					END IF 'CanNotRun
					rs.MoveNext
				loop
				set rs=nothing				
				'// Round the Category Discount to two decimals
				if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
					CatDiscTotal = Round(CatDiscTotal,2)
				end if
			' ------------------------------------------------------
			' END - Calculate category-based quantity discounts
			' ------------------------------------------------------
			
			'Display Applied Product Promotions (if any)
			TotalPromotions=0
			if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
				PromoArr1=Session("pcPromoSession")
				PromoIndex=Session("pcPromoIndex")
				For m=1 to PromoIndex
					' Show message and add to total if promotion discount is > 0
					if PromoArr1(m,2)>0 then
					%>
					<tr>
						<td colspan="5" align="right">
						<%=PromoArr1(m,1)%>
						</td>
						<td align="right">
							-<%=scCurSign  & money(PromoArr1(m,2))%>
							<%TotalPromotions=TotalPromotions+cdbl(PromoArr1(m,2))%>
						</td>
					</tr>
                    <%
					end if
				Next
			end if
			

			' Calculate & display order total
				total=total-CatDiscTotal-TotalPromotions %>
				<tr>
					<td colspan="2">
						<div><input type="image" id="submit" name="Submit" src="<%=RSlayout("recalculate")%>" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);} else {return(true);}"></div>
						<%
						'// GGG Add-on start 
						if pcv_strShowCheckoutBtn=1 then %>
								<% if HaveGcsTest=1 then %>
									<div style="padding-top: 3px;"><input type="image" id="submit" name="image1a" value="Gift Certificates" src="<%=RSlayout("checkout")%>" border="0" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);} else {document.recalculate.actGCs.value='Hello'; return(true);}"></div>
								<% else %>
									<div style="padding-top: 3px;"><a href="<%=strRedirectSSL%>" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);} return(checkQtyChange());"><img src="<%=RSlayout("checkout")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_12")%>"></a></div>
								<% end if %>
						<%
						'// GGG Add-on end
						end if
						%>
						<!--#include file="pcPay_PayPal.asp"-->
						<!--#include file="pcPay_GoogleCheckout.asp"-->
						<%'GGG Add-on start
						gHaveGR=0
						query="select pcEv_IDEvent from pcEvents where pcEv_IDCustomer=" & Session("idCustomer") & " and pcEv_Active=1"
						set rs1=conntemp.execute(query)
						if not rs1.eof then
							gHaveGR=1
						end if
						if (not (Session("idCustomer")=0)) and (gHaveGR=1) and (session("Cust_buyGift")="") then%>
							<div style="padding-top: 3px;"><a href="ggg_addtoGR.asp" onclick="javascript: if ((RemainIssue!='') || (RemainIssue1!='')) {alert('<% Response.write(dictLanguage.Item(Session("language")&"_alert_8b"))%>'); return(false);}"><img src="<%=RSlayout("AddToRegistry")%>" border="0" alt="<%=dictLanguage.Item(Session("language")&"_altTag_14")%>"></a></div>
						<%end if
						'GGG Add-on end%>
					</td>
					<td colspan="4" align="right">
					<div>
					<strong><%response.write dictLanguage.Item(Session("language")&"_showcart_12")%> <% response.write scCurSign & money(total) %></strong>
					</div>
					
					<%' Display category-based quantity discounts
					if CatDiscTotal>0 then%>
					<div style="padding-top: 3px;">
						<%response.write dictLanguage.Item(Session("language")&"_catdisc_1")%> <% response.write scCurSign & money(CatDiscTotal) %>
					</div>
					<% end if %>
					
					</td>
				</tr>
			</table>
		</td>
	</tr>
	
	<%
	' ------------------------------------------------------
	' Show horizontal line if there is more content
	' ------------------------------------------------------
	if	scShowCartWeight="-1" or (iShipService=0 and scShowEstimateLink="-1") then
	%>
	<tr>
		<td><hr></td>
	</tr>
	<%
	end if
	
	' ------------------------------------------------------
	' START - Show shopping cart content weight
	' ------------------------------------------------------
	if	scShowCartWeight="-1" then
	if cdbl(totalRowWeight)>0 AND cdbl(totalRowWeight)<1 then
		totalRowWeight=1
	end if
	totalRowWeight=round(totalRowWeight,0)
	%>
		<tr>
			<td>
			<% if scShipFromWeightUnit="KGS" then
				pKilos=Int(totalRowWeight/1000)
				pWeight_g=totalRowWeight-(pKilos*1000)
				%>
				<b><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_a")%></b> <%=pKilos&" kg "%>
				<% if pWeight_g>0 then 
					response.write pWeight_g&" g"
				end if %>
			<% else
				pPounds=Int(totalRowWeight/16)
				pWeight_oz=totalRowWeight-(pPounds*16)
				%>
				<b><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_a")%></b> <%=pPounds&" lbs. "%>
				<% if pWeight_oz>0 then 
					response.write pWeight_oz&" oz."
				end if %>
			<% end if %></td>
		</tr>
	<% end if
	' ------------------------------------------------------
	' END - Show shopping cart content weight
	' ------------------------------------------------------
	
	' ------------------------------------------------------
	' START - Show estimated shipping charges link
	' ------------------------------------------------------
		if iShipService=0 then
			if scShowEstimateLink="-1" or request("show")="1" then %>
			<tr>
				<td>
					<a href="javascript:winShipPreview('estimateShipCost.asp')"><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_b")%></a>
				</td>
			</tr>
		<%
			end if
		end if
	' ------------------------------------------------------
	' END - Show estimated shipping charges link
	' ------------------------------------------------------
	
	' ------------------------------------------------------
	' START - Promotions
	' ------------------------------------------------------
	if PromoMsgStr<>"" then
	%>
		<tr>
			<td align="center">
            <div class="pcPromoMessage">
                <span class="pcLargerText"><%=dictLanguage.Item(Session("language")&"_showcart_29")%></span>
                <ul><%=PromoMsgStr%></ul>
            </div>
            </td>
		</tr>
	<%
	end if
	' ------------------------------------------------------
	' END - Promotions
	' ------------------------------------------------------
	
	' ------------------------------------------------------
	' START - Show Gift Wrapping Overview
	' ------------------------------------------------------
	
	if (session("Cust_GW")="1") and (pcv_GWDetails<>"") and (session("Cust_GWText")="1") then%>
	<tr>
		<td>
			<%=pcv_GWDetails%>
		</td>
	</tr>
	<% end if
	
	' ------------------------------------------------------
	' END - Show Gift Wrapping Overview
	' ------------------------------------------------------
	
	' ------------------------------------------------------
	' START - Cross selling
	' ------------------------------------------------------
		
		dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText, pcv_strCSQuery, pcv_strHaveResults, pcv_intProductCount, pcArray_CSRelations
		
		'// Get Cross Sell Settings - Sitewide 
		query= "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,crossSellText,cs_CartViewCnt FROM crossSelldata WHERE id=1;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if		
		If NOT rs.EOF Then
			scCS=rs("cs_status")
			cs_showprod=rs("cs_showprod")
			cs_showcart=rs("cs_showcart")
			cs_showimage=rs("cs_showimage")
			crossSellText=rs("crossSellText")
			cs_ViewCnt=rs("cs_CartViewCnt")
		End If
		set rs=nothing
		
		'// Do Not Display if CS is turned "Off"
		If scCS=-1 AND cs_showcart="-1" Then
		
			'// Check if there are items for cross sell in the database
			pcv_strCSQuery = ""
			cs_Source=1
			pcv_cs_headerflag=0
			tmp_PList=""
			pcv_strCSQuery = pcv_strCSQuery & "cs_relationships.idproduct=0"  
			tCnt=0
			for f=pcCartIndex to 1 Step -1
			
				pidproduct=pcCartArray(f,0)
					
				IF inStr(","& tmp_PList &",",","& pidproduct &",")=0 THEN  
				
					if tCnt=0 then
						tmp_PList=tmp_PList & pidproduct 				
					else
						tmp_PList=tmp_PList & "," & pidproduct  
					end if				

					'// cs_relationships.discount)>0
										
					'// Build Cross Sell Relationship List
					pcv_strCSQuery = pcv_strCSQuery & " OR cs_relationships.idproduct="& pidproduct
					
				END IF '// Check existing IDProduct
				tCnt=tCnt+1
			next '// Move to next product in the cart
			tCnt=0
			
			If len(tmp_PList)>0 Then
				pcv_strCSunavailable = "(cs_relationships.idrelation NOT IN ("& tmp_PList &")) AND " 
			Else
				pcv_strCSunavailable = ""
			End If

			query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.discount, cs_relationships.ispercent,cs_relationships.isRequired, products.servicespec, products.price, products.description FROM cs_relationships INNER JOIN products ON cs_relationships.idrelation=products.idProduct WHERE ("& pcv_strCSunavailable &"("& pcv_strCSQuery &") AND ((products.active)=-1) AND ((products.removed)=0)) ORDER BY cs_relationships.num,cs_relationships.idrelation;"		
			set rs=server.createobject("adodb.recordset")
			set rs=conntemp.execute(query)	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			pcv_strHaveResults=0
			if NOT rs.eof then
				pcArray_CSRelations = rs.getRows()
				pcv_intProductCount = UBound(pcArray_CSRelations,2)+1
				pcv_strHaveResults=1
			end if
			set rs=nothing		
			
			tCnt=Cint(0)	
			pcsFilterOverRide="1"

			if pcv_strHaveResults=1 then	
			
				'// Start: viewPrd
				cs_pCnt=Cint(0)
				cs_pOptCnt=Cint(0)
				cs_pAddtoCart=Cint(0)
				pcv_intCategoryActive=2	'// set bundle group to inactive
				pcv_intAccessoryActive=2 '// set accessories group to inactive
				cs_count=Cint(0)
				session("listcross")=""
				
				do while ( (tCnt < pcv_intProductCount) AND (tCnt < cs_ViewCnt))				
					
					pidrelation=pcArray_CSRelations(1,tCnt) '// rs("idrelation")
					pcsType=pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
					pDiscount=pcArray_CSRelations(3,tCnt) '// rs("discount")
					cs_pserviceSpec=pcArray_CSRelations(6,tCnt)				
					pcArray_CSRelations(8,tCnt) = 1
					
					If (pcsType="Accessory") OR ((pcsType="Bundle") AND (pDiscount>0)) Then
						
						'// CHECK IF BUNDLES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY		
						'// CHECK IF ACCESSORIES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY  						
						If Session("customerType")=1 Then
							pcv_strCSTemp=""
						else
							pcv_strCSTemp=" AND pccats_RetailHide<>1 "
						end if									
						query="SELECT categories_products.idProduct "
						query=query+"FROM categories_products " 
						query=query+"INNER JOIN categories "
						query=query+"ON categories_products.idCategory = categories.idCategory "
						query=query+"WHERE categories_products.idProduct="& pidrelation &" AND iBTOhide=0 " & pcv_strCSTemp & " "
						query=query+"ORDER BY priority, categoryDesc ASC;"	
						set rsCheckCategory=server.CreateObject("ADODB.RecordSet")
						set rsCheckCategory=conntemp.execute(query)									
						If NOT rsCheckCategory.eof Then
							If pcsType="Accessory" Then
								pcv_intAccessoryActive=1
							End If
							If pcsType="Bundle" Then							
								pcv_intCategoryActive=1
							End If	
						Else
							session("listcross")=session("listcross") & "," & pidrelation					
						End If	
						set rsCheckCategory=nothing
						
					End If '// If (pcsType="Bundle") AND (pDiscount>0) Then	
	
					pcv_intOptionsExist=0
					
					'// CHECK FOR REQUIRED OPTIONS							
					pcv_intOptionsExist=pcf_CheckForReqOptions(pidrelation) '// check options function (1=YES, 2=NO)			
	
	
					'// CHECK FOR REQUIRED INPUT FIELDS
					if pcv_intOptionsExist=2 then
						pcv_intOptionsExist=pcf_CheckForReqInputFields(pidrelation)
					end if				
	
	
					'// VALIDATE
					if (cs_pserviceSpec=true) OR (pcv_intOptionsExist = 1) then
						If pcsType<>"Accessory" Then
							cs_pOptCnt=cs_pOptCnt+1
						End If
						pcArray_CSRelations(8,tCnt) = 0					
					End If	
					If pcsType<>"Accessory" Then
						cs_pCnt=cs_pCnt+1 
					End If
					tCnt=tCnt+1				
				loop
				'// End: viewPrd
			
				if pcv_intAccessoryActive=1 then
					
					cs_DisplayCheckBox=0
					cs_Bundle=0

					if pcv_cs_headerflag=0 then
						'// Only display header once
						pcv_cs_headerflag=1 %>
						<tr> 
							<td>
								<hr>
							</td>
						</tr>
						<tr>
							<td class="pcSectionTitle"><%=crossSellText%></td>
						</tr>
					<% end if %>
				
					<tr>
						<td>
							<% if cs_showImage="-1" then %>
								<!--#include file="cs_img.asp"-->
							<% else %>
								<!--#include file="cs.asp"-->
							<% end if %>
						</td>
					</tr>
			
				<%		
				end if '// if (cint(cs_pOptCnt) <> cint(cs_pCnt)) AND (pcv_intCategoryActive=1) then
				
							
			end if '// if pcv_strHaveResults=1 then			
		End If '// if scCS=-1 AND cs_showcart="-1" then
		session("listcross")=""
		%>
	
	</table>
</form>
</div>
<% call closedb() %>
<%Session("pcCartSession")=pcCartArray%>
<!--#include file="footer.asp"-->