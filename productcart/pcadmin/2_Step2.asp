<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping Configuration: Select Shipping Services" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"--> 
<% if request.form("submit")<>"" then
			
	USPS_Service=request.form("USPS_Service")
	Session("ship_USPS_Service")=USPS_Service
			
	if USPS_Service="" then
		response.redirect "2_Step2.asp?msg="&Server.URLEncode("Select at least one service.")
		response.end
	end if
	freeshipStr=""
	handlingStr=""
	
	If request.form("free9901")="YES" then
		freeamt=request.form("amt9901")
		freeshipStr=freeshipStr&"9901|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9901")<>"0" AND request.form("handling9901")<>"" then
		If isNumeric(request.form("handling9901"))=true then
			handlingStr=handlingStr&"9901|"&replacecomma(request.form("handling9901"))&"|"&request.form("shfee9901")&","
		End If
	End if
	
	If request.form("free9902")="YES" then
		freeamt=request.form("amt9902")
		freeshipStr=freeshipStr&"9902|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9902")<>"0" AND request.form("handling9902")<>"" then
		If isNumeric(request.form("handling9902"))=true then
			handlingStr=handlingStr&"9902|"&replacecomma(request.form("handling9902"))&"|"&request.form("shfee9902")&","
		End If
	End if
	
	If request.form("free9903")="YES" then
		freeamt=request.form("amt9903")
		freeshipStr=freeshipStr&"9903|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9903")<>"0" AND request.form("handling9903")<>"" then
		If isNumeric(request.form("handling9903"))=true then
			handlingStr=handlingStr&"9903|"&replacecomma(request.form("handling9903"))&"|"&request.form("shfee9903")&","
		End If
	End if
	
	If request.form("free9904")="YES" then
		freeamt=request.form("amt9904")
		freeshipStr=freeshipStr&"9904|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9904")<>"0" AND request.form("handling9904")<>"" then
		If isNumeric(request.form("handling9904"))=true then
			handlingStr=handlingStr&"9904|"&replacecomma(request.form("handling9904"))&"|"&request.form("shfee9904")&","
		End If
	End if
	
	If request.form("free9905")="YES" then
		freeamt=request.form("amt9905")
		freeshipStr=freeshipStr&"9905|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9905")<>"0" AND request.form("handling9905")<>"" then
		If isNumeric(request.form("handling9905"))=true then
			handlingStr=handlingStr&"9905|"&replacecomma(request.form("handling9905"))&"|"&request.form("shfee9905")&","
		End If
	End if
	
	If request.form("free9906")="YES" then
		freeamt=request.form("amt9906")
		freeshipStr=freeshipStr&"9906|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9906")<>"0" AND request.form("handling9906")<>"" then
		If isNumeric(request.form("handling9906"))=true then
			handlingStr=handlingStr&"9906|"&replacecomma(request.form("handling9906"))&"|"&request.form("shfee9906")&","
		End If
	End if
	
	If request.form("free9907")="YES" then
		freeamt=request.form("amt9907")
		freeshipStr=freeshipStr&"9907|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9907")<>"0" AND request.form("handling9907")<>"" then
		If isNumeric(request.form("handling9907"))=true then
			handlingStr=handlingStr&"9907|"&replacecomma(request.form("handling9907"))&"|"&request.form("shfee9907")&","
		End If
	End if
	
	If request.form("free9908")="YES" then
		freeamt=request.form("amt9908")
		freeshipStr=freeshipStr&"9908|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9908")<>"0" AND request.form("handling9908")<>"" then
		If isNumeric(request.form("handling9908"))=true then
			handlingStr=handlingStr&"9908|"&replacecomma(request.form("handling9908"))&"|"&request.form("shfee9908")&","
		End If
	End if
	
	If request.form("free9909")="YES" then
		freeamt=request.form("amt9909")
		freeshipStr=freeshipStr&"9909|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9909")<>"0" AND request.form("handling9909")<>"" then
		If isNumeric(request.form("handling9909"))=true then
			handlingStr=handlingStr&"9909|"&replacecomma(request.form("handling9909"))&"|"&request.form("shfee9909")&","
		End If
	End if
	
	If request.form("free9910")="YES" then
		freeamt=request.form("amt9910")
		freeshipStr=freeshipStr&"9910|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9910")<>"0" AND request.form("handling9910")<>"" then
		If isNumeric(request.form("handling9910"))=true then
			handlingStr=handlingStr&"9910|"&replacecomma(request.form("handling9910"))&"|"&request.form("shfee9910")&","
		End If
	End if

	If request.form("free9911")="YES" then
		freeamt=request.form("amt9911")
		freeshipStr=freeshipStr&"9911|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9911")<>"0" AND request.form("handling9911")<>"" then
		If isNumeric(request.form("handling9911"))=true then
			handlingStr=handlingStr&"9911|"&replacecomma(request.form("handling9911"))&"|"&request.form("shfee9911")&","
		End If
	End if
	
	If request.form("free9912")="YES" then
		freeamt=request.form("amt9912")
		freeshipStr=freeshipStr&"9912|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9912")<>"0" AND request.form("handling9912")<>"" then
		If isNumeric(request.form("handling9912"))=true then
			handlingStr=handlingStr&"9912|"&replacecomma(request.form("handling9912"))&"|"&request.form("shfee9912")&","
		End If
	End if
	
	If request.form("free9913")="YES" then
		freeamt=request.form("amt9913")
		freeshipStr=freeshipStr&"9913|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9913")<>"0" AND request.form("handling9913")<>"" then
		If isNumeric(request.form("handling9913"))=true then
			handlingStr=handlingStr&"9913|"&replacecomma(request.form("handling9913"))&"|"&request.form("shfee9913")&","
		End If
	End if
	
	If request.form("free9914")="YES" then
		freeamt=request.form("amt9914")
		freeshipStr=freeshipStr&"9914|"&replacecomma(freeamt)&","
	End if
	If request.form("handling9914")<>"0" AND request.form("handling9914")<>"" then
		If isNumeric(request.form("handling9914"))=true then
			handlingStr=handlingStr&"9914|"&replacecomma(request.form("handling9914"))&"|"&request.form("shfee9914")&","
		End If
	End if
	
	Session("ship_USPS_freeshipStr")=freeshipStr
	Session("ship_USPS_handlingStr")=handlingStr
	response.redirect "2_Step3.asp"
	response.end
else %>

	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
    <form name="form1" method="post" action="2_Step2.asp" class="pcForms">
        <table class="pcCPcontent">
            <tr bgcolor='#DDEEFF' class='normal'>
                <td width='4%'><input type='checkbox' name='USPS_Service' value='9901' checked></td>
                <td colspan="2"><font color='#000000'><b>USPS Priority</b></font></td>
                </tr>
            <tr class='normal'>
                <td bgcolor='F1F1F1'>&nbsp;</td>
                <td colspan='2' bgcolor='F1F1F1'>
                <input name='free9901' type='checkbox' id='free9901' value='YES' >Offer free shipping for orders over $ <input name='amt9901' type='text' id='amt9901' size='6' maxlength='10' value='0.00'></td>
            </tr>
            <tr class='normal'>
                <td bgcolor='F1F1F1'>&nbsp;</td>
                <td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td>
            </tr>
            <tr class='normal'>
                <td bgcolor='F1F1F1'>&nbsp;</td>
                <td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9901' type='text' id='handling9901' size='6' maxlength='10' value='0.00'></td>
            </tr>
            <tr class='normal'>
                <td bgcolor='F1F1F1'>&nbsp;</td>
                <td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9901' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9901' value='0' checked>Integrate into shipping rate.</td>
            </tr>
    
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9902' checked></td><td colspan="2"><font color='#000000'><b>USPS Express</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9902' type='checkbox' id='free9902' value='YES' >Offer free shipping for orders over $ <input name='amt9902' type='text' id='amt9902' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9902' type='text' id='handling9902' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9902' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9902' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9903' checked></td><td colspan="2"><font color='#000000'><b>USPS Parcel</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9903' type='checkbox' id='free9903' value='YES' >Offer free shipping for orders over $ <input name='amt9903' type='text' id='amt9903' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9903' type='text' id='handling9903' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9903' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9903' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9904' checked></td><td colspan="2"><font color='#000000'><b>USPS First Class</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9904' type='checkbox' id='free9904' value='YES' >Offer free shipping for orders over $ <input name='amt9904' type='text' id='amt9904' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9904' type='text' id='handling9904' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9904' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9904' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9915' checked></td><td colspan="2"><font color='#000000'><b>USPS Bound Printed Matter</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9915' type='checkbox' id='free9915' value='YES' >Offer free shipping for orders over $ <input name='amt9915' type='text' id='amt9915' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9915' type='text' id='handling9915' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9915' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9915' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9916' checked></td><td colspan="2"><font color='#000000'><b>USPS Media Mail</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9916' type='checkbox' id='free9916' value='YES' >Offer free shipping for orders over $ <input name='amt9916' type='text' id='amt9916' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9916' type='text' id='handling9916' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9916' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9916' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9917' checked></td><td colspan="2"><font color='#000000'><b>USPS Library Mail</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9917' type='checkbox' id='free9917' value='YES' >Offer free shipping for orders over $ <input name='amt9917' type='text' id='amt9917' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9917' type='text' id='handling9917' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9917' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9917' value='0' checked>Integrate into shipping rate.</td></tr>
            <tr> 
                <td colspan="2">&nbsp;</td>
            </tr>
            <tr> 
                <td colspan="2"><h2>USPS - International shipping services</h2></td>
            </tr>
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9914' checked></td><td colspan="2"><font color='#000000'><b>Global Express Guaranteed<sup>&reg;</sup></b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9914' type='checkbox' id='free9914' value='YES' >Offer free shipping for orders over $ <input name='amt9914' type='text' id='amt9914' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9914' type='text' id='handling9914' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9914' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9914' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9905' checked></td><td colspan="2"><font color='#000000'><b>Global Express Guaranteed<sup>&reg;</sup> Non-Document Rectangular</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9905' type='checkbox' id='free9905' value='YES' >Offer free shipping for orders over $ <input name='amt9905' type='text' id='amt9905' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9905' type='text' id='handling9905' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9905' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9905' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9906' checked></td><td colspan="2"><font color='#000000'><b>Express Mail<sup>&reg;</sup> International (EMS)</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9906' type='checkbox' id='free9906' value='YES' >Offer free shipping for orders over $ <input name='amt9906' type='text' id='amt9906' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9906' type='text' id='handling9906' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9906' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9906' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9907' checked></td><td colspan="2"><font color='#000000'><b>Priority Mail<sup>&reg;</sup> International</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9907' type='checkbox' id='free9907' value='YES' >Offer free shipping for orders over $ <input name='amt9907' type='text' id='amt9907' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9907' type='text' id='handling9907' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9907' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9907' value='0' checked>Integrate into shipping rate.</td></tr>
            
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9908' checked></td><td colspan="2"><font color='#000000'><b>Priority Mail<sup>&reg;</sup> International Flat Rate Envelope</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9908' type='checkbox' id='free9908' value='YES' >Offer free shipping for orders over $ <input name='amt9908' type='text' id='amt9908' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9908' type='text' id='handling9908' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9908' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9908' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9909' checked></td><td colspan="2"><font color='#000000'><b>Priority Mail<sup>&reg;</sup> International Flat Rate Box</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9909' type='checkbox' id='free9909' value='YES' >Offer free shipping for orders over $ <input name='amt9909' type='text' id='amt9909' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9909' type='text' id='handling9909' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9909' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9909' value='0' checked>Integrate into shipping rate.</td></tr>
            
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9910' checked></td><td colspan="2"><font color='#000000'><b>Global Express Guaranteed<sup>&reg;</sup> Non-Document Non-Rectangular</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9910' type='checkbox' id='free9910' value='YES' >Offer free shipping for orders over $ <input name='amt9910' type='text' id='amt9910' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9910' type='text' id='handling9910' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9910' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9910' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9911' checked></td><td colspan="2"><font color='#000000'><b>Express Mail<sup>&reg;</sup> International (EMS) Flat Rate Envelope</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9911' type='checkbox' id='free9911' value='YES' >Offer free shipping for orders over $ <input name='amt9911' type='text' id='amt9911' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9911' type='text' id='handling9911' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9911' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9911' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9912' checked></td><td colspan="2"><font color='#000000'><b>First-Class Mail<sup>&reg;</sup> International</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9912' type='checkbox' id='free9912' value='YES' >Offer free shipping for orders over $ <input name='amt9912' type='text' id='amt9912' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9912' type='text' id='handling9912' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9912' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9912' value='0' checked>Integrate into shipping rate.</td></tr>
            
            <tr bgcolor='#DDEEFF' class='normal'><td width='4%'><input type='checkbox' name='USPS_Service' value='9913' checked></td><td colspan="2"><font color='#000000'><b>USPS Economy (Surface) Standard Post</b></font></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input name='free9913' type='checkbox' id='free9913' value='YES' >Offer free shipping for orders over $ <input name='amt9913' type='text' id='amt9913' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><hr align='left' width='325' size='1' noshade></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'>Add Handling Fee $ <input name='handling9913' type='text' id='handling9913' size='6' maxlength='10' value='0.00'></td></tr><tr class='normal'><td bgcolor='F1F1F1'>&nbsp;</td><td colspan='2' bgcolor='F1F1F1'><input type='radio' name='shfee9913' value='-1' checked>Display as a &quot;Shipping &amp; Handling&quot; charge.<br><input type='radio' name='shfee9913' value='0' checked>Integrate into shipping rate.</td></tr>
    
            <tr> 
                <td colspan="2"><hr></td>
            </tr>
            <tr> 
                <td colspan="2"><input type="submit" name="Submit" value="Continue" class="submit2"></td>
            </tr>
        </table>
    </form>
<% end if %>
<!--#include file="AdminFooter.asp"-->