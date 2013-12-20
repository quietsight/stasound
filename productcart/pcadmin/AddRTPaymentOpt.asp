<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pageTitle="Activate a Payment Gateway"
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
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<% pageTitle="Add a Real-Time Payment Option" %>
<!--#include file="RTGatewayConstants.asp"-->
<% 

dim query, connTemp, rs

'Create Temp Array for gateways that are not active
strTempGWArray=pcGWArray

' get payment types
err.clear
err.number=0
call openDb()  

query="SELECT gwCode,paymentDesc,active FROM paytypes;"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=connTemp.execute(query)

if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if NOT rs.eof then 
	iCnt=1
	pcv_activeString=""
	
	do until rs.eof
		varTemp=rs("gwCode")
		if varTemp<>999999 then
			varPaymentDesc=rs("paymentDesc")
			varActive=rs("active")
			if varActive<>"0" then
				pcvActiveAry=split(pcGWArray,",")
				for t=lbound(pcvActiveAry) to ubound(pcvActiveAry)-1
					intGWCode=pcvActiveAry(t)
					if intGWCode="//" then
						intGWCode=0
					end if
					if cint(varTemp)=cint(intGWCode) then
						strTempGWArray=replace(strTempGWArray,","&intGWCode&",",",")
						pcv_activeString=pcv_activeString&"<tr><td width='6%' height='21'><b><img src='images/lighton.gif' width='18' height='22'></b></td><td height='21'>"&varPaymentDesc&" is enabled</td></tr>"
					end if
				Next
			end if
		end if
		rs.moveNext
	loop
	set rs=nothing
end if

call closedb()
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<% ' Highlight NetSource Commerce Payment Gateway %>
<div style="position: relative;">
    <div style="position: absolute; left: 550px; top: 50px;"><a href="AddModRTPayment.asp?gwchoice=67"><img src="images/ei_logo_gradient_payment_gateway_175.jpg" alt="NetSource Commerce Payment Gateway" border="0" /></a></div>
</div>


<table class="pcCPcontent">

    <% if pcv_activeString<>"" then %>
        <tr>
            <td><strong>The following real-time payment options have been configured:</strong></td>
        </tr>
    <% end if %>

    <tr> 
        <td valign="top"> 
            <table border="0" width="100%" cellpadding="2" cellspacing="0">
                <% =pcv_activeString %>
                <% if pcv_activeString<>"" then %>
                <tr> 
                    <td colspan="2">&nbsp;</td>
                </tr>
                <% end if %>
                <tr> 
                    <td colspan="2">Select &amp; configure the payment gateway that you would like to use:</td>
                </tr>
                <tr>
                    <td class="pcCPspacer"></td>
                </tr>
                <% 
				pcvNonActiveAry=split(strTempGWArray,",")
                for v=lbound(pcvNonActiveAry) to ubound(pcvNonActiveAry)-1
                    varNonActive=pcvNonActiveAry(v)
					if varNonActive="//" then
						varNonActive=0
					else
						varNonActive=Cint(varNonActive)
					end if
                    call gwShowDesc()
					if pcGWPaymentDesc="" AND  pcGWPaymentURL="" then
					else %>
                    <tr> 
                        <td width="6%" height="21"></td>
                        <td height="21"><a href="pcConfigurePayment.asp?gwchoice=<%=pcGWCode&cint(pcvNonActiveAry(v))%>">Enable <%=pcGWPaymentDesc%></a>&nbsp;&nbsp;-&nbsp;&nbsp;<a href="<%=pcGWPaymentURL%>" target="_blank">Visit web site</a></td>
                    </tr>
                    <% end if
                Next %>
                <% if not request("gwChoice")="" then 'If gwChoice is empty, hide the rest of the page %>
                <tr> 
                    <td colspan="2">&nbsp;</td>
                </tr>
                <% end if %>
            </table>
        </td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->