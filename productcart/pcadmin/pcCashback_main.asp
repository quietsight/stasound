<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export to Bing Cashback" %>
<% section="genRpts" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/CashbackConstants.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
If len(LSCB_KEY)>0 Then
	pcv_strCashbackActive = 1
End If
%>
<% 
'// Checks for cookie
Dim CookieVar, ShowAgreement
ShowAgreement=0
CookieVar=Request.Cookies("AgreeCBLicense")

if request("RedirectURL")<>"" then
	Session("RedirectURL")=getUserInput(request("RedirectURL"),0)
end if

If NOT CookieVar="Agreed" then
	ShowAgreement=1
End If
If request.form("Submit2")<>"" then
	AgreeVar=request.form("agree")
	If AgreeVar=1 then
		'// Set cookie
		Response.Cookies("AgreeCBLicense")="Agreed"
		Response.Cookies("AgreeCBLicense").Expires=Date() + 365
		MyCookiePath=Request.ServerVariables("PATH_INFO")
		do while not (right(MyCookiePath,1)="/")
		MyCookiePath=mid(MyCookiePath,1,len(MyCookiePath)-1)
		loop
		Response.Cookies("AgreeCBLicense").Path=MyCookiePath
		response.redirect "pcCashback_main.asp"
	else
		'// Send message to agree
		AgreeMsg="Agree to the terms and conditions of the Microsoft Live Cashback End User License Agreement to continue."
		response.redirect "pcCashback_main.asp?AM="&server.URLEncode(AgreeMsg)
	end if
End If
%>
<% if ShowAgreement=0 then %>
    <table class="pcCPcontent">
        <tr>
        <td class="pcCPspacer"></td>
      </tr>
      <tr>
        <th>About the integration with Bing Cashback&reg;</th>
      </tr>
        <tr>
        <td class="pcCPspacer"></td>
      </tr>
      <tr>
            <tr>
                <td>
          		Bing cashback is a new offering from Microsoft that combines the search power of Bing with a comparison-shopping engine to bring consumers some of the best deals on the Web. <a href="http://wiki.earlyimpact.com/productcart/marketing-generate_cashback_file" target="_blank">See the documentation</a> for more information.
				<%if pcv_strCashbackActive=1 then%>
                    <% if LSCB_STATUS="0" OR LSCB_STATUS="" then %>
                        <div class="pcCPmessage">Bing Cashback is currently turned &quot;Off&quot;</div> 
                    <% else %>
                        <div class="pcCPmessageSuccess">Bing Cashback is currently turned &quot;On&quot;</div> 
                  <% end if %>
                   <p>&nbsp;</p> 
                <% else %>
                    <p>&nbsp;</p>
                    <p><span style="font-weight: bold">Bing Cashback is not Active.</span> 
                    To activate Bing Cashback click the &quot;Register for Merchant ID&quot; link below. Fill out the registration form and wait for your merchant key to arrive via email. Once you recieve your merchant key return to this page and click &quot;Activate Bing Cashback&quot;. On the following page save your merchant key and turn on Cashback tracking.
                    </p> 
                    <p>&nbsp;</p>	
                                 <p>For <a href="http://advertising.microsoft.com/advertising/cashback" target="_blank">details</a> about selling your products on Bing Cashback, <a href="http://advertising.microsoft.com/advertising/cashback" target="_blank">click here</a>.</p>
                    <p>&nbsp;</p>
                <%end if%>               
                
                <p>
                  <ul class="pcListIcon">
                    <%if pcv_strCashbackActive=1 then%>
                    <li><a href="pcCashback_settings.asp">View/Modify Settings</a></li>
                    <li><a href="exportCashBack.asp">Export products to Bing Cashback</a></li>
                    <% else %>
                    <li><a href="pcCashback_register.asp" target="_blank">Register for Merchant ID</a></li>
                    <li><a href="pcCashback_settings.asp">Activate Bing Cashback</a></li>
                    <%end if%>
                  </ul>
                </p>
            </td>
        </tr>
    </table>
    <script>
        function newWindow(file,window) {
            msgWindow=open(file,window,'resizable=yes,scrollbars=no,width=650,height=525');
            if (msgWindow.opener == null) msgWindow.opener = self;
        }
    </script>
<% else %>
    <form action="pcCashback_main.asp" method="post" name="IAgree" id="IAgree" class="pcForms">
        <table class="pcCPcontent">
            <tr> 
                <td colspan="2">
                <% if request.querystring("AM")<>"" then %>
                <div class="pcCPmessage">
                <%=request.querystring("AM")%>
                </div>
                <% end if %>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                	<iframe src="https://www.earlyimpact.com/productcart/microsoft/LiveSeachCashbackAgreement1008.html" width="700" marginwidth="2" height="150" marginheight="2" scrolling="auto" frameborder="0" hspace="2" vspace="20" style="border: 1px solid #e1e1e1;background-color: #F5F5F5;"></iframe>
              </td>
            </tr>
            <tr> 
                <td colspan="2">
                <input type="checkbox" name="agree" value="1" class="clearBorder"> I agree to the terms and conditions of the Microsoft Live Cashback End User License Agreement, which are listed above.
                </td>
            </tr>
            <tr> 
                <td colspan="2">
                <input type="submit" name="Submit2" value="Submit" class="submit2">
                </td>
            </tr>
        </table>
    </form> 
<% end if %>
<!--#include file="AdminFooter.asp"-->