<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Notify Affiliate" %>
<% section="misc" %>
<%PmAdmin=8%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartfolder.asp"-->
<!--#include file="AdminHeader.asp"-->

<%

Dim rs, conntemp, query, pAffiliateId

pAffiliateId = getUserInput(request("idAffiliate"),4)
	if not validNum(pAffiliateId) then
		response.Redirect "AdminAffiliates.asp"
	end if
	
call openDb()
	query="SELECT affiliateName, affiliateCompany, commission, affiliateEmail, pcAff_Password FROM affiliates WHERE idaffiliate = " & pAffiliateId
	set rs=server.CreateObject("ADODB.Recordset")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
			response.redirect "techErr.asp?error="& Server.Urlencode("Error loading affiliate information: "&Err.Description) 
	end if
	if rs.eof then
			response.redirect "AdminAffiliates.asp"
	end if
	pcv_affiliateName = rs("affiliateName")
	pcv_affiliateCommission = rs("commission")
	pcv_affiliateEmail = rs("affiliateEmail")
	pcv_affiliatePass = rs("pcAff_Password")
	pcv_affiliatePass = enDeCrypt(pcv_affiliatePass, scCrypPass)
	pcv_affiliateSubject = "Your " & scCompanyName & " affiliate account is ready."
	set rs=nothing
call closeDb()

dim pcv_fromname, pcv_fromemail, pcv_toname, pcv_toemail, pcv_subject, pcv_message, pcv_success

pcv_FormAction=request.Form("pcFormAction")

if pcv_FormAction = "send" then

		pcv_fromemail=getUserInput(request.form("pcFromEmail"),0)
		pcv_fromname=getUserInput(request.form("pcFromName"),0)
		pcv_toname=getUserInput(request.form("pcAffName"),0)
		pcv_toemail=getUserInput(request.form("pcAffEmail"),0)
		pcv_message=getUserInput(request.form("pcEmailTestMessage"),0)
		pcv_subject=getUserInput(request.form("pcEmailSubject"),0)
		
		'// Write to the page for debugging
		'response.write "pcv_fromemail="&pcv_fromemail&"<br>" 
		'response.write "pcv_fromname="&pcv_fromname&"<br>" 
		'response.write "pcv_toname="&pcv_toname&"<br>" 
		'response.write "pcv_toemail="&pcv_toemail&"<br>" 
		'response.write "pcv_subject="&pcv_subject&"<br>" 
		'response.write pcv_message
		'response.End()
		'// End dubugging
		
		call sendmail (pcv_fromname, pcv_fromemail, pcv_toemail, pcv_subject, pcv_message)
%>
    <table class="pcCPcontent">
        <tr>
            <td class="pcCPspacer"></td>
        </tr>
        <tr> 
            <td><div class="pcCPmessageSuccess">The message was successfully sent to the affiliate. <a href="AdminAffiliates.asp">Manage Affiliates</a>.</div></td>
        </tr>
	</table>
<%
else

		pcv_affiliateLogin = replace((scStoreURL&"/"&scPcFolder&"/pc/AffiliateLogin.asp"),"//","/")
		pcv_affiliateLogin=replace(pcv_affiliateLogin,"http:/","http://")

%>

<form action="pcAffiliateSendEmail.asp" method="post" class="pcForms">
<input type="hidden" name="pcFormAction" value="send">
<input type="hidden" name="idAffiliate" value="<%=pAffiliateId%>">
    <table class="pcCPcontent">
        <tr>
            <td class="pcCPspacer" colspan="2"></td>
        </tr>
        <tr> 
            <th colspan="2">Send E-mail to Affiliate</th>
        </tr>
        <tr>
            <td class="pcCPspacer" colspan="2"></td>
        </tr>
        <tr>       
          <td align="right">From Name:</td>
          <td><input type="text" value="<%=scCompanyName%>" name="pcFromName" size="30"></td>
        </tr>
        <tr>
          <td align="right">From Email:</td>
            <td><input type="text" value="<%=scEmail%>" name="pcFromEmail" size="30"></td>
          </tr>
          <tr> 
            <td align="right">Affiliate Name:</td>
            <td><input type="text" value="<%=pcv_affiliateName%>" name="pcAffName" size="30"></td>
          </tr>
          <tr> 
            <td align="right">Affiliate Email:</td>
            <td><input type="text" value="<%=pcv_affiliateEmail%>" name="pcAffEmail" size="30"></td>
          </tr>
          <tr> 
            <td align="right">Subject:</td>
            <td><input type="text" value="<%=pcv_affiliateSubject%>" name="pcEmailSubject" size="50"></td>
          </tr>
          <tr> 
            <td align="right">Message:</td>
            <td></td>
          </tr>
          <tr> 
            <td colspan="2" align="center"><textarea name="pcEmailTestMessage" cols="80" rows="6">Dear <%=pcv_affiliateName%>, 
            
            your affiliate account is ready for use. Your User Name is your e-mail address (<%=pcv_affiliateEmail%>) and your Password is <%=pcv_affiliatePass%>.  You can log into your account at <%=pcv_affiliateLogin%> to retrieve special links to our store catalog that contain your affiliate ID.
            
            Best Regards,
            
            The <%=scCompanyName%> Team
            </textarea></td>
          </tr>
          <tr> 
            <td colspan="2" align="center">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" align="center"><input type="submit" value="Send Message" class="submit2"></td>
          </tr>
      </table>
  </form>

<% end if %>
<!--#include file="AdminFooter.asp"-->