<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Email Settings" 
pageIcon="pcv4_email_settings.png"
Section="layout"
%>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
dim mySQL, conntemp, rstemp, CustServEmail
on error resume next
call openDb()
mySQL="SELECT * FROM emailsettings WHERE id=1"

set rstemp=Server.CreateObject("ADODB.Recordset")     
set rstemp=conntemp.execute(mySQL)

	if err.number <> 0 then
		response.write "Error in emailSettings: "&Err.Description
	end If 

ownerEmail=rstemp("ownerEmail")
frmEmail=rstemp("frmEmail")
CustServEmail=scCustServEmail
	if trim(CustServEmail)="" then CustServEmail=frmEmail
ConfirmEmail=replace(rstemp("ConfirmEmail"),"<br>",vbCrlf)
ReceivedEmail=replace(rstemp("ReceivedEmail"),"<br>",vbCrlf)
ShippedEmail=replace(rstemp("ShippedEmail"),"<br>",vbCrlf)
CancelledEmail=replace(rstemp("CancelledEmail"),"<br>",vbCrlf)
PayPalEmail=replace(rstemp("PayPalEmail"),"<br>",vbCrlf)
%>

<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
function winemailobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=no,status=no,width=400,height=450')
	myFloater.location.href=fileName;
	}

function whatTripSelected(){
	if ( eval("document.form1.EmailComObj").value == 'CDOSYS' || eval("document.form1.EmailComObj").value == 'CDO' ){
		if (document.getElementById) {
			document.getElementById('divShowHide').style.display = ''; 
		} 
	}
	else{
		if (document.getElementById) {
		document.getElementById('divShowHide').style.display = 'none'; 
		} 
	}
}

//-->
</SCRIPT>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form  name="form1" method="post" action="../includes/PageCreateEmailSettings.asp" class="pcForms">
	<input type="hidden" name="Page_Name" value="emailsettings.asp">
	<input type="hidden" name="PayPalEmail2" value="<%=PayPalEmail%>">
   
	<table class="pcCPcontent">			
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Choose &amp; Test Your E-mail Component</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>				
		<tr> 
			<td align="right" width="20%" nowrap>Select an E-mail Component:</td>
			<td width="80%"> 
				<SELECT NAME="EmailComObj" ONCHANGE="whatTripSelected();">
					<option value="CDONTS" selected>CDONTS</option>
					<option value="CDOSYS" <% if scEmailComObj="CDOSYS" then%>selected<%end if%>>CDOSYS</option>
					<option value="CDO" <% if scEmailComObj="CDO" then%>selected<%end if%>>CDO</option>
					<option value="ABMailer" <% if scEmailComObj="ABMailer" then%>selected<%end if%>>ABMailer</option>
					<option value="Bamboo" <% if scEmailComObj="Bamboo" then%>selected<%end if%>>Bamboo SMTP</option>
					<option value="PersitsASPMail" <% if scEmailComObj="PersitsASPMail" then%>selected<%end if%>>Persits ASPMail</option>
					<option value="JMail3" <% if scEmailComObj="JMail3" then%>selected<%end if%>>JMail 3.7</option>
					<option value="JMail4" <% if scEmailComObj="JMail4" then%>selected<%end if%>>JMail 4</option>
					<option value="ServerObjectsASPMail" <% if scEmailComObj="ServerObjectsASPMail" then%>selected<%end if%>>ServerObjects ASPMail</option>        
				</select>
                &nbsp;
			<a href="javascript:winemailobj('EmailComCheck.asp')">Detect supported components</a>
            &nbsp;|&nbsp;
			<a href="javascript:winemailobj('EmailSettingsCheck.asp')">Test your settings</a>
			</td>
		</tr>
		<tr>
		  <td></td>
			<td class="pcSmallText">If CDOSYS is <a href="javascript:winemailobj('EmailComCheck.asp')">detected</a>, but cannot send e-mails, try using CDO instead.</td>
		</tr>					
		<tr>				
			<td align="right">SMTP Server Address:</td>
			<td><input type="text" name="SMTP" value="<%=scSMTP%>" size="40"></td>
		</tr>
        <% 
		if scEmailComObj="CDOSYS" or scEmailComObj="CDOSYS" then
			divShowHideStyle=""
		else
			divShowHideStyle="none"
		end if
		%>
		<tr id="divShowHide" style="display: <%=divShowHideStyle%>;">
		    <td valign="top" align="right">SMTP Settings</td>
		    <td>
                <input type="radio" name="optLocalRemote" value="1" <%if scLocalOrRemote = "1" then%>checked<%end if%> class="clearBorder"> Local 
                <input type="radio" name="optLocalRemote" value="2" <%if scLocalOrRemote <> "1" then%>checked<%end if%> class="clearBorder"> Remote 
                <div class="pcSmallText" style="margin-top: 5px;">Indicates whether the SMTP server is on a different server than the Web site</div>
              	<div style="margin-top: 10px;">
                <% 
                Dim pcStrPort
                pcStrPort = scPort
                If trim(pcStrPort)="" then pcStrPort="25"
                %>
                Port: <input type="text" size="4" name="optPort" value="<%=pcStrPort%>"> 
                <span class="pcSmallText">(if 25 does not work, try 2525, 23 or 26)</span>
                </div> 
		    </td>
		</tr>	
        <tr>
        	<td colspan="2"><hr></td>
        </tr>				
		<tr valign="top"> 				
			<td align="right">SMTP Server Authentication</td>
			<td align="left">
            	<input name="SmtpAuth" type="checkbox" value="1" <% if scSMTPAuthentication="Y" then%>checked<% end if %> class="clearBorder">           	   
            	 My SMTP server requires authentication <span class="pcSmallText"> - Ask your Web host if unsure</span></td>
		</tr>
		<tr> 
			<td></td>
			<td>User name: <input type="text" name="SmtpAuthUID" value="<%=scSMTPUID%>" autocomplete="off"></td>
		</tr>
        <tr> 
            <td></td>
            <td>Password: <input type="password" name="SmtpAuthPWD" value="<%=scSMTPPWD%>" autocomplete="off"></td>
        </tr>
        <tr>
        	<td colspan="2"><hr></td>
        </tr>
	<tr> 
        <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="Update " class="submit2">
        </td>
	</tr>
	<tr> 
        <td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">E-mail Addresses</th>
	</tr>
	<tr> 
	<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
	<td align="right">&quot;Store Manager's&quot; E-mail:</td>
	<td align="left"><input type="text" name="frmEmail" size="30" value="<%=frmEmail%>">&nbsp;<span class="pcSmallText">This is the address that will receive order confirmations, etc.</span></td>
	</tr>
	<tr> 
	<td align="right">&quot;Customer Service&quot; E-mail:</td>
	<td align="left"><input type="text" name="CustServEmail" size="30" value="<%=CustServEmail%>">&nbsp;<span class="pcSmallText">This is the address that will receive &quot;Contact Us&quot; and Help Desk notifications</span></td>
	</tr>
	<tr> 
	<td align="right">&quot;From&quot; E-mail:</td>
	<td align="left"><input type="text" name="ownerEmail" size="30" value="<%=ownerEmail%>">&nbsp;<span class="pcSmallText">This is the &quot;From&quot; address in all messages sent by the store.</span></td>
	</tr>
	<tr> 
	<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
	<td align="right"><input type="checkbox" name="NoticeNewCust" value="1" <%if scNoticeNewCust="1" then%>checked<%end if%> class="clearBorder"></td>
	<td align="left">Notify store manager when a new customer registers &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=404')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>	
	<tr> 
	<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
	<td colspan="2" align="center"> 
	<input type="submit" name="Submit" value="Update " class="submit2">
	</td>
	</tr>
	<tr> 
        <td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Default Messages</th>
	</tr>
	<tr> 
	<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>		
	<td align="right" valign="top">&quot;Order Received&quot; E-mail:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=405')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	<td><textarea name="ReceivedEmail" cols="60" rows="6"><%=ReceivedEmail%></textarea></td>
	</tr>					
	<tr> 		
	<td align="right" valign="top">&quot;Order Confirmation&quot; E-mail:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=406')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	<td><textarea name="ConfirmEmail" cols="60" rows="6"><%=ConfirmEmail%></textarea></td>
	</tr>
	<tr> 
	<td align="right" valign="top">Additional Copy for &quot;Order Shipped&quot; E-mail:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=407')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	<td><textarea name="ShippedEmail" cols="60" rows="6"><%=ShippedEmail%></textarea></td>
	</tr>					
	<tr> 
	<td align="right" valign="top">Additional Copy for &quot;Order Cancelled&quot; E-mail:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=408')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	<td><textarea name="CancelledEmail" cols="60" rows="6"><%=CancelledEmail%></textarea></td>
	</tr>
	<tr>
		<td colspan="2">						
		<p style="margin-bottom: 10px;">You may use any of the following <strong>dynamic fields</strong> in your e-mails.</p>
		<p>Your Company Name: <strong><font color="#FF0000">&lt;COMPANY&gt;</font></strong></p>
		<p>Company's URL: <strong><font color="#FF0000">&lt;COMPANY_URL&gt;</font></strong></p>
		<p>Today's Date: <strong><font color="#FF0000">&lt;TODAY_DATE&gt;</font></strong></p>
		<p>Customer's Full Name: <strong><font color="#FF0000">&lt;CUSTOMER_NAME&gt;</font></strong></p>
		<p>Order ID: <strong><font color="#FF0000">&lt;ORDER_ID&gt;</font></strong></p>
		<p>Order Date: <strong><font color="#FF0000">&lt;ORDER_DATE&gt;</font></strong></p>
	 </td>
	</tr>
    <tr>
        <td colspan="2"><hr></td>
    </tr>	
	<tr> 
	<td colspan="2" align="center"><input type="submit" name="Submit" value="Update " class="submit2">
	</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->