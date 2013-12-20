<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Google Checkout Configuration" %>
<% Section="paymntOpt" %>
<%PmAdmin=1%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/rc4_GoogleCheckout.asp" -->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"--> 
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<style>

	#pcCPmain ul {
		margin: 0px;
		padding: 0;
	}

	#pcCPmain ul li {
		margin: 0px;
	}

	div.TabbedMenu ul {
	text-align:left;
	margin:0 0 0 60px;
	padding:0;
	cursor:pointer;
	}

	div.TabbedMenu ul li {
	display:inline;
	list-style:none;
	margin:0 0.3em;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a {
	position:relative;
	z-index:0;
	font-weight:bold;
	border:solid 2px #e1e1e1;
	border-bottom-width:0;
	padding:0.3em;
	background-color:#ffffcc;
	color:black;
	text-decoration:none;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a.current {
	background-color:#F5F5F5;
	border:solid 2px #CCCCCC;
	border-bottom-width:0;
	position:relative;z-index:2;
	cursor:pointer;
	font-size:12px;
	}
	
	div.TabbedMenu ul li a.current:hover {
	background-color:#F5F5F5;
	cursor:pointer;
	font-size:12px;
	}

	div.TabbedMenu ul li a:hover {
	z-index:2;
	background-color:#F5F5F5;
	border-bottom:0;
	cursor:pointer;
	font-size:12px;
	}
	
	div.TabbedMenu a span {display:none;}
	
	div.TabbedMenu a:hover span {
		display:block;
		position:absolute;
		top:2.3em;
		background-color:#F5F5F5;
		border-bottom:thin dotted gray;
		border-top:thin dotted gray;
		font-weight:normal;
		left:0;
		padding:1px 2px;
		cursor:pointer;
		font-size:12px;
	}
	
	div.TabbedPanes {
		padding: 1em;
		border: dashed 2px #CCCCCC;
		background-color: #F5F5F5;
		display: none;
		text-align:left;
		position:relative;z-index:1;
		margin-top:0.15em;
	}
	
</style>
<% 
'// Define Tab Count
dim k, pcTabCount, strTabCnt
pcTabCount=2
strTabCnt=""
for k=1 to pcTabCount
	if k=1 then
		strTabCnt=strTabCnt&"""tab"&k&""""
	else
		strTabCnt=strTabCnt&",""tab"&k&""""
	end if
next
%>

<!--#include file="../includes/javascripts/pcCPTabs.asp"-->

<% 
Dim connTemp, qry_ID, query, rs, mySQL, rstemp

call openDB()


pcv_intGoogleActive=GOOGLEACTIVE

pcStrPageName="ConfigureGoogleCheckout2.asp"

'// Set Required Fields
pcv_isMerchantIDRequired=true
pcv_isMerchantKeyRequired=true
pcv_isGoogleCurrencyRequired=true

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF Request.Form("submit1")<>"" THEN
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))


	pcs_ValidateTextField	"merchantID", pcv_isMerchantIDRequired, 0
	pcs_ValidateTextField	"merchantKey", pcv_isMerchantKeyRequired, 0
	pcs_ValidateTextField	"SandboxMerchantID", false, 0
	pcs_ValidateTextField	"SandboxMerchantKey", false, 0
	pcs_ValidateTextField	"GoogleTestMode", false, 0
	pcs_ValidateTextField	"GoogleCurrency", pcv_isGoogleCurrencyRequired, 0
	pcs_ValidateTextField	"GoogleTaxShipping", false, 0
	pcs_ValidateTextField	"pcv_processOrder", false, 0
	pcs_ValidateTextField	"pcv_setPayStatus", false, 0
	
	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// Insert the Google Payment Status 
		query="SELECT idPayment FROM paytypes WHERE gwCode=50;"
		set rsGoogle=Server.CreateObject("ADODB.Recordset")     
		set rsGoogle=connTemp.execute(query)		
		if rsGoogle.eof then
			'// Insert the Google Payment Status 
			query="INSERT INTO payTypes (paymentDesc, sslURL, active, quantityFrom, quantityUntil, weightFrom, weightUntil, priceFrom, priceUntil,ssl,priceToAdd, percentageToAdd,gwCode,Type) VALUES ('Google Checkout','',-1,0,9999,0,9999,0,9999,-1,0,0,'50','G')"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			set rs=nothing
		else
			'// Update the Google Payment Status 
			query="UPDATE payTypes SET active=-1 WHERE gwcode=50;"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
		set rsGoogle=nothing
		
		Session("pcAdminmerchantKey")=GDeCrypt(Session("pcAdminmerchantKey"), scCrypPass)
		Session("pcAdminSandboxMerchantKey")=GDeCrypt(Session("pcAdminSandboxMerchantKey"), scCrypPass)
		
		'// Redirect to Save the Constants	
		response.redirect("../includes/PageCreateGoogleCheckoutConstants.asp")
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	End If
ELSE
	'// Load values from ../includes/PageCreateGoogleCheckoutConstants.asp
	if msg="" then 
		Session("pcAdminmerchantID")=GOOGLEMERCHANTID
		Session("pcAdminmerchantKey")=GDeCrypt(GOOGLEMERCHANTKEY, scCrypPass)
		Session("pcAdminSandboxMerchantID")=GOOGLESANDBOXID
		Session("pcAdminSandboxMerchantKey")=GDeCrypt(GOOGLESANDBOXKEY, scCrypPass)
		Session("pcAdminGoogleCurrency")=GOOGLECURRENCY
		Session("pcAdminGoogleTaxShipping")=GOOGLETAXSHIPPING
		Session("pcAdminGoogleTestMode")=GOOGLETESTMODE
		Session("pcAdminpcv_processOrder")=GOOGLEPROCESS
		Session("pcAdminpcv_setPayStatus")=GOOGLEPAYSTATUS
		Session("pcAdminDelete")="NO"
	end if	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Delete Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if Request("mode")="Del" then

	'// Update the Google Payment Status 
	query="UPDATE payTypes SET active=0 WHERE gwcode=50;"
	set rs=Server.CreateObject("ADODB.Recordset")     
	set rs=connTemp.execute(query)
	set rs=nothing

	'// Load values from ../includes/PageCreateGoogleCheckoutConstants.asp
	Session("pcAdminmerchantID")=GOOGLEMERCHANTID
	Session("pcAdminmerchantKey")=GOOGLEMERCHANTKEY
	Session("pcAdminSandboxMerchantID")=GOOGLESANDBOXID
	Session("pcAdminSandboxMerchantKey")=GOOGLESANDBOXKEY
	Session("pcAdminGoogleCurrency")=GOOGLECURRENCY
	Session("pcAdminGoogleTaxShipping")=GOOGLETAXSHIPPING
	Session("pcAdminGoogleTestMode")=GOOGLETESTMODE
	Session("pcAdminpriceToAddType")=GOOGLEPRICETOADDTYPE
	Session("pcAdminpriceToAdd")=GOOGLEPRICETOADD
	Session("pcAdminpercentageToAdd")=GOOGLEPERCENTAGETOADD
	Session("pcAdminpcv_processOrder")=GOOGLEPROCESS
	Session("pcAdminpcv_setPayStatus")=GOOGLEPAYSTATUS
	Session("pcAdminDelete")="YES"
	
	'// Redirect to Save the Constants	
	response.redirect("../includes/PageCreateGoogleCheckoutConstants.asp")
	
end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Delete Mode
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"merchantID", pcv_isMerchantIDRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"merchantKey", pcv_isMerchantKeyRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
<script>
<!--
function openshipwin(fileName)
{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=500,height=350')
	myFloater.location.href=fileName;
	checkwin();
}
function checkwin()
{
	if (myFloater.closed)
	{
		location="Orddetails.asp?id=<%=qry_ID%>&ActiveTab=2";
	}
	else
	{
		setTimeout('checkwin()',500);
	}
}
function CalPop(sInputName)
{
	window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
}
function isDigit(s)
{
var test=""+s;
var OK2reset;
if(test=="."||test==","||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
function checkRandomly()
{
  //the next line generates a random number between 0 and 
  //checkAr.length - 1
  var intRandom = floor(Math.random() * checkAr.length);

  for (var i = 0; i <= intRandom; i++)
  {
    var myElem = document.myForm.elements[checkAr[i]]
    
    if (!myElem.checked)    
      myElem.checked = true;
    else
      myElem.checked = false;
  }
}
function disableCheckBox (checkBox) {
  if (!checkBox.disabled) {
		checkBox.checked = false;
    checkBox.disabled = true;
    if (!document.all && !document.getElementById) {
      checkBox.storeChecked = checkBox.checked;
      checkBox.oldOnClick = checkBox.onclick;
    }
  }
}
function enableCheckBox (checkBox) {
  if (checkBox.disabled) {
    checkBox.disabled = false;
    if (!document.all && !document.getElementById)
      checkBox.onclick = checkBox.oldOnClick;
  }
}
//-->
</script>

<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer">
			<% 
			msg=getUserInput(request.querystring("msg"),0)
			if msg<>"" then
			%>
				<div class="pcCPmessage">
					<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=msg%>
				</div>
			<% end if %>
		</td>
	</tr>
</table>
<% dim intActiveTab, pcTab1Style, pcTab2Style, pcTab3Style, pcTab4Style, pcTab5Style

intActiveTab=request("ActiveTab")
if intActiveTab="" then
	intActiveTab=1
end if

pcTab1Style="display:none"
pcTab2Style="display:none"
pcTab1Class=""
pcTab2Class=""

select case intActiveTab
	case "1"
		pcTab1Style="display:block"
		pcTab1Class="current"
	case "2"
		pcTab2Style="display:block"
		pcTab2Class="current"
	case else
		pcTab1Style="display:block"
		pcTab1Class="current"
end select
%>

<form id="form2" name="form2" method="post" action="<%=pcStrPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="ActiveTab" value="<%=intActiveTab%>">
	<table class="pcCPcontent">	
		<tr>
			<td valign="top">
				<div class="TabbedMenu">
					<ul>
						<li><a id="tabs1" class="<%=pcTab1Class%>" onclick="change('tabs1', 'current');change('tabs2', '');showTab('tab1');form2.ActiveTab.value = 1">Settings</a></li>
						<li><a id="tabs2" class="<%=pcTab2Class%>" onclick="change('tabs1', '');change('tabs2', 'current');showTab('tab2');form2.ActiveTab.value = 2">Read Me</a></li>						
					</ul>
				</div>
				
				<%
				'--------------
				' START TAB 1
				'--------------
				%>
				<div id="tab1" class="TabbedPanes" style="<%=pcTab1Style%>">
					<table class="pcCPcontent">
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Google Checkout Status<a name="top"></a></th>
						</tr>
						<tr> 
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2">
							<% if pcv_intGoogleActive=-1 then %>
								<p>
								<span class="pcCPnotes"><strong>Google Checkout is enabled. </strong></span>
								A button for Google Checkout will appear on your "View Cart" page. If you need to disable Google Checkout click the button below.
								<br />
								<br />
								<input name="submit2" type="button" value="Disable Google Checkout" class="submit2" onClick="location.href='ConfigureGoogleCheckout2.asp?mode=Del'">
								</p>
							<% else %>
								<p><span class="pcCPnotes"><strong>Google Checkout is not enabled.</strong></span></p>
							<% end if %>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Set my &quot;Google Merchant ID&quot; and &quot;Google Merchant Key&quot;</th>
						</tr>
						<tr> 
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2"><p><strong>LIVE MODE</strong>:&nbsp;&nbsp;<span class="pcCPnotes">Enter the Google credentials that you obtained from your Google Merchant account settings.</span></p></td>
						</tr>
						<tr> 
							<td width="267"><p>Google Merchant ID:</p></td>
							<td width="1039">
							<p>
							<input name="merchantID" type="text" value="<% =pcf_FillFormField ("merchantID", pcv_isMerchantIDRequired) %>" size="40" maxlength="50">
							<% pcs_RequiredImageTag "merchantID", pcv_isMerchantIDRequired %>
							</p>			</td>
						</tr>
						<tr> 
							<td><p>Google Merchant Key:</p></td>
							<td>
							<p>
							<input name="merchantKey" type="text" value="<% =pcf_FillFormField ("merchantKey", pcv_isMerchantKeyRequired) %>" size="40" maxlength="50">
							<% pcs_RequiredImageTag "merchantKey", pcv_isMerchantKeyRequired %>
							</p>			
							</td>
						</tr>	
						<tr> 
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr> 
							<td colspan="2"><p><strong>TEST MODE</strong>:&nbsp;&nbsp;<span class="pcCPnotes">Enter the Google credentials that you obtained from your Google &quot;Sandbox&quot; Merchant account settings.</span></p></td>
						</tr>
						<tr> 
							<td width="267"><p>Sandbox Merchant ID:</p></td>
							<td width="1039">
							<p>
							<input name="SandBoxMerchantID" type="text" value="<% =pcf_FillFormField ("SandBoxMerchantID", false) %>" size="40" maxlength="50">
							<% pcs_RequiredImageTag "SandBoxMerchantID", false %> (required for Test Mode only)
							</p>			</td>
						</tr>
						<tr> 
							<td><p>Sandbox Merchant Key:</p></td>
							<td>
							<p>
							<input name="SandBoxMerchantKey" type="text" value="<% =pcf_FillFormField ("SandBoxMerchantKey", false) %>" size="40" maxlength="50">
							<% pcs_RequiredImageTag "SandBoxMerchantKey", false %> (required for Test Mode only)
							</p>			
							</td>
						</tr>
						<tr>
							<td><p></p></td>
							<td>
							<p><input name="GoogleTestMode" type="checkbox" class="clearBorder" id="GoogleTestMode" value="YES" <%if Session("pcAdminGoogleTestMode")="YES" then%>checked<%end if%>>
							<b>Enable Test Mode </b>(Credit cards will not be charged. Requires Sandbox Merchant ID and Key.)							
								</p>			
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<th colspan="2">Miscellaneous Settings</th>
						</tr>
						<tr> 
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
						<tr>
							<td><p>Currency:</p></td>
							<td>
							<p>
							<select name="GoogleCurrency">
								<option value="USD" <% if Session("pcAdminGoogleCurrency")="USD" then%>selected<% end if %>>U.S. Dollars ($)</option>
								<option value="GBP" <% if Session("pcAdminGoogleCurrency")="GBP" then%>selected<% end if %>>Pounds Sterling (&pound;)</option>													
							</select>&nbsp;&nbsp;<span class="pcCPnotes"><i>Note: Pounds Sterling required for Google Checkout <strong>UK</strong>.</i></span>						
							</p>			
							</td>
						</tr>
						<tr>					
							<td><p>Shipping Handling Charges include VAT?</p></td>
							<td>
								<p>
								<input type="radio" name="GoogleTaxShipping" value="false" checked> No
								<input type="radio" name="GoogleTaxShipping" value="true" <% If Session("pcAdminGoogleTaxShipping")="true" then%>checked<% end if %>> 
								Yes&nbsp;&nbsp;
								If set to &quot;Yes&quot;, ProductCart assumes that Shipping &amp; Handling charges include VAT, and Google Checkout will be instructed to include these charges. 
								<br />
								<span class="pcCPnotes"><i>Note: This setting only applies to Google Checkut <strong>UK</strong>.</i></span>
								</p>
							</td>
						</tr>	
						<tr> 
							<td colspan="2" class="pcCPspacer"></td>
						</tr>					
		
						<tr>
							<td colspan="2" class="pcCPspacer">
							<input name="pcv_processOrder" type="hidden" value="0" />
							<input name="pcv_setPayStatus" type="hidden" value="0" />
							</td>
						</tr>

						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
					</table>
				</div>					
				<%
				'--------------
				' END TAB 1
				'--------------
				
				'--------------
				' START TAB 2
				'--------------
				%>					
				<div id="tab2" class="TabbedPanes" style="<%=pcTab2Style%>">
				<table class="pcCPcontent">
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<th colspan="2">Using Google Checkout with ProductCart</th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<td colspan="2" valign="top">
						<p>The following links contain important reference material for using Google Checkout with ProductCart.</strong></p>
						<ul class="pcListIcon" style="margin: 10px 0 0 30px;">
						  <li><a href="http://wiki.earlyimpact.com/widgets/integrations/googlecheckout" target="_blank">Using Google Checkout with ProductCart.</a></li>
						  <li><a href="http://checkout.google.com/support" target="_blank">Google Checkout Help Center.</a></li>
						</ul>
						</td>
					 </tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
						<th colspan="2">Default Shipping Rates</th>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr>
					  <td colspan="2" valign="top"> 
						<p>
						Google Checkout will determine shipping rates for the order by dynamically exchanging information with your ProductCart-powered store, invisibly to the customer.	To prevent a situation in which there are no shipping rates to choose from (e.g. a communication problem between your Web server and Google&rsquo;s servers), Google Checkout requires that your store provide a set of default shipping rates. ProductCart automatically calculates reasonable rates based a number of parameters.&nbsp; Advanced users have the option to override the ProductCart default rates with personalized rates. </p>
						<p>&nbsp;</p>
						<p><a href="GoogleCheckout_DefaultRates.asp" target="_blank">Advanced Settings: Override Default Shipping Rates</a>
						  </p>
						  </p>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
				</table>
				</div>
				<%
				'--------------
				' END TAB 2
				'--------------
				%>	
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="center">			
			<% if pcv_intGoogleActive=-1 then %>
				<input name="submit1" type="submit" value="Save Settings" class="submit1">&nbsp;
			<% else %>
				<input name="submit1" type="submit" value="Save Settings and Enable Google Checkout" class="submit1">&nbsp;
			<% end if %>
				<input type="button" value="Payment Options" onClick="location.href='PaymentOptions.asp'" class="submit2">			
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	</table>
<% 
call closedb() 
Session("pcAdminmerchantID")=""
Session("pcAdminmerchantKey")=""
Session("pcAdminSandboxMerchantID")=""
Session("pcAdminSandboxMerchantKey")=""
Session("pcAdminGoogleTestMode")=""
Session("pcAdminpriceToAddType")=""
Session("pcAdminpriceToAdd")=""
Session("pcAdminpercentageToAdd")=""
Session("pcAdminpcv_processOrder")=""
Session("pcAdminpcv_setPayStatus")=""
Session("pcAdminGoogleCurrency")=""
%>
</form>

<!--#include file="AdminFooter.asp"-->