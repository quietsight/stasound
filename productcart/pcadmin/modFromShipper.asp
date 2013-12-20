<%
response.Buffer=true
Response.Expires = -1
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
pageTitle="General Shipping Settings"
pageIcon="pcv4_icon_ship.png"
Section="shipOpt"
%>
<%PmAdmin=4%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->  
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->


<%
dim mySQL, conntemp, rstemp
call openDb()
'// DMB - Added below this line
pcStrPageName="modFromShipper.asp"
'// Set Required Fields
pcv_isnameRequired=true
pcv_iscompanyRequired=true
pcv_isdepartmentRequired=false
pcv_iscountryRequired=true

'// Use the Request object to toggle State (based of Country selection)
pcv_isstateRequired=true
pcv_strStateCodeRequired=request("pcv_isStateCodeRequired")
if  len(pcv_strStateCodeRequired)>0 then
	pcv_isstateRequired=pcv_strStateCodeRequired
end if

'// Use the Request object to toggle Province (based of Country selection)
pcv_isprovinceRequired=false
pcv_strProvinceCodeRequired=request("pcv_isProvinceCodeRequired")
if  len(pcv_strProvinceCodeRequired)>0 then
	pcv_isprovinceRequired=pcv_strProvinceCodeRequired
end if

pcv_isaddress1Required=true
pcv_isaddress2Required=false  
pcv_isaddress3Required=false  
pcv_iscityRequired=true
pcv_ispostalcodeRequired=true
pcv_isphoneRequired=true
pcv_ispageRequired=false
pcv_isfaxRequired=false



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script language=""JavaScript"">"&vbcrlf
response.write "<!--"&vbcrlf	
response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf
pcs_JavaTextField	"pShipFromPersonName", pcv_isnameRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromName", pcv_iscompanyRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromPhone", pcv_isphoneRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromPostalCountry", pcv_iscountryRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromAddress1", pcv_isaddress1Required, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromCity", pcv_iscityRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
pcs_JavaTextField	"pShipFromPostalCode", pcv_ispostalcodeRequired, dictLanguage.Item(Session("language")&"_NewCust_3")
response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf
response.write "//-->"&vbcrlf
response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
IF request.Form("Submit")<>"" THEN
	'/////////////////////////////////////////////////////
	'// Validate Fields and Set Sessions	
	'/////////////////////////////////////////////////////
	
	'// set errors to none
	pcv_intErr=0
	
	'// generic error for page
	pcv_strGenericPageError = Server.Urlencode(dictLanguage.Item(Session("language")&"_Custmoda_18"))


	pcs_ValidateTextField	"pShipFromPersonName", pcv_isnameRequired, 50
	pcs_ValidateTextField	"pShipFromName", pcv_iscompanyRequired, 50
	pcs_ValidateTextField	"pShipFromDepartment", pcv_isdepartmentRequired, 50
	pcs_ValidateTextField	"pShipFromAddress1", pcv_isaddress1Required, 75
	pcs_ValidateTextField	"pShipFromAddress2", pcv_isaddress2Required, 75 
	pcs_ValidateTextField	"pShipFromAddress3", pcv_isaddress3Required, 75 	
	pcs_ValidateTextField	"pShipFromPostalCountry", pcv_iscountryRequired, 50
	pcs_ValidateTextField	"pShipFromState", pcv_isstateRequired, 50
	pcs_ValidateTextField	"pShipFromProvince", pcv_isprovinceRequired, 50
	pcs_ValidateTextField	"pShipFromCity", pcv_iscityRequired, 50
	pcs_ValidateTextField	"pShipFromPostalCode", pcv_ispostalcodeRequired, 12
	pcs_ValidateTextField	"pShipFromZip4", false, 4
	pcs_ValidatePhoneNumber	"pShipFromPhone", pcv_isphoneRequired, 30
	pcs_ValidatePhoneNumber	"pShipFromFax", pcv_isfaxRequired, 30
	pcs_ValidatePhoneNumber	"pShipFromPage", pcv_ispagerequired, 30
		
	pcs_ValidateTextField	"packageWeightLimit", false, 0	
	pcs_ValidateTextField	"DefaultProvider", false, 0
	pcs_ValidateTextField	"AlwAltShipAddress", false, 0
	pcs_ValidateTextField	"ComResShipAddress", false, 0
	pcs_ValidateTextField	"AlwNoShipRates", false, 0
	pcs_ValidateTextField	"pShowProductWeight", false, 0
	pcs_ValidateTextField	"sds_NotifySeparate", false, 0
	pcs_ValidateTextField	"pShowCartWeight", false, 0
	pcs_ValidateTextField	"pShowEstimateLink", false, 0
	pcs_ValidateTextField	"pHideProductPackage", false, 0
	
	pcs_ValidateTextField	"pHideEstimateDeliveryTimes", false, 0
	
	pcs_ValidateTextField	"sectionShow", false, 0
	pcs_ValidateTextField	"shipDetailTitle", false, 0
	pcs_ValidateHTMLField	"shipDetails", false, 0
	pcs_ValidateTextField	"ratesOnly", false, 0	
	
	'// Fix Quotes
	Session("pcAdminpShipFromName")=Replace(Session("pcAdminpShipFromName"),"&quot;","""""")
	Session("pcAdminpShipFromName")=Replace(Session("pcAdminpShipFromName"),"""","""""")

	'/////////////////////////////////////////////////////
	'// Check for Validation Errors
	'/////////////////////////////////////////////////////
	
	If pcv_intErr>0 Then
		response.redirect pcStrPageName&"?msg="&pcv_strGenericPageError
	Else
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' START: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		response.redirect("../includes/PageCreateShipFromSettings.asp")
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		' END: Run Code
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	End If
ELSE
	'// Load values from ../includes/shipFromSettings.asp
	msg=request.QueryString("msg")
	if msg="" then 
		Session("pcAdminpShipFromPersonName")=scOriginPersonName
		Session("pcAdminpShipFromName")=scShipFromName
		Session("pcAdminpShipFromDepartment")=scOriginDepartment
		Session("pcAdminpShipFromPhone")=scOriginPhoneNumber
		Session("pcAdminpShipFromPage")=scOriginPagerNumber
		Session("pcAdminpShipFromFax")=scOriginFaxNumber
		Session("pcAdminpShipFromAddress1")=scShipFromAddress1
		Session("pcAdminpShipFromAddress2")=scShipFromAddress2
		Session("pcAdminpShipFromAddress3")=scShipFromAddress3
		Session("pcAdminpShipFromCity")=scShipFromCity
		Session("pcAdminpShipFromPostalCode")=scShipFromPostalCode
		Session("pcAdminpShipFromZip4")=scShipFromZip4
		Session("pcAdminpShipFromProvince")=scShipFromState
		Session("pcAdminpShipFromState")=scShipFromState	
		Session("pcAdminpShipFromPostalCountry")=scShipFromPostalCountry
	end if	
END IF	
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: POSTBACK
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// DMB - Added above this line
%>
<SCRIPT>
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
      //checkBox.onclick = preserve;
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
</SCRIPT>

<form name="fromShipper" method="post" action="<%=pcStrPageName%>" onSubmit="return Form1_Validator(this)" class="pcForms">  
	<input type="hidden" name="gsShipFromName" value="<%=Replace(scCompanyName, """", "&quot;")%>">
	<input type="hidden" name="gsShipCompanyName" value="<%=Replace(scCompanyName, """", "&quot;")%>">
	<input type="hidden" name="gsShipFromAddress1" value="<%=scCompanyAddress%>">
	<input type="hidden" name="gsShipFromCity" value="<%=scCompanyCity%>">
	<input type="hidden" name="gsShipFromState" value="<%=scCompanyState%>">
	<input type="hidden" name="gsShipFromPostalCountry" value="<%=scCompanyCountry%>">
	<input type="hidden" name="gsShipFromPostalCode" value="<%=scCompanyZip%>">		
	<input type="hidden" name="gsShipFromPersonName" value="<%=scOriginPersonName%>">
	<input type="hidden" name="gsShipFromDepartment" value="<%=scOriginDepartment%>">
	<input type="hidden" name="gsShipFromPhone" value="<%=scOriginPhoneNumber%>">
	<input type="hidden" name="gsShipFromPage" value="<%=scOriginPagerNumber%>">
	<input type="hidden" name="gsShipFromFax" value="<%=scOriginFaxNumber%>">

	<table class="pcCPcontent">
		<tr>
			<th colspan="2">&quot;Ship From&quot; Address</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
		</tr>
		<tr> 
			<td colspan="2">
				<p>Enter the address that you will be using as your &quot;Ship From&quot; address. This address must be recognized as a valid address by the shipping provider(s) that you will be using on this store.</p>
			</td>
		</tr>
		<% if scCompanyAddress <> "" AND scCompanyCountry <> "" AND scCompanyState <> "" then  %>
			<tr> 
				<td colspan="2">
					<p><input type="checkbox" name="C1" value="1" onClick="CopyInfo(this);" class="clearBorder"> Use the same address as &quot;Store Settings&quot;</p>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
		<script language="JavaScript">
		<!--
		function CopyInfo(theBox) {	
										
			var checkbox_status = theBox.checked
		
			if ( !checkbox_status )
			{
				theBox.checked = checkbox_status

				//document.forms['fromShipper'].elements['pShipFromPersonName'].value = ''
				document.forms['fromShipper'].elements['pShipFromName'].value = ''
				document.forms['fromShipper'].elements['pShipFromAddress1'].value = ''
				document.forms['fromShipper'].elements['pShipFromCity'].value = ''
				document.forms['fromShipper'].elements['pShipFromPostalCode'].value = ''
				document.forms['fromShipper'].elements['pShipFromPostalCountry'].value = ''
				document.forms['fromShipper'].elements['pShipFromProvince'].value = ''
				
			} else { 
			
				//document.forms['fromShipper'].elements['pShipFromPersonName'].value = 
				document.forms['fromShipper'].elements['pShipFromName'].value = document.forms['fromShipper'].elements['gsShipCompanyName'].value
				document.forms['fromShipper'].elements['pShipFromAddress1'].value = document.forms['fromShipper'].elements['gsShipFromAddress1'].value
				document.forms['fromShipper'].elements['pShipFromCity'].value = document.forms['fromShipper'].elements['gsShipFromCity'].value
				document.forms['fromShipper'].elements['pShipFromPostalCode'].value = document.forms['fromShipper'].elements['gsShipFromPostalCode'].value
				document.forms['fromShipper'].elements['pShipFromPostalCountry'].value = document.forms['fromShipper'].elements['gsShipFromPostalCountry'].value
				document.forms['fromShipper'].elements['pShipFromProvince'].value = document.forms['fromShipper'].elements['gsShipFromState'].value
				var a = document.forms['fromShipper'].elements['gsShipFromState'].value;
				SelectState('pShipFromPostalCountry', 'pShipFromState', 'pShipFromProvince', a, '');
			}
		}
		//-->
		</script>
		<% end if %>
		<tr> 
			<td><p>Shipping Contact Name:</p></td>
			<td>
                        <p>
                        <input name="pShipFromPersonName" type="text" value="<% =pcf_FillFormField ("pShipFromPersonName", pcv_isnameRequired) %>" size="20" maxlength="50">
                        <% pcs_RequiredImageTag "pShipFromPersonName", pcv_isnameRequired %>
                        </p>
                        </td>
		</tr>
                        <%
                        Session("pcAdminpShipFromName")=Replace(Session("pcAdminpShipFromName"), """", "&quot;")
                        %>
		<tr> 
			<td><p>Company Name:</p></td>
			<td>
                        <p>
                        <input name="pShipFromName" type="text" value="<% =pcf_FillFormField ("pShipFromName", pcv_iscompanyRequired) %>" size="20" maxlength="50">
                        <% pcs_RequiredImageTag "pShipFromName", pcv_iscompanyRequired %>
                        </p>
                        </td>
                    </tr>
                    <tr> 
                        <td><p>Department:</p></td>
                        <td>
                        <p>
                        <input name="pShipFromDepartment" type="text" value="<% =pcf_FillFormField ("pShipFromDepartment", pcv_isDepartmentRequired) %>" size="20" maxlength="50">
                        <% pcs_RequiredImageTag "pShipFromDepartment", pcv_isDepartmentRequired %>
                        </p>
                        </td>
		</tr>
		<%	'// Phone Custom Error
		if session("ErrpShipFromPhone")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">You must enter a valid phone number</td>
			</tr>
			<% session("ErrpShipFromPhone") = ""
		end if %>
		<tr> 
                        <td><p>Phone:</p></td>
                        <td>
                        <p>
                        <input name="pShipFromPhone" type="text" value="<% =pcf_FillFormField ("pShipFromPhone", pcv_isPhoneRequired) %>" size="20" maxlength="50">
                        <% pcs_RequiredImageTag "pShipFromPhone", pcv_isPhoneRequired %>
                        </p>
                        </td>
                    </tr>
		<%	'// Pager Custom Error
		if session("ErrpShipFromPage")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">You must enter a valid pager number</td>
			</tr>
			<% session("ErrpShipFromPage") = ""
		end if %>
                    <tr> 
			<td><p>Pager:</p></td>
                        <td>
                        <p>
			<input name="pShipFromPage" type="text" value="<% =pcf_FillFormField ("pShipFromPage", pcv_isPageRequired) %>" size="20" maxlength="50">
			<% pcs_RequiredImageTag "pShipFromPage", pcv_isPageRequired %>
			</p>
			</td>
		</tr>
						<%	'// Fax Custom Error
                        if session("ErrpShipFromFax")<>"" then %>
			<tr> 
				<td>&nbsp;</td>
				<td><img src="<%=pcf_GenerateIconURL("images/sample/pc_icon_next.gif")%>" width="10" height="10">You must enter a valid phone number</td>
			</tr>
						<% session("ErrpShipFromFax") = ""
                        end if %>
		<tr> 
			<td><p>Fax:</p></td>
                        <td>
                        <p>
			<input name="pShipFromFax" type="text" value="<% =pcf_FillFormField ("pShipFromFax", pcv_isFaxRequired) %>" size="20" maxlength="50">
			<% pcs_RequiredImageTag "pShipFromFax", pcv_isFaxRequired %>
                        </p>
                        </td>
                    </tr>
								<%
                                '///////////////////////////////////////////////////////////
                                '// START: COUNTRY AND STATE/ PROVINCE CONFIG
                                '///////////////////////////////////////////////////////////
                                ' 
                                ' 1) Place this section ABOVE the Country field
                                ' 2) Note this module is used on multiple pages. Transfer your local variable into this rountine via the section below.
                                ' 3) Additional Required Info
                                
                                '// #2 Transfer local variable into this rountine here. (Format: Required Variable = Local Variable)
                                pcv_isStateCodeRequired = pcv_isstateRequired '// determines if validation is performed (true or false)
                                pcv_isProvinceCodeRequired = pcv_isprovinceRequired '// determines if validation is performed (true or false)
                                pcv_isCountryCodeRequired = pcv_iscountryRequired '// determines if validation is performed (true or false)
                                
                                '// #3 Additional Required Info
                                pcv_strTargetForm = "fromShipper" '// Name of Form
                                pcv_strCountryBox = "pShipFromPostalCountry" '// Name of Country Dropdown
                                pcv_strTargetBox = "pShipFromState" '// Name of State Dropdown
                                pcv_strProvinceBox =  "pShipFromProvince" '// Name of Province Field
                                
                                '// Set local Country to Session
                                if Session(pcv_strSessionPrefix&pcv_strCountryBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strCountryBox) = pShipFromPostalCountry
                                end if
                                
                                '// Set local State to Session
                                if Session(pcv_strSessionPrefix&pcv_strTargetBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strTargetBox) = pShipFromState
                                end if
                                
                                '// Set local Province to Session
                                if Session(pcv_strSessionPrefix&pcv_strProvinceBox) = "" then
                                    Session(pcv_strSessionPrefix&pcv_strProvinceBox) = pShipFromProvince
                                end if
                                %>					
                                <!--#include file="../includes/javascripts/pcStateAndProvince.asp"-->
                                <%
                                '///////////////////////////////////////////////////////////
                                '// END: COUNTRY AND STATE/ PROVINCE CONFIG
                                '///////////////////////////////////////////////////////////
                                %>		
                                <%
                                '// Display the Country Dropdown (/includes/javascripts/pcStateAndProvince.asp)
                                pcs_CountryDropdown
                        
                                %>	
                    <tr> 
                        <td><p>Address:</p></td>
                        <td> 
                        <p>
                        <input name="pShipFromAddress1" type="text" value="<% =pcf_FillFormField ("pShipFromAddress1", pcv_isAddress1Required) %>" size="30" maxlength="75">
                        <% pcs_RequiredImageTag "pShipFromAddress1", pcv_isAddress1Required %>
                        </p>
                        </td>
		</tr>
		<tr> 
                        <td>&nbsp;</td>
                        <td> 
                        <p>
                        <input type="text" name="pShipFromAddress2" value="<% =pcf_FillFormField ("pShipFromAddress2", pcv_isAddress2Required) %>" size="30" maxlength="75">
                        <% pcs_RequiredImageTag "pShipFromAddress2", pcv_isAddress2Required %>
                        </p>
                        </td>
                    </tr>
                    <tr> 
                        <td>&nbsp;</td>
			<td> 
                        <p>
                        <input type="text" name="pShipFromAddress3" value="<% =pcf_FillFormField ("pShipFromAddress3", pcv_isAddress3Required) %>" size="30" maxlength="75">
                        <% pcs_RequiredImageTag "pShipFromAddress3", pcv_isAddress3Required %>
                        </p>
                        </td>
                    </tr>
                    <tr> 
                        <td><p>City:</p></td>
                        <td>
                        <p>
                        <input type="text" name="pShipFromCity" value="<% =pcf_FillFormField ("pShipFromCity", pcv_isCityRequired) %>" size="20" maxlength="50">
                        <% pcs_RequiredImageTag "pShipFromCity", pcv_isCityRequired %>
                        </p>
                        </td>
		</tr>
		
		<%
		'// Display the State/ Province Zones (/includes/javascripts/pcStateAndProvince.asp)
		pcs_StateProvince
		%>	
			
		<tr> 
                        <td><p>Postal Code:</p></td>
                        <td> 
                        <p>
                        <input name="pShipFromPostalCode" type="text" value="<% =pcf_FillFormField ("pShipFromPostalCode", pcv_isPostalCodeRequired) %>" size="10" maxlength="10">
                        <% pcs_RequiredImageTag "pShipFromPostalCode", pcv_isPostalCodeRequired %>-<input name="pShipFromZip4" type="text" value="<% =pcf_FillFormField ("pShipFromZip4", pcv_isPostalCodeRequired) %>" size="10" maxlength="10">
                        <% pcs_RequiredImageTag "pShipFromZip4", false %>
                        *Last 4 digits required to use USPS labels wizard</p>
                        </td>
                    </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Maximum Weight per Package</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
        <tr>
            <td colspan="2">
                <p>Enter the Maximum Weight per Package (if any): <input name="packageWeightLimit" type="text" id="packageWeightLimit" size="4" maxlength="4" value="<%=scPackageWeightLimit%>"> <%=scShipFromWeightUnit%> &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=474')"><img src="images/pcv3_infoIcon.gif" alt="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_400")%>"></a></p>
            </td>
        </tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="2">Other settings</th>
	</tr>
	<tr> 
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<table class="pcCPcontent">
      <%
				intShipTypeCnt=0
				strShipTypeOpt=""
				dim query, rs
				query="SELECT serviceCode,serviceActive,serviceDescription FROM shipService WHERE serviceActive=-1;"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if NOT rs.eof then
					do until rs.eof
						if instr(rs("serviceCode"),"C") then
							if scDefaultProvider = "CUSTOM" then
								if  intShipTypeCnt=0 then
									strShipTypeOpt=strShipTypeOpt&"<option value=CUSTOM selected>Custom Shipping Options</option>"
								end if
							else
								if  intShipTypeCnt=0 then
									strShipTypeOpt=strShipTypeOpt&"<option value=CUSTOM>Custom Shipping Options</option>"
								end if
							end if
							intShipTypeCnt=1
						end if
						rs.moveNext
					loop
				end if
				query="SELECT shipmentDesc FROM shipmentTypes WHERE active<>0;"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if NOT rs.eof then
					do until rs.eof
						strTempShipmentDesc=rs("shipmentDesc")
						strTempShipmentDescShown=strTempShipmentDesc
						if strTempShipmentDesc="Canada Post" then
							strTempShipmentDesc="CP"
							strTempShipmentDescShown="CANADA POST"
						end if
						intShipTypeCnt=intShipTypeCnt+1
						if scDefaultProvider = strTempShipmentDesc then
							strShipTypeOpt=strShipTypeOpt&"<option value="&strTempShipmentDesc&" selected>"&strTempShipmentDescShown&"</option>"
						else
							strShipTypeOpt=strShipTypeOpt&"<option value="&strTempShipmentDesc&">"&strTempShipmentDescShown&"</option>"
						end if
						rs.moveNext
					loop
				end if
				call closedb()
				set rs=nothing
				if intShipTypeCnt>1 then %>
					<tr> 
						<td colspan="2" valign="top">
							<p>Select a default shipping provider:&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=426')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p>
						</td>
					</tr>
                    <tr> 
						<td colspan="2" valign="top">
							<p><select name="DefaultProvider"><%=strShipTypeOpt%></select></p></td>
					</tr>
                    <tr>
                        <td colspan="2"><hr></td>
                    </tr>
				<%
				end if
				%>
				<tr> 
					<td align="right" valign="top"><input type="checkbox" name="AlwNoShipRates" value="-1" <% if scAlwNoShipRates="-1" then%>checked<% end if %> class="clearBorder"></td>
					<td>Allow customers to complete order if no shipping rates are returned. 
                    <div class="pcSmallText">This will allow you to manually calculate and add the shipping rates before actual shipment of the order.</div>
                    </td>
				</tr>
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr> 
                    <th colspan="2">Shipping Address and Address Type</th>
                </tr>
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr> 
                    <td colspan="2">When to show the Shipping Address:</td>
                </tr>
				<tr> 
					<td align="right"><input type="radio" name="AlwAltShipAddress" value="0" <% if scAlwAltShipAddress<>"1" OR scAlwAltShipAddress<>"2" then%>checked<% end if %> class="clearBorder"></td>
					<td>Show shipping address only if order requires shipping</td>
				</tr>
				<tr> 
					<td align="right"><input type="radio" name="AlwAltShipAddress" value="1" <% if scAlwAltShipAddress="1" then%>checked<% end if %> class="clearBorder"></td>
					<td>Disable shipping address (billing address = shipping address)</td>
				</tr>
				<tr> 
					<td align="right"><input type="radio" name="AlwAltShipAddress" value="2" <% if scAlwAltShipAddress="2" then%>checked<% end if %> class="clearBorder"></td>
					<td>Always show shipping address (even if the order does not require shipping)</td>
				</tr>
                <tr> 
                    <td colspan="2"><hr></td>
                </tr>
                <tr> 
                    <td colspan="2">Commercial vs. Residential &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=475')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
                </tr>
				<tr> 
					<td align="right"><input type="radio" name="ComResShipAddress" value="0" <% if scComResShipAddress="0" then%>checked<% end if %> class="clearBorder"></td>
					<td>Let the customer choose</td>
				</tr>
				<tr> 
					<td align="right"><input type="radio" name="ComResShipAddress" value="1" <% if scComResShipAddress="1" then%>checked<% end if %> class="clearBorder"></td>
					<td>Always use <em>Residential</em></td>
				</tr>
				<tr> 
					<td align="right"><input type="radio" name="ComResShipAddress" value="2" <% if scComResShipAddress="2" then%>checked<% end if %> class="clearBorder"></td>
					<td>Always use <em>Commercial</em></td>
				</tr>
				<tr> 
					<td align="right"><input type="radio" name="ComResShipAddress" value="3" <% if scComResShipAddress="3" then%>checked<% end if %> class="clearBorder"></td>
					<td>Use <em>Commercial</em> with wholesale customers, <em>Residential</em> with retail customers</td>
				</tr>
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr> 
                    <th colspan="2">Shipping-related Display Settings</th>
                </tr>
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
				<tr> 
					<td align="right" valign="top"><input type="checkbox" name="sds_NotifySeparate" value="1" <%if scShipNotifySeparate="1" then%>checked<%end if%> class="clearBorder"></td>
					<td>Notify customers when order might be shipped in separate shipments.
                    <div class="pcSmallText">You can let customers decide whether to receive one or multiple shipments using the <a href="AdminSettings.asp?tab=2" target="_blank">Allow Separate Shipments</a> setting.</div></td>
				</tr>
				<tr> 
					<td align="right"><input type="checkbox" name="pShowProductWeight" value="-1" <% if scShowProductWeight="-1" then%>checked<% end if %> class="clearBorder"> 
					</td>
					<td>Display product weight on the product's detail page.</td>
				</tr>
				<tr> 
					<td align="right"><input type="checkbox" name="pShowCartWeight" value="-1" <% if scShowCartWeight="-1" then%>checked<% end if %> class="clearBorder"> 
					</td>
					<td>Display total cart weight on view cart page.</td>
				</tr>
				<tr> 
					<td align="right"><input type="checkbox" name="pShowEstimateLink" value="-1" <% if scShowEstimateLink="-1" then%>checked<% end if %> class="clearBorder"> 
					</td>
					<td>Display 'Estimated Shipping Charges' link on the View Shopping Cart page.</td>
				</tr>
				<tr> 
					<td valign="top" align="right"><input type="checkbox" name="pHideProductPackage" value="-1" <% if scHideProductPackage="-1" then%>checked<% end if %> class="clearBorder"> 
					</td>
					<td>Hide number of packages on shipping service selection page.
                    <div class="pcSmallText">Note that the number of packages is only shown when the order is shipped in 2 or more packages.</div>
                    </td>
				</tr>
				<tr> 
					<td valign="top" align="right"><input type="checkbox" name="pHideEstimateDeliveryTimes" value="-1" <% if scHideEstimateDeliveryTimes="-1" then%>checked<% end if %> class="clearBorder"> 
					</td>
					<td>
					    Hide Estimated Delivery Time
                        <div class="pcSmallText">This is the center column in the table that shows available shipping services in the storefront.</div>
                    </td>
				</tr>					
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr> 
                    <th colspan="2">Shipping Instructions / Disclaimer</th>
                </tr>
                <tr> 
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
				<tr> 
					<td colspan="2" valign="top">Display instructions, a disclaimer, or other shipping-related information on the page on which customers select shipping. For example, your message could explain that &quot;In store pickup&quot; is not available on &quot;Sundays&quot;.</td>
				</tr>
        <tr> 
          <td colspan="2" valign="top">
					<table class="pcCPcontent">
            			<tr> 
							<td colspan="2">
							<table class="pcCPcontent">
									<tr> 
										<td><input type="radio" name="sectionShow" value="NA" checked onClick="disableCheckBox(ratesOnly)" class="clearBorder"></td>
										<td>Don't show a message.</td>
									</tr>
									<tr> 
										<td>
											<% if PC_SECTIONSHOW="TOP" then %>
												<input type="radio" name="sectionShow" value="TOP" checked onClick="disableCheckBox(ratesOnly)" class="clearBorder"> 
											<%else %>
												<input type="radio" name="sectionShow" value="TOP" onClick="disableCheckBox(ratesOnly)" class="clearBorder"> 
											<% end if %>
										</td>
										<td>Display a message at the top of the page, before shipping options are shown.</td>
									</tr>
									<tr> 
										<td>
											<% if PC_SECTIONSHOW="BTM" then %>
												<input type="radio" name="sectionShow" value="BTM" checked onClick="enableCheckBox(ratesOnly)" class="clearBorder"> 
											<%else %>
												<input type="radio" name="sectionShow" value="BTM" onClick="enableCheckBox(ratesOnly)" class="clearBorder"> 
											<% end if %>
										</td>
										<td>Display a message at the bottom of the page, after shipping options are shown.</td>
									</tr>
									<tr> 
										<td>&nbsp;</td>
										<td><input type="checkbox" name="ratesOnly" value="YES" <% if PC_RATESONLY="YES" then%>checked<% end if %><%if PC_SECTIONSHOW<>"BTM" then %> disabled<% end if %> class="clearBorder">&nbsp;Only show the message if shipping rates are returned. Note that if you choose to display a message at the top of the page, the message will be displayed even if no shipping rates are returned</td>
									</tr>
									<tr> 
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									</tr>
								</table>
						</td>
						</tr>
						<tr> 
							<td width="10%">Title:</td>
							<td width="90%"><input type="text" name="ShipDetailTitle" value="<%=PC_SHIP_DETAIL_TITLE%>" size="60"></td>
						</tr>
						<tr valign="top"> 
							<td>Message:</td>
							<%
							pcv_strShipDetails = replace(PC_SHIP_DETAILS,"<br>",vbcrlf)
							pcv_strShipDetails = replace(pcv_strShipDetails,"<BR>",vbcrlf)
							%>
							<td><textarea name="ShipDetails" cols="60" rows="8"><%=pcv_strShipDetails%></textarea></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
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
		<input type="submit" name="Submit" value="Update" class="submit2">
		&nbsp;
		<input type="button" value="View/Edit Shipping Options" class="ibtnGrey" onClick="location.href='viewShippingOptions.asp'">
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->