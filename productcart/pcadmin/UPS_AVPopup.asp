<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/languages_ship.asp" --> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/UPSconstants.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<SCRIPT LANGUAGE="JavaScript"><!--
function copyForm() {
    opener.document.form1.ShipToCity.value = document.popupForm.ShipToCity.value;
    opener.document.form1.ShipToStateOrProvinceCode.value = document.popupForm.ShipToStateCode.value;
    opener.document.form1.ShipToPostalCode.value = document.popupForm.ShipToPostalCode.value;
    //opener.document.form1.submit();
    window.close();
    return false;
}
//--></SCRIPT>
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style1 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style></head>
<body>

<% If request("Submit")<>"" then
	strNewAddress=getUserInput(request("NewAddress"),0)
	strNewAddressArry=split(strNewAddress,"||")
	ShipToCity=strNewAddressArry(0)
	ShipToState=strNewAddressArry(1)
	ShipToPC=strNewAddressArry(2)
	%>
	<table>
		<tr>
			<td bgcolor="#E1E1E1"><span class="style1">UPS OnLine&reg; Tools Address Validation </span></td>
		</tr>
		<tr>
			<td><p><strong><br />
			  You chose</strong><br />
			  City: <%=ShipToCity%><br />
			  State: <%=ShipToState%><br />
			  Postal Code: <%=ShipToPC%></p>
		  <p>&nbsp;</p></td>
		</tr>
		<tr><td>
		<FORM NAME="popupForm" onSubmit="return copyForm()">
			<p>
				<input type="hidden" name="ShipToCity" value="<%=ShipToCity%>" />
				<input type="hidden" name="ShipToStateCode" value="<%=ShipToState%>" />
				<input type="hidden" name="ShipToPostalCode" value="<%=ShipToPC%>" />
				If this is the address that you wish to use, click on the &quot;Submit&quot; button below and the appropriate fields will be populated with the new address information. Remember to &quot;Apply&quot; the changes on the main page, so that the new address is saved to the database.	  </p>
			<p>  
			<INPUT TYPE="BUTTON" VALUE="Submit" onClick="copyForm()">
			</p>
		</FORM>
		</td></tr>
		<tr>
		  <td>&nbsp;</td>
	  </tr>
		<tr>
		  <td><table width="100%" bgcolor="#FFFFFF">
        <tr>
          <td width="45" valign="middle"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></td>
          <td width="654" valign="middle"><span class="pcSmallText">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</span></td>
        </tr>
      </table></td>
	  </tr>
		</table>
<% else
	
	PassedShipToState=request("State")
	PassedShipToCity=request("City")
	PassedShipToPC=request("PC")
	PassedShipToCountry=request("Country") 
	
	intCityMatch=0
	intStateMatch=0
	intPCMatch=0
	intPerfectMatch=0
	strOptionString=""
	
	dim query, rs, conntemp
	
	'//OPEN DB CONNECTION
	call openDb()
	
	'//UPS Variables from database
	query="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	
	ups_active=rs("active")
	ups_userid=trim(rs("userID"))
	ups_password=trim(rs("password"))
	ups_license_key=trim(rs("AccessLicense"))
	
	set rs=nothing
	call closedb()
	
	ups_postdata=""
	ups_postdata="<?xml version=""1.0""?>"&vbcrlf
	ups_postdata=ups_postdata&"<AccessRequest xml:lang=""en-US"">"&vbcrlf
	ups_postdata=ups_postdata&"<AccessLicenseNumber>"&ups_license_key&"</AccessLicenseNumber>"&vbcrlf
	ups_postdata=ups_postdata&"<UserId>"&ups_userid&"</UserId>"&vbcrlf
	ups_postdata=ups_postdata&"<Password>"&ups_password&"</Password>"&vbcrlf
	ups_postdata=ups_postdata&"</AccessRequest>"&vbcrlf
	ups_postdata=ups_postdata&"<?xml version=""1.0""?>"&vbcrlf
	ups_postdata=ups_postdata&"<AddressValidationRequest xml:lang=""en-US"">"&vbcrlf
	ups_postdata=ups_postdata&"<Request>"&vbcrlf
	ups_postdata=ups_postdata&"<TransactionReference>"&vbcrlf
	ups_postdata=ups_postdata&"<CustomerContext>Maryam Dennis-Customer Data</CustomerContext>"&vbcrlf
	ups_postdata=ups_postdata&"<XpciVersion>1.0001</XpciVersion>"&vbcrlf
	ups_postdata=ups_postdata&"</TransactionReference>"&vbcrlf
	ups_postdata=ups_postdata&"<RequestAction>AV</RequestAction>"&vbcrlf
	ups_postdata=ups_postdata&"</Request>"&vbcrlf
	ups_postdata=ups_postdata&"<Address>"&vbcrlf
	ups_postdata=ups_postdata&"<City>"&PassedShipToCity&"</City>"&vbcrlf
	ups_postdata=ups_postdata&"<StateProvinceCode>"&PassedShipToState&"</StateProvinceCode>"&vbcrlf
	ups_postdata=ups_postdata&"</Address>"&vbcrlf
	ups_postdata=ups_postdata&"</AddressValidationRequest>"&vbcrlf
	
	ups_URL="https://www.ups.com/ups.app/xml/AV"
	
	Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
	srvUPSXmlHttp.open "POST", ups_URL, false
	srvUPSXmlHttp.send(ups_postdata)
	UPS_result = srvUPSXmlHttp.responseText
	
	'FOR DEBUGGING - SEE RESPONSE.BACK FROM UPS.
	'If the XML parser is working, then uncomment "on error resume next" at top of file.
	'uncomment the following two lines also
	'response.write UPS_result
	'response.end
	
	Set USPSINTXMLdoc = server.CreateObject("Msxml2.DOMDocument"&scXML)
	USPSINTXMLDoc.async = false 
	if USPSINTXMLDoc.loadXML(UPS_result) then ' if loading from a string
		set objLst = USPSINTXMLDoc.getElementsByTagName("AddressValidationResult")
		for i = 0 to (objLst.length-1)
			intNoMatchCnt=0
			for j=0 to ((objLst.item(i).childNodes.length)-1)
			If objLst.item(i).childNodes(j).nodeName="Rank" then
				ShipToRank=objLst.item(i).childNodes(j).text
			End if
			If objLst.item(i).childNodes(j).nodeName="Address" then
				for m=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
					If objLst.item(i).childNodes(j).childNodes(m).nodeName="City" then
						ShipToCity=objLst.item(i).childNodes(j).childNodes(m).text
					end if
					If objLst.item(i).childNodes(j).childNodes(m).nodeName="StateProvinceCode" then
						ShipToState=objLst.item(i).childNodes(j).childNodes(m).text
					end if
				next
			End if
			If objLst.item(i).childNodes(j).nodeName="StateProvinceCode" then
				ShipToState=objLst.item(i).childNodes(j).text
			End if
			If objLst.item(i).childNodes(j).nodeName="PostalCodeLowEnd" then
				ShipToPC=objLst.item(i).childNodes(j).text
			End if
			If objLst.item(i).childNodes(j).nodeName="PostalCodeHighEnd" then
				ShipToPCHighEnd=objLst.item(i).childNodes(j).text
			End if
	
			next
			'compare data to the submitted data
			if lcase(ShipToCity)=lcase(PassedShipToCity) then
				intCityMatch=1
			else
				intNoMatchCnt=intNoMatchCnt+1
			end if
			
			if lcase(ShipToState)=lcase(PassedShipToState) then
				intStateMatch=1
			else
				intNoMatchCnt=intNoMatchCnt+1
			end if
			
			if ShipToPCHighEnd<>ShipToPC then
				'loop through range
				ShipToPCCnt=int(ShipToPC)
				do until ShipToPCCnt=int(ShipToPCHighEnd)
					if int(ShipToPCCnt)=int(PassedShipToPC) then
						intPCMatch=1
					end if
					ShipToPCCnt=int(ShipToPCCnt)+1
				loop
				if intPCMatch=1 then
				else
					intNoMatchCnt=intNoMatchCnt+1
				end if
			else
				if lcase(ShipToPC)=lcase(PassedShipToPC) then
					intPCMatch=1
				else
					intNoMatchCnt=intNoMatchCnt+1
				end if
			end if
			
			if intNoMatchCnt=0 then 
				intPerfectMatch=1
			end if 
			
			'create dropdown string to show if needed later 
			strOptionString=strOptionString& "<option value='"&ShipToCity&"||"&ShipToState&"||"&ShipToPC&"'>"&ShipToCity&", "&ShipToState&" "&ShipToPC&"</option>"
	
		next
	end if 
	
	if intPerfectMatch=1 then %>
		<table>
      <tr>
        <td bgcolor="#E1E1E1"><span class="style1">UPS OnLine&reg; Tools Address Validation </span></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><img src="images/pcadmin_successful.gif" width="18" height="18" /> <strong> This address has passed UPS address validation. </strong></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr bgcolor="#FFFF99">
        <td><strong>NOTICE:</strong> UPS assumes no liability for the information 
          provided by the address validation functionality.<br /> 
          The 
          address validation functionality does not support the 
        identification or verification of occupants at an address.</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>
					<table width="100%" bgcolor="#FFFFFF">
						<tr>
							<td width="45" valign="middle"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></td>
							<td width="654" valign="middle"><span class="pcSmallText">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</span></td>
						</tr>
        	</table>        </td>
      </tr>
</table>
	<% else %>
		<form id="form1" name="form1" method="post" action="UPS_AVPopup.asp">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td colspan="2" bgcolor="#E1E1E1"><span class="style1">UPS OnLine&reg; Tools Address Validation </span></td>
			  </tr>
				<tr>
				  <td colspan="2">&nbsp;</td>
			  </tr>
				<tr>
					<td colspan="2"><p><strong><img src="images/note.gif" width="20" height="20" /> An address error was found</strong><br />
				  <br />
					Behind the scenes, the UPS Address Validation Tool checked this shipping address. There seems to be a problem.</p>
					<p>Address Submitted: <%=PassedShipToCity&", "&PassedShipToState&" "&PassedShipToPC%> </p></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
					<% if len(strOptionString)<1 then %>
					No suggestions were found for this address
					<% else %>
					Suggested Value: 
						<select name="NewAddress">
							<%=strOptionString %>
						</select>
					<% end if %></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr bgcolor="#FFFF99">
					<td colspan="2"><strong>NOTICE:</strong> The address validation functionality will validate 
						P.O. Box addresses, however, UPS does not deliver to P.O.
						boxes, attempts by customer to ship to a P.O. Box via UPS 
					may result in additional charges.</td>
				</tr>
				<tr>
					<td width="101">&nbsp;</td>
					<td width="625">&nbsp;</td>
				</tr>
				<% if len(strOptionString)>0 then %>
					<tr>
						<td><input type="submit" name="Submit" value="Select Address" /></td>
						<td>&nbsp;</td>
					</tr>
				<% end if %>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">NOTICE: UPS assumes no liability for the information 
						provided by the address validation functionality. The 
						address validation functionality does not support the 
						identification or verification of occupants at an address.</td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><table width="100%" bgcolor="#FFFFFF">
            <tr>
              <td width="45" valign="middle"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50" /></td>
              <td width="654" valign="middle"><span class="pcSmallText">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</span></td>
            </tr>
          </table></td>
				</tr>
			</table>
		</form>
	<% end if
end if %>
</body>
</html>
