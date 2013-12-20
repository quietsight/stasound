<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.buffer=true %>
<!--#include file="CustLIv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/bto_language.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/UPSconstants.asp" -->
<!--#include file="../includes/rc4.asp" --> 

<!--#include file="header.asp"-->
<style>
.resultstable{
    color:<%=FColor%>;
    font-family:<%=FFType%>;
    font-size:12px;
	}
TD{	font-family:<%=FFType%>;
	font-size: 12px;
	color:<%=FColor%>;
	}
</style>
<% dim mySQL, conntemp, rstemp
call openDb()

'//UPS Variables
mySQL="SELECT active, userID, [password], AccessLicense FROM ShipmentTypes WHERE idshipment=3"
set rstemp=conntemp.execute(mySQL)
ups_license_key=trim(rstemp("AccessLicense"))
ups_userid=trim(rstemp("userID"))
ups_password=trim(rstemp("password"))
ups_active=rstemp("active")

call closedb()
if request.form("SubmitTracking")<>"" then
	dim pitracknumber
	pitracknumber=request("itracknumber")
	pitracknumber=replace(pitracknumber," ","")
	session("itracknumber")=uCase(pitracknumber)
	if request.form("iagree")="" then
		response.redirect "CustUPSTracking.asp?msg="&server.URLEncode("You must agree to the UPS Tracking Terms and Conditions before continuing.")
	end if
	'//UPS Rates
	ups_trackdata=""
	ups_trackdata="<?xml version=""1.0""?>"
	ups_trackdata=ups_trackdata&"<AccessRequest xml:lang=""en-US"">"
	ups_trackdata=ups_trackdata&"<AccessLicenseNumber>"&ups_license_key&"</AccessLicenseNumber>"
	ups_trackdata=ups_trackdata&"<UserId>"&ups_userid&"</UserId>"
	ups_trackdata=ups_trackdata&"<Password>"&ups_password&"</Password>"
	ups_trackdata=ups_trackdata&"</AccessRequest>"
	ups_trackdata=ups_trackdata&"<?xml version=""1.0""?>"
	ups_trackdata=ups_trackdata&"<TrackRequest xml:lang=""en-US"">"
	ups_trackdata=ups_trackdata&"<Request>"
	ups_trackdata=ups_trackdata&"<TransactionReference>"
	ups_trackdata=ups_trackdata&"<CustomerContext>Example 1</CustomerContext>"
	ups_trackdata=ups_trackdata&"<XpciVersion>1.0001</XpciVersion>"
	ups_trackdata=ups_trackdata&"</TransactionReference>"
	ups_trackdata=ups_trackdata&"<RequestAction>Track</RequestAction>"
	ups_trackdata=ups_trackdata&"<RequestOption>activity</RequestOption>"
	ups_trackdata=ups_trackdata&"</Request>"
	ups_trackdata=ups_trackdata&"<TrackingNumber>"&session("itracknumber")&"</TrackingNumber>"
	ups_trackdata=ups_trackdata&"</TrackRequest>"
	'get URL to post to
	ups_URL="https://www.ups.com/ups.app/xml/Track"
	
	toResolve = 3000
	toConnect = 3000
	toSend = 3000
	toReceive = 3000
	
	Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
	'srvXmlHttp.setTimeouts toResolve, toConnect, toSend, toReceive ' not needed but a handy feature
	srvXmlHttp.open "POST", ups_URL, false
	srvXmlHttp.send(ups_trackdata)
	result = srvXmlHttp.responseText
	'response.write result&"<BR>"
	'response.end
	Set XMLdoc = server.CreateObject("Msxml2.DOMDocument")
	XMLDoc.async = false 
	if xmldoc.loadXML(result) then ' if loading from a string
		set objLst = xmldoc.getElementsByTagName("ResponseStatusCode") 
		for i = 0 to (objLst.length - 1)
			varStatus=objLst.item(i).text
			if varStatus="0" then
				set objLst = xmldoc.getElementsByTagName("Error") 
				for j = 0 to (objLst.length - 1)
					for k=0 to ((objLst.item(j).childNodes.length)-1)
						If objLst.item(j).childNodes(k).nodeName="ErrorDescription" then
							varErrorDescription=objLst.item(j).childNodes(k).text
						end if
					next
				next
			end if
		next
			
		if varStatus="0" then %>
			<table width="100%" border="0" cellspacing="0" cellpadding="4">
				<tr> 
					<td>
						<table width="100%" border="0" cellspacing="0" cellpadding="2">
							<tr> 
								<td width="9%"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
								<td width="91%"><b><font face="<%=FFType%>"  color="<%=FColor%>" size="3">UPS 
								Tracking</font></b></td>
							</tr>
						</table></td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td bgcolor="<%=AColor%>"><font face="<%=FFType%>"  color="<%=AFColor%>" size="2">Error 
						retrieving tracking information</font></td>
				</tr>
				<tr> 
					<td><font face="<%=FFType%>"  color="<%=AFColor%>" size="2"> 
						<%response.write varErrorDescription%>
						</font></td>
				</tr>
				<tr> 
					<td>&nbsp;</td>
				</tr>
				<tr> 
					<td><font face="<%=FFType%>"  color="<%=FColor%>" size="2"> 
						<input class=ibtn type="button" name="Button" value="Back" onClick="javascript:history.back()">
						</font></td>
				</tr>
				<tr>
					<td colspan=4>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan=4><p><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>UPS Tracking Terms &amp; Conditions:</b></font></p>
						<P><font face="<%=FFType%>"  color="<%=FColor%>" size="1">
						NOTICE: The UPS package tracking systems accessed via this
						service (the "Tracking Systems") and tracking information obtained
						through this service (the "Information") are the private property of
						UPS. UPS authorizes you to use the Tracking Systems solely to track
						shipments tendered by or for you to UPS for delivery and for no
						other purpose. Without limitation, you are not authorized to make
						the Information available on any web site or otherwise reproduce,
						distribute, copy, store, use or sell the Information for commercial
						gain without the express written consent of UPS. This is a personal
						service, thus your right to use the Tracking Systems or Information
						is non-assignable. Any access or use that is inconsistent with these
						terms is unauthorized and strictly prohibited.</font>
						</P></td>
				</tr>
				<tr> 
					<td colspan=4>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan=4><font face="<%=FFType%>"  color="<%=FColor%>" size="2">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF <br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</font> </td>
				</tr>
				<tr> 
					<td colspan=4>&nbsp;</td>
				</tr>
			</table>
		<% else
			%>
			<table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr> 
                      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                          <tr> 
                            <td width="9%"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
                            <td width="91%"><b><font face="<%=FFType%>"  color="<%=FColor%>" size="3">UPS 
                              Tracking</font></b></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td colspan="4">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="4"><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>Tracking 
                        Summary</b></font></td>
                    </tr>
										<% set objLst = xmldoc.getElementsByTagName("ShipTo")
										for i = 0 to (objLst.length - 1)
										for j=0 to ((objLst.item(i).childNodes.length)-1)
											If objLst.item(i).childNodes(j).nodeName="Address" then
												for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
													if objLst.item(i).childNodes(j).childNodes(k).nodeName="AddressLine1" then
														upstAddressLine1= objLst.item(i).childNodes(j).childNodes(k).text
													end if
													if objLst.item(i).childNodes(j).childNodes(k).nodeName="City" then
														upstCity= objLst.item(i).childNodes(j).childNodes(k).text
													end if
													if objLst.item(i).childNodes(j).childNodes(k).nodeName="StateProvinceCode" then
														upstStateProvinceCode= objLst.item(i).childNodes(j).childNodes(k).text
													end if
													if objLst.item(i).childNodes(j).childNodes(k).nodeName="PostalCode" then
														upstPostalCode= objLst.item(i).childNodes(j).childNodes(k).text
													end if
													if objLst.item(i).childNodes(j).childNodes(k).nodeName="CountryCode" then
														upstCountryCode= objLst.item(i).childNodes(j).childNodes(k).text
													end if
												next
											End if
										next
									 next
									 %>
									 <% set objLst = xmldoc.getElementsByTagName("Service")
										for i = 0 to (objLst.length - 1)
										for j=0 to ((objLst.item(i).childNodes.length)-1)
											If objLst.item(i).childNodes(j).nodeName="Code" then
												upstServiceCode= objLst.item(i).childNodes(j).text
											End if
											If objLst.item(i).childNodes(j).nodeName="Description" then
												upstServiceDescription= objLst.item(i).childNodes(j).text
											End if
										next
									 next
									 select case upstServiceCode
										case "01"
											upstService="UPS Next Day Air <sup>&reg;</sup>"
										case "02"
											upstService="UPS 2nd Day Air <sup>&reg;</sup>"
										case "03"
											upstService="UPS Ground"
										case "07"
											upstService="UPS Worldwide Express <sup>SM</sup>"
										case "08"
											upstService="UPS Worldwide Expedited <sup>SM</sup>"
										case "11"
											upstService="UPS Standard To Canada"
										case "12"
											upstService="UPS 3 Day Select <sup>&reg;</sup>"
										case "13"
											upstService="UPS Next Day Air Saver <sup>&reg;</sup>"
										case "14"
											upstService="UPS Next Day Air<sup>&reg;</sup> Early A.M. <sup>&reg;</sup>"
										case "54"
											upstService="UPS Worldwide Express Plus <sup>SM</sup>"
										case "59"
											upstService="UPS 2nd Day Air A.M. <sup>&reg;</sup>"
										case "65"
											upstService="UPS Express Saver <sup>SM</sup>"
										case else
											upstService=upstServiceDescription
										end select
									 %>
									 <% set objLst = xmldoc.getElementsByTagName("Package")
										for i = 0 to (objLst.length - 1)
										for j=0 to ((objLst.item(i).childNodes.length)-1)
											If objLst.item(i).childNodes(j).nodeName="TrackingNumber" then
												upstNumber= objLst.item(i).childNodes(j).text
											End if
										next
									 next
									 %>
                    <tr> 
                      <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td width="95%"><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>Tracking Number:</b>&nbsp;<%=upstNumber%></font></td>
                            <td width="5%">&nbsp;</td>
                          </tr>
													<% if upstAddressLine1<>"" then %>
                          <tr> 
                            <td><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>Shipped To:</b>&nbsp;<%=upstAddressLine1%>&nbsp;<%=upstCity%>,&nbsp;<%=upstStateProvinceCode%>&nbsp;<%=upstPostalCode%>&nbsp;<%=upstCountryCode%></font></td>
                            <td>&nbsp;</td>
                          </tr>
													<% end if %>
                          <tr>
                            <td><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>Service:</b>&nbsp;<%=upstService%></font></td>
                            <td>&nbsp;</td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="19%" bgcolor="<%=AColor%>"><font face="<%=FFType%>"  color="<%=AFColor%>" size="2"><b>Date</b></font></td>
                      <td width="24%" bgcolor="<%=AColor%>"><font face="<%=FFType%>"  color="<%=AFColor%>" size="2"><b>Time</b></font></td>
                      <td width="33%" bgcolor="<%=AColor%>"><font face="<%=FFType%>"  color="<%=AFColor%>" size="2"><b>Location</b></font></td>
                      <td width="24%" bgcolor="<%=AColor%>"><font face="<%=FFType%>"  color="<%=AFColor%>" size="2"><b>Activity</b></font></td>
                    </tr>
                    <%
				set objLst = xmldoc.getElementsByTagName("Activity") 
				for i = 0 to (objLst.length - 1)
					for j=0 to ((objLst.item(i).childNodes.length)-1)
						If objLst.item(i).childNodes(j).nodeName="ActivityLocation" then
							for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="Address" then
									UPSLOCATION= objLst.item(i).childNodes(j).childNodes(k).text
								end if
							next
						End if
						If objLst.item(i).childNodes(j).nodeName="Status" then
							for k=0 to ((objLst.item(i).childNodes(j).childNodes.length)-1)
								if objLst.item(i).childNodes(j).childNodes(k).nodeName="StatusType" then
									for l=0 to ((objLst.item(i).childNodes(j).childNodes(k).childNodes.length)-1)
										if objLst.item(i).childNodes(j).childNodes(k).childNodes(l).nodeName="Description" then
											UPSACTIVITY= objLst.item(i).childNodes(j).childNodes(k).childNodes(l).text
										end if
									next
								end if
							next
						End if
						If objLst.item(i).childNodes(j).nodeName="Date" then
							UPSDATE= objLst.item(i).childNodes(j).text
							UPSYEAR=left(UPSDATE,4)
							UPSMONTH=Mid(UPSDATE, 5, 2)
							UPSDAY=right(UPSDATE, 2) 
						End If
						If objLst.item(i).childNodes(j).nodeName="Time" then
							UPSTIME= objLst.item(i).childNodes(j).text
							UPSHOUR=left(UPSTIME,2)
							UPSMINUTE=Mid(UPSTIME, 3, 2)
							if CINT(UPSHOUR)>11 then
							 UPSAMPM="PM"
							else
							 UPSAMPM="AM"
							End If
							SELECT CASE UPSHOUR
								CASE "00"
									UPSHOUR="12"
								CASE "01"
									UPSHOUR="1"
								CASE "02"
									UPSHOUR="2"
								CASE "03"
									UPSHOUR="3"
								CASE "04"
									UPSHOUR="4"
								CASE "05"
									UPSHOUR="5"
								CASE "06"
									UPSHOUR="6"
								CASE "07"
									UPSHOUR="7"
								CASE "08"
									UPSHOUR="8"
								CASE "09"
									UPSHOUR="9"
								CASE "10"
									UPSHOUR="10"
								CASE "11"
									UPSHOUR="11"
								CASE "12"
									UPSHOUR="12"
								CASE "13"
									UPSHOUR="1"
								CASE "14"
									UPSHOUR="2"
								CASE "15"
									UPSHOUR="3"
								CASE "16"
									UPSHOUR="4"
								CASE "17"
									UPSHOUR="5"
								CASE "18"
									UPSHOUR="6"
								CASE "19"
									UPSHOUR="7"
								CASE "20"
									UPSHOUR="8"
								CASE "21"
									UPSHOUR="9"
								CASE "22"
									UPSHOUR="10"
								CASE "23"
									UPSHOUR="11"
								END SELECT
						End If
					next %>
                    <tr valign="top"> 
                      <td CLASS="resultstable"><%=UPSMONTH&"/"&UPSDAY&"/"&UPSYEAR%></td>
                      <td CLASS="resultstable"><%=UPSHOUR&":"&UPSMINUTE&" "&UPSAMPM%></td>
                      <td CLASS="resultstable"><%=UPSLOCATION%></td>
                      <td CLASS="resultstable"><%=UPSACTIVITY%></td>
                    </tr>
                    <% 	next %>
                    <tr> 
                      <td colspan=4>&nbsp;</td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan=4><font face="<%=FFType%>"  color="<%=FColor%>" size="2"> 
                        <input class=ibtn type="button" name="Button" value="Back" onClick="javascript:history.back()">
                        </font></td>
                    </tr>
                    <tr> 
                      <td colspan=4>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan=4><p><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>UPS 
                          Tracking Terms &amp; Conditions:</b></font></p>
                        <P><font face="<%=FFType%>"  color="<%=FColor%>" size="1"> 
                          NOTICE: The UPS package tracking systems accessed via 
                          this service (the "Tracking Systems") and tracking information 
                          obtained through this service (the "Information") are 
                          the private property of UPS. UPS authorizes you to use 
                          the Tracking Systems solely to track shipments tendered 
                          by or for you to UPS for delivery and for no other purpose. 
                          Without limitation, you are not authorized to make the 
                          Information available on any web site or otherwise reproduce, 
                          distribute, copy, store, use or sell the Information 
                          for commercial gain without the express written consent 
                          of UPS. This is a personal service, thus your right 
                          to use the Tracking Systems or Information is non-assignable. 
                          Any access or use that is inconsistent with these terms 
                          is unauthorized and strictly prohibited.</font> </P></td>
                    </tr>
                    <tr> 
                      <td colspan=4>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan=4><font face="<%=FFType%>"  color="<%=FColor%>" size="2">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF <br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</font> </td>
                    </tr>
                    <tr> 
                      <td colspan=4>&nbsp;</td>
                    </tr>
                  </table>	
		<% end if %>
	<% end if
else 'form is submitted
	itracknumber=request.querystring("itracknumber") 
	session("itracknumber")=itracknumber %>
	<form name="trackform" method="post" action="custUPSTracking.asp">
		<table width="100%" border="0" cellspacing="0" cellpadding="2">
			<tr> 
				<td width="12%"><img src="../UPSLicense/LOGO_S2.jpg" width="45" height="50"></td>
				<td width="88%"><b><font face="<%=FFType%>"  color="<%=FColor%>" size="3">UPS 
					Tracking</font></b></td>
			</tr>
			<tr> 
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<% if request.QueryString("msg")<>"" then %>
			<tr> 
				<td colspan="2"><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><%=request.QueryString("msg")%></font></td>
			</tr>
			<tr> 
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<% end if %>
		</table>
		<font face="<%=FFType%>"  color="<%=FColor%>" size="2">UPS 
			Tracking Number:</font> 
			<input name="itracknumber" type="text" id="itracknumber" size="30" value="<%=session("itracknumber")%>">
		<p>
		<table border="0" width="500" cellspacing="0" cellpadding="0">
			<tr>
				<td> 
					<p><font face="<%=FFType%>"  color="<%=FColor%>" size="2"><b>UPS Tracking Terms &amp; Conditions:</b></font></p><P><font face="<%=FFType%>"  color="<%=FColor%>" size="1">
					NOTICE: The UPS package tracking systems accessed via this
					service (the "Tracking Systems") and tracking information obtained
					through this service (the "Information") are the private property of
					UPS. UPS authorizes you to use the Tracking Systems solely to track
					shipments tendered by or for you to UPS for delivery and for no
					other purpose. Without limitation, you are not authorized to make
					the Information available on any web site or otherwise reproduce,
					distribute, copy, store, use or sell the Information for commercial
					gain without the express written consent of UPS. This is a personal
					service, thus your right to use the Tracking Systems or Information
					is non-assignable. Any access or use that is inconsistent with these
					terms is unauthorized and strictly prohibited.</font>
					</P></td>
			</tr>
		</table>
		<p> 
			<input type=checkbox name="iagree" value="1">
			<font face="<%=FFType%>"  color="<%=FColor%>" size="2">I agree to the Terms &amp; Conditions Set Forth Above</font></p>
		<p><br>
			<input class=ibtn type="button" name="Button" value="Back" onClick="javascript:history.back()">
			&nbsp; 
			<input type="submit" name="SubmitTracking" value="Submit">
			&nbsp; 
			<input type="reset" name="Reset" value="Reset">
		</p>
		</p><font face="<%=FFType%>"  color="<%=FColor%>" size="2">UPS, THE UPS SHIELD TRADEMARK, THE UPS READY MARK, <br />THE UPS ONLINE TOOLS MARK AND THE COLOR BROWN ARE TRADEMARKS OF <br />UNITED PARCEL SERVICE OF AMERICA, INC. ALL RIGHTS RESERVED.</font> 
		<p></p>
		<p>&nbsp; </p>
  </form>
<% end if %> 
<!--#include file="footer.asp"-->