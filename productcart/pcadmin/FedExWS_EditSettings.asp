<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Web Services Shipping Configuration - Edit Settings" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% dim query, rs, conntemp
call opendb()
query="SELECT ShipmentTypes.AccessLicense FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=9));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
strAccessLicense=rs("AccessLicense")
if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if
set rs=nothing
call closedb()

pcv_strFEDEXWS_DROPOFF_TYPE = FEDEXWS_DROPOFF_TYPE

if pcv_strFEDEXWS_DROPOFF_TYPE = "" then
	pcv_strFEDEXWS_DROPOFF_TYPE = "REGULARPICKUP"
end if

pcv_FEDEXWS_LISTRATE = FEDEXWS_LISTRATE
if pcv_FEDEXWS_LISTRATE = "" then
	pcv_FEDEXWS_LISTRATE = "0"
end if

pcv_FEDEXWS_SATURDAYDELIVERY = FEDEXWS_SATURDAYDELIVERY
if pcv_FEDEXWS_SATURDAYDELIVERY = "" then
	pcv_FEDEXWS_SATURDAYDELIVERY = "0"
end if

pcv_FEDEXWS_SATURDAYPICKUP = FEDEXWS_SATURDAYPICKUP
if pcv_FEDEXWS_SATURDAYPICKUP = "" then
	pcv_FEDEXWS_SATURDAYPICKUP = "0"
end if

if request.form("submit")<>"" then
	Session("ship_FEDEXWS_FEDEX_PACKAGE")=request.form("FEDEXWS_PACKAGE")
	Session("ship_FEDEXWS_DROPOFF_TYPE")=request.form("FEDEXWS_DROPOFF_TYPE")
	Session("ship_FEDEXWS_HEIGHT")=request.form("FEDEXWS_HEIGHT")
	Session("ship_FEDEXWS_WIDTH")=request.form("FEDEXWS_WIDTH")
	Session("ship_FEDEXWS_LENGTH")=request.form("FEDEXWS_LENGTH")
	Session("ship_FEDEXWS_DIM_UNIT")=request.form("FEDEXWS_DIM_UNIT")
	Session("ship_FEDEXWS_ADDDAY")=request.form("FEDEXWS_ADDDAY")
	Session("ship_FEDEXWS_LISTRATE")=request.form("FEDEXWS_LISTRATE")
	Session("ship_FEDEXWS_SATURDAYDELIVERY")=request.form("FEDEXWS_SATURDAYDELIVERY")
	Session("ship_FEDEXWS_SATURDAYPICKUP")=request.form("FEDEXWS_SATURDAYPICKUP")
	Session("ship_FEDEXWS_DYNAMICINSUREDVALUE")=request.form("DynamicInsuredValue")
	Session("ship_FEDEXWS_INSUREDVALUE")=request.form("InsuredValue")
	Session("ship_FEDEXWS_SMHUBID")=request.form("SMHubID")
	response.redirect "../includes/PageCreateFedExWSConstants.asp?refer=viewShippingOptions.asp#FedExWS"
	response.end
else %>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

	<form name="form1" method="post" action="FedExWS_EditSettings.asp" class="pcForms">
		<table class="pcCPcontent">
				<tr>
				  <td colspan="2">
					<h2>Default Package Type</h2>
					Typically orders are shipped in different boxes depending on what customers purchased. Therefore, in most cases, you will select <em>Custom Packaging</em> here, and  specify the most common box size under <em>Default Package Size</em>. <a href="http://wiki.earlyimpact.com/productcart/shipping-federal_express_ws#packaging_type" target="_blank">See the documentation for details</a>. </td>
				</tr>
				<tr>
					<td align="right">
						<input name="FEDEXWS_PACKAGE" type="radio" value="YOUR_PACKAGING" checked>
					</td>
					<td>Custom Packaging</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_10KG_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_10KG_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; 10kg Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_25KG_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_25KG_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; 25kg Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_BOX" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_BOX" then%>checked<%end if%>>								</td>
					<td>FedEx Box</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_ENVELOPE" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_ENVELOPE" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Envelope</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_PAK" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_PAK" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Pak</td>
		  </tr>
				<tr>
					<td align="right">
					<input type="radio" name="FEDEXWS_PACKAGE" value="FEDEX_TUBE" <%if FEDEXWS_FEDEX_PACKAGE="FEDEX_TUBE" then%>checked<%end if%>>								</td>
					<td>FedEx&reg; Tube</td>
		  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
					<h2>Default Package Size</h2>
					If you selected <em>Custom Packaging</em> under <em>Default Package Type</em>, enter the most common package size below. This should refer to the size of the box used for the majority of your shipments. <a href="http://wiki.earlyimpact.com/productcart/shipping-federal_express_ws#packaging_type" target="_blank">See the documentation for details.</a></td>
				</tr>
				<tr>
					<td align="right">Height: </td>
				  <td><input name="FEDEXWS_HEIGHT" type="text" id="FEDEXWS_HEIGHT" value="<%=FEDEXWS_HEIGHT%>" size="4" maxlength="4"></td>
		  </tr>
				<tr>
					<td align="right">Width: </td>
				  <td><input name="FEDEXWS_WIDTH" type="text" id="FEDEXWS_WIDTH" value="<%=FEDEXWS_WIDTH%>" size="4" maxlength="4"></td>
		  </tr>
				<tr>
					<td align="right">Length:</td>
					<td> <input name="FEDEXWS_LENGTH" type="text" id="FEDEXWS_LENGTH" value="<%=FEDEXWS_LENGTH%>" size="4" maxlength="4">
				  <span class="pcSmallText">This is the measurement of the longest side</span></td>
		  </tr>
				<tr>
					<td align="right">Measurement Unit:</td>
					<td>
					<% if FEDEXWS_DIM_UNIT="CM" then%>
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="IN">
					Inches
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="CM" checked>
					Centimeters
					<% else %>
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="IN" checked>
					Inches
					<input type="radio" name="FEDEXWS_DIM_UNIT" value="CM">
					Centimeters
					<% end if %>
				  </td>
		  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>FedEx Insurance Settings</h2>
					</td>
				</tr>
				<tr>
					<td colspan="2">To use the order total (total of products listed in the shopping cart) as the amount to insure, choose <em>Use Dynamic Insurance Rate</em>. If you are not using <em>Dynamic Insurance Rate</em>, you can set a <em>Flat Rate</em> as the amount to insure, and that amount will be used for every FedEx rate calculation in the storefront. The default flat rate will be set to $100.00 if one is not set and you have not selected to use <em>Dynamic Insurance Rate</em>. </td>
				</tr>
				<tr>
					<td align="right"><input name="DynamicInsuredValue" type="radio" value="1" <%if FDXWS_DYNAMICINSUREDVALUE="1" then%>checked<% end if %>></td>
					<td>Use Dynamic Insurance Rate</td>
				</tr>
				<tr>
					<td align="right"><input type="radio" name="DynamicInsuredValue" value="0" <%if FDXWS_DYNAMICINSUREDVALUE="0" then%>checked<% end if %>></td>
					<td>Use Flat Rate</td>
				</tr>
				<tr>
					<td width="14%" nowrap>Flat  Rate Value:</td>
					<td width="86%"><input type="text" name="InsuredValue" value="<%=FDXWS_INSUREDVALUE%>" size="15" maxlength="14" /></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					 <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>FedEx SmartPost</h2>
					</td>
				</tr>
				<tr>
					<td colspan="2">To use SmartPost you must have SmartPost enabled on your account. When you have SmartPost enabled you will be supplied a HUB ID that should be entered here. If you are not using SmartPost you can leave this field empty.<br>
					  <br>
					<font color="#FF0000">*To enable SmartPost for your FedEx account, please contact your FedEx Account Representative.</font></td>
				</tr>
				<tr>
					<td width="14%"><p>Hub ID: </p></td>
					<td width="86%"><p>
					<select name="SMHubID">
						<option value="" selected>Select HUB ID</option>
						<option value="5015" <%if FDXWS_SMHUBID="5015" then%>selected<%end if%>>5015</option>Northborough, MA</option>
						<option value="5087" <%if FDXWS_SMHUBID="5087" then%>selected<%end if%>>5087</option>Edison, NJ</option>
						<option value="5150" <%if FDXWS_SMHUBID="5150" then%>selected<%end if%>>5150</option>Pittsburgh, PA</option>
						<option value="5185" <%if FDXWS_SMHUBID="5185" then%>selected<%end if%>>5185</option>Allentown, PA</option>
						<option value="5254" <%if FDXWS_SMHUBID="5254" then%>selected<%end if%>>5254</option>Martinsburg, WV</option>
						<option value="5281" <%if FDXWS_SMHUBID="5281" then%>selected<%end if%>>5281</option>Charlotte, NC</option>
						<option value="5303" <%if FDXWS_SMHUBID="5303" then%>selected<%end if%>>5303</option>Atlanta, GA</option>
						<option value="5327" <%if FDXWS_SMHUBID="5327" then%>selected<%end if%>>5327</option>Orlando, FL</option>
						<option value="5379" <%if FDXWS_SMHUBID="5379" then%>selected<%end if%>>5379</option>Memphis, TN</option>
						<option value="5431" <%if FDXWS_SMHUBID="5431" then%>selected<%end if%>>5431</option>Grove City, OH</option>
						<option value="5465" <%if FDXWS_SMHUBID="5465" then%>selected<%end if%>>5465</option>Indianapolis, IN</option>
						<option value="5481" <%if FDXWS_SMHUBID="5481" then%>selected<%end if%>>5481</option>Detroit, MI</option>
						<option value="5531" <%if FDXWS_SMHUBID="5531" then%>selected<%end if%>>5531</option>New Berlin, WI</option>
						<option value="5552" <%if FDXWS_SMHUBID="5552" then%>selected<%end if%>>5552</option>Minneapolis, MN</option>
						<option value="5631" <%if FDXWS_SMHUBID="5631" then%>selected<%end if%>>5631</option>St. Louis, MO</option>
						<option value="5648" <%if FDXWS_SMHUBID="5648" then%>selected<%end if%>>5648</option>Kansas, KS</option>
						<option value="5751" <%if FDXWS_SMHUBID="5751" then%>selected<%end if%>>5751</option>Dallas, TX</option>
						<option value="5771" <%if FDXWS_SMHUBID="5771" then%>selected<%end if%>>5771</option>Houston, TX</option>
						<option value="5802" <%if FDXWS_SMHUBID="5802" then%>selected<%end if%>>5802</option>Denver, CO</option>
						<option value="5843" <%if FDXWS_SMHUBID="5843" then%>selected<%end if%>>5843</option>Salt Lake City, UT</option>
						<option value="5854" <%if FDXWS_SMHUBID="5854" then%>selected<%end if%>>5854</option>Phoenix, AZ</option>
						<option value="5902" <%if FDXWS_SMHUBID="5902" then%>selected<%end if%>>5902</option>Los Angeles, CA</option>
						<option value="5929" <%if FDXWS_SMHUBID="5929" then%>selected<%end if%>>5929</option>Chino, CA</option>
						<option value="5958" <%if FDXWS_SMHUBID="5958" then%>selected<%end if%>>5958</option>Sacramento, CA</option>
						<option value="5983" <%if FDXWS_SMHUBID="5983" then%>selected<%end if%>>5983</option>Seattle, WA</option>
					</select>
					</p></td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2"><h2>Default Pickup Method</h2></td>
				</tr>
				<tr>
					<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="REGULAR_PICKUP" <%if pcv_strFEDEXWS_DROPOFF_TYPE="REGULAR_PICKUP" then%>checked<%end if%>>                </td>
					<td>Regular Pick-up </td>
				</tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="REQUEST_COURIER" <%if pcv_strFEDEXWS_DROPOFF_TYPE="REQUEST_COURIER" then%>checked<%end if%>>                </td>
				  <td>Request Courier  </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="DROP_BOX" <%if pcv_strFEDEXWS_DROPOFF_TYPE="DROP_BOX" then%>checked<%end if%>>                </td>
				  <td>Dropbox </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="BUSINESS_SERVICE_CENTER" <%if pcv_strFEDEXWS_DROPOFF_TYPE="BUSINESS_SERVICE_CENTER" then%>checked<%end if%>>                </td>
				  <td>Business Service Center  </td>
		  </tr>
				<tr>
	<td align="right"><input name="FEDEXWS_DROPOFF_TYPE" type="radio" value="STATION" <%if pcv_strFEDEXWS_DROPOFF_TYPE="STATION" then%>checked<%end if%>>                </td>
				  <td>Station </td>
		  </tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
					<h2>Rate Type</h2>
					Select which shipping rates you want to be shown in the storefront. NOTE: List Rates and Discounted Rates will not always be different amounts and such behavior does not indicate this feature isn't working.
					<a href="http://wiki.earlyimpact.com/productcart/shipping-federal_express_ws#rate_type" target="_blank">See the documentation for details.</a></td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="-1" <%if pcv_FEDEXWS_LISTRATE="-1" then%>checked<%end if%>>
					</td>
					<td>Show List Rates</td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="0" <%if pcv_FEDEXWS_LISTRATE="0" then%>checked<%end if%>>
					</td>
					<td>Show Discounted Rates<br><em>Do not select this option unless your account is eligible for Discount Rates</em></td>
				</tr>
				<tr>
					<td align="right" valign="top">
					<input name="FEDEXWS_LISTRATE" type="radio" value="-2" <%if pcv_FEDEXWS_LISTRATE="-2" then%>checked<%end if%>>
					</td>
					<td>Show Discounted Rates + Multiweight Ground Rates<br><em>Do not select this option unless your account is eligible for Multiweight Ground Rates</em></td>
				</tr>

				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>Saturday Delivery</h2>
						<strong>Note:</strong> When this feature is "On" Saturday Delivery options are displayed only when available.
					</td>
				</tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYDELIVERY" type="radio" value="0" <%if pcv_FEDEXWS_SATURDAYDELIVERY="0" then%>checked<%end if%>>
					</td>
					<td>Off (recommended)</td>
				  </tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYDELIVERY" type="radio" value="-1" <%if pcv_FEDEXWS_SATURDAYDELIVERY="-1" then%>checked<%end if%>>
					</td>
					<td>On</td>
				  </tr>
				<tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2">
						<h2>Saturday Pickup</h2>
						<strong>Note:</strong> When this feature is "On" Saturday Pickup pricing is displayed  when available.
					If you never will ship on a Satuday, you will turn this feature off.</td>
				</tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYPICKUP" type="radio" value="0" <%if pcv_FEDEXWS_SATURDAYPICKUP="0" then%>checked<%end if%>>
					</td>
					<td>Off (recommended)</td>
				  </tr>
				<tr>
					<td align="right">
					<input name="FEDEXWS_SATURDAYPICKUP" type="radio" value="-1" <%if pcv_FEDEXWS_SATURDAYPICKUP="-1" then%>checked<%end if%>>
					</td>
					<td>On</td>
				  </tr>
				<tr>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td>
				</tr>
				<tr>
				<td colspan="2">
				  <h2>Delay Shipment Setting</h2>
						<strong>Note:</strong> The system will add a lead time to the expected shipping date for the rate requests in the store front. Set this to the number of days after in which you  normally ship out packages  after the date they are ordered. This will give proper delivery dates for &quot;Next Day&quot; and other expedited shipment methods.<br>
						<br>
						<br>
						Example: If you set this to 2 days and an order was placed on Monday, the ship date will be set for that Wednesday. Therefore the customer will see the service for &quot;FedEx Next Day&quot; with a expected Delivery day of Thursday instead of Tuesday.</td>
				</tr>
				<tr>
					<td align="right">Number of Days:
			  </td>
					<td><input name="FEDEXWS_ADDDAY" type="text" id="FEDEXWS_ADDDAY" value="<%=FEDEXWS_ADDDAY%>" size="2" maxlength="2">
				  </td>
		  </tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
					<hr>
					<input type="submit" name="Submit" value="Submit" class="submit2"></td>
				</tr>
	  </table>
</form>
<% end if %>

<!--#include file="AdminFooter.asp"-->