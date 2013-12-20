<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx<sup>&reg;</sup> Shipping Configuration - Edit Settings" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<% dim query, rs, conntemp
call opendb()
query="SELECT ShipmentTypes.AccessLicense FROM ShipmentTypes WHERE (((ShipmentTypes.idShipment)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
strAccessLicense=rs("AccessLicense")
if len(strAccessLicense)<1 then
	strAccessLicense="TEST"
end if 
set rs=nothing
call closedb() 

pcv_strFEDEX_DROPOFF_TYPE = FEDEX_DROPOFF_TYPE

if pcv_strFEDEX_DROPOFF_TYPE = "" then
	pcv_strFEDEX_DROPOFF_TYPE = "REGULARPICKUP"
end if

pcv_FEDEX_LISTRATE = FEDEX_LISTRATE
if pcv_FEDEX_LISTRATE = "" then
	pcv_FEDEX_LISTRATE = "0"
end if
%>
<table width="94%" border="0" align="center" cellpadding="1" cellspacing="0">
	<tr>
		<td>
			<% if request.form("submit")<>"" then
				Session("ship_FEDEX_FEDEX_PACKAGE")=request.form("FEDEX_PACKAGE")
				Session("ship_FEDEX_DROPOFF_TYPE")=request.form("FEDEX_DROPOFF_TYPE")
				Session("ship_FEDEX_HEIGHT")=request.form("FEDEX_HEIGHT")
				Session("ship_FEDEX_WIDTH")=request.form("FEDEX_WIDTH")
				Session("ship_FEDEX_LENGTH")=request.form("FEDEX_LENGTH")
				Session("ship_FEDEX_DIM_UNIT")=request.form("FEDEX_DIM_UNIT")
				Session("ship_FEDEX_LISTRATE")=request.form("FEDEX_LISTRATE")
				Session("ship_FEDEX_DYNAMICINSUREDVALUE")=request.form("DynamicInsuredValue")
				Session("ship_FEDEX_INSUREDVALUE")=request.form("InsuredValue")
				response.redirect "../includes/PageCreateFedExConstants.asp?refer=viewShippingOptions.asp#FedEx"
				response.end
			else %>
				<form name="form1" method="post" action="FedEx_EditSettings.asp" class="pcForms">
					<table class="pcCPcontent">
						<% if request.querystring("msg")<>"" then %>
							<tr> 
								<td colspan="2"> 
									<table width="100%" border="0" cellspacing="0" cellpadding="4">
										<tr> 
											<td width="4%"><img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"></td>
											<td width="96%"><font color="#FF9900"><b><%=request.querystring("msg")%></b></font></td>
										</tr>
									</table></td>
							</tr>
						<% end if %>
							<tr> 
								  <td colspan="2"><strong>FedEx API Status</strong></td>
							</tr>
							<tr>
								 <td colspan="2">
									<% 
									if strAccessLicense="TEST" then
										response.write "You are currently running in <b>&quot;TEST&quot;</b> mode. <a href='ConfigureOption3.asp?changeMode=Y'>Click here to switch to Production Mode</a>."
									else
									
									end if 
									%>
								</td>
							  </tr>
							<tr>
							  <td colspan="2">&nbsp;</td>
							  </tr>
							<tr> 
								  <td colspan="2"><h2>Default Package Type</h2></td>
							</tr>
							<tr> 
								<td> 
									<input name="FEDEX_PACKAGE" type="radio" value="YOURPACKAGING" checked>								</td>
								<td>Customer Package </td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEX10KGBOX" <%if FEDEX_FEDEX_PACKAGE="FEDEX10KGBOX" then%>checked<%end if%>>								</td>
								<td>FedEx 10kg Box</td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEX25KGBOX" <%if FEDEX_FEDEX_PACKAGE="FEDEX25KGBOX" then%>checked<%end if%>>								</td>
								<td>FedEx 25kg Box</td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEXBOX" <%if FEDEX_FEDEX_PACKAGE="FEDEXBOX" then%>checked<%end if%>>								</td>
								<td>FedEx&reg; Box</td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEXENVELOPE" <%if FEDEX_FEDEX_PACKAGE="FEDEXENVELOPE" then%>checked<%end if%>>								</td>
								<td>FedEx&reg; Envelope</td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEXPAK" <%if FEDEX_FEDEX_PACKAGE="FEDEXPAK" then%>checked<%end if%>>								</td>
								<td>FedEx&reg; Pak</td>
							</tr>
							<tr> 
								<td> 
								<input type="radio" name="FEDEX_PACKAGE" value="FEDEXTUBE" <%if FEDEX_FEDEX_PACKAGE="FEDEXTUBE" then%>checked<%end if%>>								</td>
								<td>FedEx&reg; Tube</td>
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
                                <td colspan="2">To use the value of the cart as the insurance rate value, choose to <span style="font-weight: bold; font-style: italic">Use Dynamic Insurance Rate</span>. If you are not using Dynamic Insurance Rate, you can set a <span style="font-weight: bold; font-style: italic">Flat Rate</span> that will be used for every FedEx rate calculation in the store front. The default flat rate will be set to $100.00 if one is not set and you have not selected to use dynamic insurance rates. </td>
                            </tr>
                            <tr>
                                <td align="right"><p><input name="DynamicInsuredValue" type="radio" value="1" <%if FDX_DYNAMICINSUREDVALUE="1" then%>checked<% end if %>></p></td>
                                <td>Use Dynamic Insurance Rate</td>
                            </tr>
                            <tr>
                                <td align="right"><p><input type="radio" name="DynamicInsuredValue" value="0" <%if FDX_DYNAMICINSUREDVALUE="0" then%>checked<% end if %>></p></td>
                                <td>Use Flat Rate</td>
                            </tr>
                            <tr>
                                <td width="14%"><p>Flat  Rate Value: </p></td>
                                <td width="86%"><p><input type="text" name="InsuredValue" value="<%=FDX_INSUREDVALUE%>" size="15" maxlength="14" /> </p></td>
                            </tr>
							<tr>
							<tr>
							  <td>&nbsp;</td>
							  <td>&nbsp;</td>
							</tr>
                            <tr>
                            	<td colspan="2"><h2>Default Pickup Method</h2></td>
                            </tr>
							<tr>
                				<td><input name="FEDEX_DROPOFF_TYPE" type="radio" value="REGULARPICKUP" <%if pcv_strFEDEX_DROPOFF_TYPE="REGULARPICKUP" then%>checked<%end if%>>                </td>
							  	<td>Regular Pick-up </td>
							</tr>
							<tr>
                <td><input name="FEDEX_DROPOFF_TYPE" type="radio" value="REQUESTCOURIER" <%if pcv_strFEDEX_DROPOFF_TYPE="REQUESTCOURIER" then%>checked<%end if%>>                </td>
							  <td>Request Courier  </td>
							  </tr>
							<tr>
                <td><input name="FEDEX_DROPOFF_TYPE" type="radio" value="DROPBOX" <%if pcv_strFEDEX_DROPOFF_TYPE="DROPBOX" then%>checked<%end if%>>                </td>
							  <td>Dropbox </td>
							  </tr>
							<tr>
                <td><input name="FEDEX_DROPOFF_TYPE" type="radio" value="BUSINESSSERVICECENTER" <%if pcv_strFEDEX_DROPOFF_TYPE="BUSINESSSERVICECENTER" then%>checked<%end if%>>                </td>
							  <td>Business Service Center  </td>
							  </tr>
							<tr>
                <td><input name="FEDEX_DROPOFF_TYPE" type="radio" value="STATION" <%if pcv_strFEDEX_DROPOFF_TYPE="STATION" then%>checked<%end if%>>                </td>
							  <td>Station </td>
							  </tr>
							<tr> 
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
                				<td colspan="2">
				<h2>Default Rate Type (displayed on storefront)</h2><br>
                </strong>Note: List Rates and Discounted Rates will not always be different amounts and such behavior does not indicate this feature isn't working. 
				</td>
					  </tr>
							<tr>
                				<td>
								<input name="FEDEX_LISTRATE" type="radio" value="0" <%if pcv_FEDEX_LISTRATE="0" then%>checked<%end if%>>                				
								</td>
							  	<td>Show Discounted Rates</td>
							  </tr>							
							<tr>
                				<td>
								<input name="FEDEX_LISTRATE" type="radio" value="-1" <%if pcv_FEDEX_LISTRATE="-1" then%>checked<%end if%>>
                				</td>
							  	<td>Show List Rates</td>
							  </tr>
							<tr>
							  <td>&nbsp;</td>
							  <td>&nbsp;</td>
							</tr>
							<tr> 
								<td colspan="2">                                
								<h2>Default Package Size</h2><br>
								If you are using your own packaging to ship products, please choose a default
								package size below. This size should be the most common size for the majority
								of packages that you ship.<br>
								<br>
								Packages are measured as package length plus girth:<br>
								(length + ((width x 2) + (height x 2)))</p>
								<ul>
								<li>A package weighing less than 30 lbs. and measuring more than 84 inches and equal
								to or less than 108 inches in combined length and girth will be classified as
								an Oversize 1 (OS1) package. The transportation charges for an Oversize 1 (OS1) package
								will be the same as a 30-lb. package being transported under the same circumstances.<br>
								<br>
								</li>
								<li>A package weighing less than 50 lbs. and measuring more than 108 inches in combined
								length and girth will be classified as an Oversize 2 (OS2) package. The transportation
								charges for an Oversize 2 (OS2) package will be the same as a 50-lb. package being
								transported under the same circumstances. </li>
								</ul></td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>Height: 
								<input name="FEDEX_HEIGHT" type="text" id="FEDEX_HEIGHT" value="<%=FEDEX_HEIGHT%>" size="4" maxlength="4">								</td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>Width: 
								<input name="FEDEX_WIDTH" type="text" id="FEDEX_WIDTH" value="<%=FEDEX_WIDTH%>" size="4" maxlength="4">								</td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>Length: 
								<input name="FEDEX_LENGTH" type="text" id="FEDEX_LENGTH" value="<%=FEDEX_LENGTH%>" size="4" maxlength="4">
								<font color="#FF0000"> *This is the measurement of the longest side</font></td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>Measurement Unit:
								<% if FEDEX_DIM_UNIT="CM" then%>
								<input type="radio" name="FEDEX_DIM_UNIT" value="IN">
								Inches 
								<input type="radio" name="FEDEX_DIM_UNIT" value="CM" checked>
								Centimeters
								<% else %>
								<input type="radio" name="FEDEX_DIM_UNIT" value="IN" checked>
								Inches 
								<input type="radio" name="FEDEX_DIM_UNIT" value="CM">
								Centimeters
								<% end if %>								</td>
							</tr>
							<tr> 
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
							<tr> 
								<td colspan="2" align="center">
								<input type="submit" name="Submit" value="Submit" class="ibtnGrey"></td>
							</tr>
						</table>
					</form>
				<% end if %>
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->