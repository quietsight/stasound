<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS - Edit Shipping Services" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/USPSconstants.asp"-->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td>
			<% if request.form("submit")<>"" then
				Session("ship_USPS_EM_PACKAGE")=request.form("EMPackage")
				Session("ship_USPS_PM_PACKAGE")=request.form("PMPackage")
				Session("ship_USPS_HEIGHT")=request.form("USPS_HEIGHT")
				Session("ship_USPS_WIDTH")=request.form("USPS_WIDTH")
				Session("ship_USPS_LENGTH")=request.form("USPS_LENGTH")

				If (Int(Session("ship_USPS_LENGTH"))+((int(Session("ship_USPS_WIDTH"))*2)+(int(Session("ship_USPS_HEIGHT"))*2)))>108 then
					response.redirect "USPS_EditSettings.asp?msg="&server.URLEncode("You default package measurements excede USPS standards of 108 inches.")
				end if
				Session("ship_USPS_EM_FREWeightLimit")=request.form("EM_FREWeightLimit")
				if NOT isNumeric(Session("ship_USPS_EM_FREWeightLimit")) OR Session("ship_USPS_EM_FREWeightLimit")="" then
					Session("ship_USPS_EM_FREWeightLimit")="0"
				end if
				Session("ship_USPS_EM_FREOption")=request.form("EM_FREOption")
				If Session("ship_USPS_EM_FREOption")="" then
					Session("ship_USPS_EM_FREOption")="0"
				End If
				Session("ship_USPS_PM_FREWeightLimit")=request.form("PM_FREWeightLimit")
				if NOT isNumeric(Session("ship_USPS_PM_FREWeightLimit")) OR Session("ship_USPS_PM_FREWeightLimit")="" then
					Session("ship_USPS_PM_FREWeightLimit")="0"
				end if
				Session("ship_USPS_PM_FREOption")=request.form("PM_FREOption")
				If Session("ship_USPS_PM_FREOption")="" then
					Session("ship_USPS_PM_FREOption")="0"
				End If
				response.redirect "../includes/PageCreateUSPSConstants.asp?refer=viewShippingOptions.asp#USPS"
				response.end
			else %>
            
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

				<form name="form1" method="post" action="USPS_EditSettings.asp">
                    <table class="pcCPcontent">
                        <tr>
                            <td class="pcCPspacer" colspan="2"></td>
                        </tr>
                        <tr> 
                          <th colspan="2">Express Mail Default Package Type</th>
                        </tr>
                        <tr>
                            <td class="pcCPspacer" colspan="2"></td>
                        </tr>
                        <tr> 
                          <td> 
                            <input name="EMPackage" type="radio" value="NONE" checked>
                            </td>
                          <td>Your Packaging</td>
                        </tr>
                        <tr> 
                          <td> 
                            <input type="radio" name="EMPackage" value="Flat Rate Envelope" <% if USPS_EM_PACKAGE="Flat Rate Envelope" then%>checked<%end if%>>
                            </td>
                          <td>Express Mail Flat Rate Envelope, 12.5&quot; x 9.5&quot;</td>
                        </tr>
                        <tr> 
                          <td>&nbsp;</td>
                          <td bgcolor="#e1e1e1">
                          <table width="100%" border="0" cellspacing="1" cellpadding="1">
                            <tr>
                              <td width="32%">
                                  <div>Weight limit for Flat Rate Envelope:</div></td>
								<% if USPS_EM_FREWeightLimit="" then
									USPS_EM_FREWeightLimit="0"
                                end if %>
                              <td width="68%"><input name="EM_FREWeightLimit" type="text" id="textfield" size="4" value="<%=USPS_EM_FREWeightLimit%>"></td>
                            </tr>

                        </table>
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">

                            <tr>
                                <td width="5%" valign="top"><div align="right">
                              <input type="checkbox" name="EM_FREOption" value="1" <% if USPS_EM_FREOption="1" then%>checked<%end if%>>
                                </div></td>
                                <td width="95%"><p>Use &quot;Your Packaging&quot; when Flat Rate Envelope weight limit is exceeded?</p>
                                <p><strong>Note:</strong> If you do not choose to use &quot;Your Packaging&quot; when a weight limit is set, Express Mail will not be shown as an option when the weight limit is exceeded.</p></td>
                            </tr>
                        </table>
                        </td>
                    </tr>
                    <tr>
                        <td class="pcCPspacer" align="center" colspan="2"></td>
                    </tr>
                    <tr> 
                        <th colspan="2">Priority Mail Default Package Type</th>
                	</tr>
                    <tr>
                        <td class="pcCPspacer" align="center" colspan="2"></td>
                    </tr>

                    <tr> 
                      <td> 
                        <input name="PMPackage" type="radio" value="NONE" checked>
                        </td>
                      <td>Your Packaging</td>
                    </tr>
                    <tr>
                      <td>
                        <input type="radio" name="PMPackage" value="Flat Rate Envelope" <% if USPS_PM_PACKAGE="Flat Rate Envelope" then%>checked<%end if%>>
                      </td>
                      <td>Priority Mail Flat Rate Envelope, 12.5&rdquo; x 9.5&rdquo;</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td bgcolor="#e1e1e1"><table width="100%" border="0" cellspacing="1" cellpadding="1">
                        <tr>
                          <td width="32%"><div>Weight limit for Flat Rate Envelope:</div></td>
								<% if USPS_PM_FREWeightLimit="" then
									USPS_PM_FREWeightLimit="0"
                                end if %>
                          <td width="68%"><input name="PM_FREWeightLimit" type="text" id="textfield2" size="4" value="<%=USPS_PM_FREWeightLimit%>"></td>
                        </tr>
    
                      </table>
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">
    
                          <tr>
                            <td width="5%"><div align="right">
                              <input type="radio" name="PM_FREOption" value="0" <% if USPS_PM_FREOption="0" then%>checked<%end if%>>
                            </div></td>
                            <td>Do not show Priority Mail as an option if weight limit is exceeded.</td>
                          </tr>
                          <tr>
                            <td width="5%"><div align="right">
                              <input type="radio" name="PM_FREOption" value="NONE" <% if USPS_PM_FREOption="NONE" then%>checked<%end if%>>
                            </div></td>
                            <td width="95%">Use &quot;Your Packaging&quot; when Flat Rate Envelope weight limit is exceeded?</td>
                          </tr>
                          <tr>
                            <td width="5%"><div align="right">
                              <input type="radio" name="PM_FREOption" value="Flat Rate Box" <% if USPS_PM_FREOption="Flat Rate Box" then%>checked<%end if%>>
                            </div></td>
                            <td>Priority Mail Small Flat Rate Box, 8-5/8" x 5-3/8" x 1-5/8"</td>
                          </tr>
                          <tr>
                            <td width="5%"><div align="right">
                              <input type="radio" name="PM_FREOption" value="Flat Rate Box1" <% if USPS_PM_FREOption="Flat Rate Box1" then%>checked<%end if%>>
                            </div></td>
                            <td>                            Priority Mail Medium   Flat Rate Box, 11" x 8-1/2" x 5-1/2"<br>
                              &nbsp;&nbsp;OR
                              <br>
                            Priority Mail Medium   Flat Rate Box, 13-5/8" x 11-7/8" x 3-3/8"</td>
                          </tr>
                          <tr>
                            <td width="5%"><div align="right">
                              <input type="radio" name="PM_FREOption" value="Flat Rate Box2" <% if USPS_PM_FREOption="Flat Rate Box2" then%>checked<%end if%>>
                            </div></td>
                            <td>Priority Mail Large Flat Rate Box , 12" x 12" x 5-1/2"</td>
                          </tr>
                      </table></td>
                    </tr>
                    <tr> 
                      <td> 
                        <input type="radio" name="PMPackage" value="Flat Rate Box" <% if USPS_PM_PACKAGE="Flat Rate Box" then%>checked<%end if%>>
                        </td>
                      <td>Priority Mail Small Flat Rate Box, 8-5/8" x 5-3/8" x 1-5/8"</td>
                    </tr>
                    <tr> 
                      <td>
                        <input type="radio" name="PMPackage" value="Flat Rate Box1" <% if USPS_PM_PACKAGE="Flat Rate Box1" then%>checked<%end if%>>
                        </td>
                      <td>Priority Mail Medium   Flat Rate Box, 11" x 8-1/2" x 5-1/2"<br>
&nbsp;&nbsp;OR <br>
Priority Mail Medium   Flat Rate Box, 13-5/8" x 11-7/8" x 3-3/8"</td>
                    </tr>
                    <tr> 
                      <td> 
                        <input type="radio" name="PMPackage" value="Flat Rate Box2" <% if USPS_PM_PACKAGE="Flat Rate Box2" then%>checked<%end if%>>
                        </td>
                      <td>                      Priority Mail Large Flat Rate Box , 12" x 12" x 5-1/2"</td>
                    </tr>
                    <tr>
                      <td class="pcCPspacer" align="center" colspan="2"></td>
                    </tr>
                    <tr>
                      <th colspan="2">Default Package Size</th>
                    </tr>
                    <tr>
                      <td class="pcCPspacer" align="center" colspan="2"></td>
                    </tr>
                    <tr> 
                      <td colspan="2"><p>Choose a default package size. This size will be the common size of the majority of packages that you ship. If an item is larger, such as an oversized product, you can specify that on a per product basis in the product details.<br>
                          <br>
                          Packages are measured as package length plus girth: (length + ((width x 2) + (height x 2)))</p>
                        <p>Your default size should measure no more then 108 inches (length plus girth)</p></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>Height: 
                        <input name="USPS_HEIGHT" type="text" id="USPS_HEIGHT" value="<%=USPS_HEIGHT%>" size="4" maxlength="4">
                        </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>Width: 
                        <input name="USPS_WIDTH" type="text" id="USPS_WIDTH" value="<%=USPS_WIDTH%>" size="4" maxlength="4">
                        </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>Length: 
                        <input name="USPS_LENGTH" type="text" id="USPS_LENGTH" value="<%=USPS_LENGTH%>" size="4" maxlength="4">
                        <span class="pcSmallText">* This is the measurement of the longest side</span></td>
                    </tr>
                    <tr>
                      <td class="pcCPspacer" align="center" colspan="2"><hr></td>
                    </tr>
                    <tr> 
                      <td colspan="2" align="center"><input type="submit" name="Submit" value="Submit" class="submit2"></td>
                    </tr>
                  </table>
                  </form>
			<% end if %>
        </td>
    </tr>
</table>
<!--#include file="AdminFooter.asp"-->