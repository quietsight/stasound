<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="USPS Shipping Configuration - Default Package Settings" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% if request.form("submit")<>"" then
	EMPackage=request.form("EMPackage")
	Session("ship_USPS_EM_PACKAGE")=EMPackage
	PMPackage=request.form("PMPackage")
	Session("ship_USPS_PM_PACKAGE")=PMPackage
	USPS_HEIGHT=request.form("USPS_HEIGHT")
	Session("ship_USPS_HEIGHT")=USPS_HEIGHT
	USPS_WIDTH=request.form("USPS_WIDTH")
	Session("ship_USPS_WIDTH")=USPS_WIDTH
	USPS_LENGTH=request.form("USPS_LENGTH")
	Session("ship_USPS_LENGTH")=USPS_LENGTH
	If (Int(USPS_LENGTH)+((int(USPS_WIDTH)*2)+(int(USPS_HEIGHT)*2)))>108 then
		response.redirect "2_Step3.asp?msg="&server.URLEncode("You default package measurements excede USPS standards of 108 inches.")
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
	response.redirect "2_Step5.asp"
	response.end
else %>
	<% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>

	<form name="form1" method="post" action="2_Step3.asp" class="pcForms">
	<table class="pcCPcontent">
          <tr> 
            <td colspan="2"><h2>Default Package Types:&nbsp;</h2></td>
          </tr>
          <tr> 
            <td colspan="2"><b>Express Mail</b></td>
          </tr>
          <tr> 
            <td align="right"><input name="EMPackage" type="radio" value="None" checked> </td>
            <td>Your Packaging</td>
          </tr>
          <tr> 
            <td align="right"><input type="radio" name="EMPackage" value="Flat Rate Envelope"></td>
            <td>Express Mail Flat Rate Envelope, 12.5&quot; x 9.5&quot;</td>
          </tr>
            <tr> 
              <td>&nbsp;</td>
              <td>
              <table class="pcCPcontent" style="background-color: #e1e1e1;">
                <tr>
                  <td colspan="2">
                    Weight limit for Flat Rate Envelope:
                    <% if USPS_EM_FREWeightLimit="" then
                        USPS_EM_FREWeightLimit="0"
                    end if %>
                    <input name="EM_FREWeightLimit" type="text" id="textfield" size="4" value="0"></td>
                </tr>
                <tr>
                  <td width="5%" valign="top"><div align="right">
                  <input type="checkbox" name="EM_FREOption" value="1" class="clearBorder">
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
            <td colspan="2"><b>Priority Mail</b></td>
          </tr>
          <tr> 
            <td align="right"><input name="PMPackage" type="radio" value="None" checked> </td>
            <td>Your Packaging</td>
          </tr>
            <tr>
              <td align="right"><input type="radio" name="PMPackage" value="Flat Rate Envelope"></td>
              <td>Priority Mail Flat Rate Envelope, 12.5&rdquo; x 9.5&rdquo;</td>
            </tr>
            <tr> 
              <td>&nbsp;</td>
              <td>
              <table class="pcCPcontent" style="background-color: #e1e1e1;">
                <tr>
                  <td colspan="2">Weight limit for Flat Rate Envelope: 
                        <% if USPS_PM_FREWeightLimit="" then
                            USPS_PM_FREWeightLimit="0"
                        end if %>
                        <input name="PM_FREWeightLimit" type="text" id="textfield2" size="4" value="0">
                  </td>
                </tr>
                <tr>
                    <td colspan="2">When Flat Rate Envelope weight limit is exceeded...</td>
                </tr>
                  <tr>
                    <td width="5%"><div align="right"><input type="radio" name="PM_FREOption" value="0" checked></div></td>
                    <td>Do not show Priority Mail</td>
                  </tr>
                  <tr>
                    <td width="5%"><div align="right"><input type="radio" name="PM_FREOption" value="NONE"></div></td>
                    <td width="95%">Use &quot;Your Packaging&quot;</td>
                  </tr>
                  <tr>
                    <td><div align="right"><input type="radio" name="PM_FREOption" value="Flat Rate Box"></div></td>
                    <td>Priority Mail Small Flat Rate Box, 8-5/8" x 5-3/8" x 1-5/8"</td>
                  </tr>
                  <tr>
                    <td><div align="right"><input type="radio" name="PM_FREOption" value="Flat Rate Box1"></div></td>
                    <td>Priority Mail Medium   Flat Rate Box, 11" x 8-1/2" x 5-1/2"<br>
&nbsp;&nbsp;OR <br>
Priority Mail Medium   Flat Rate Box, 13-5/8" x 11-7/8" x 3-3/8"</td>
                  </tr>
                  <tr>
                    <td><div align="right"><input type="radio" name="PM_FREOption" value="Flat Rate Box2"></div></td>
                    <td>Priority Mail Large Flat Rate Box , 12" x 12" x 5-1/2"</td>
                  </tr>
              </table></td>
            </tr>
            <tr> 
          <tr> 
            <td align="right"><input type="radio" name="PMPackage" value="Flat Rate Box" <% if USPS_PM_PACKAGE="Flat Rate Box" then%>checked<%end if%>>
            </td>
            <td>Priority Mail Small Flat Rate Box, 8-5/8" x 5-3/8" x 1-5/8"</td>
          </tr>
          <tr> 
            <td align="right"><input type="radio" name="PMPackage" value="Flat Rate Box1" <% if USPS_PM_PACKAGE="Flat Rate Envelope" then%>checked<%end if%>></td>
            <td>Priority Mail Medium   Flat Rate Box, 11" x 8-1/2" x 5-1/2"<br>
&nbsp;&nbsp;OR <br>
Priority Mail Medium   Flat Rate Box, 13-5/8" x 11-7/8" x 3-3/8"</td>
          </tr>
          <tr> 
            <td align="right"><input type="radio" name="PMPackage" value="Flat Rate Box2" <% if USPS_PM_PACKAGE="Flat Rate Envelope" then%>checked<%end if%>></td>
            <td>Priority Mail Large Flat Rate Box , 12" x 12" x 5-1/2"</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"> 
            <h2>Default Package Size</h2>
            If you are using your own packaging to ship products, please specify a default package size below. This size should be the most common size for the majority of packages that you ship.<br>
            Packages are measured as package length plus girth: (length + ((width x 2) + (height x 2)))
            <br>
            Your default size should measure no more then 108 inches (length plus girth)</td>
          </tr>
          <tr> 
            <td>Height: </td>
            <td>
              <input name="USPS_HEIGHT" type="text" id="USPS_HEIGHT" value="10" size="4" maxlength="4"> 
            </td>
          </tr>
          <tr> 
            <td>Width:</td>
            <td> 
              <input name="USPS_WIDTH" type="text" id="USPS_WIDTH" value="10" size="4" maxlength="4"> 
            </td>
          </tr>
          <tr> 
            <td>Length:</td>
            <td> 
              <input name="USPS_LENGTH" type="text" id="USPS_LENGTH" value="10" size="4" maxlength="4"> 
              <font color="#FF0000"> This is the measurement of the longest side</font></td>
          </tr>
          <tr> 
			<td colspan="2"><hr></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> <input type="submit" name="Submit" value="Submit" class="submit2"></td>
          </tr>
        </table>
	</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->