<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="FedEx Shipping Configuration" %>
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
			FedExPackage=request.form("FedExPackage")
			Session("ship_FEDEX_FEDEX_PACKAGE")=FedExPackage
			FEDEX_HEIGHT=request.form("FEDEX_HEIGHT")
			Session("ship_FEDEX_HEIGHT")=FEDEX_HEIGHT
			FEDEX_WIDTH=request.form("FEDEX_WIDTH")
			Session("ship_FEDEX_WIDTH")=FEDEX_WIDTH
			FEDEX_LENGTH=request.form("FEDEX_LENGTH")
			Session("ship_FEDEX_LENGTH")=FEDEX_LENGTH
			FEDEX_DIM_UNIT=request.form("FEDEX_DIM_UNIT")
			Session("ship_FEDEX_DIM_UNIT")=FEDEX_DIM_UNIT

			response.redirect "3_Step5.asp"
			response.end
		else %>
<form name="form1" method="post" action="3_Step3.asp" class="pcForms">
        <table class="pcCPcontent">
          <% if request.querystring("msg")<>"" then %>
          <tr> 
            <td colspan="2">
				<% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
          </tr>
          <% end if %>
          <tr> 
            <th colspan="2">Default Package Types</th>
          </tr>
          <tr> 
            <td colspan="2" class="pcCPspacer"></td>
          </tr>
          <tr> 
            <td align="right"> <input name="FedExPackage" type="radio" value="1" checked class="clearBorder"> 
            </td>
            <td>Your Packaging</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="5" class="clearBorder"> 
            </td>
            <td>FedEx 10k Box</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="25" class="clearBorder"> 
            </td>
            <td>FedEx 25k Box</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="3" class="clearBorder"> 
            </td>
            <td>FedEx Box</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="6" class="clearBorder"> 
            </td>
            <td>FedEx Envelope</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="2" class="clearBorder"> 
            </td>
            <td>FedEx Pak</td>
          </tr>
          <tr> 
            <td align="right"> <input type="radio" name="FedExPackage" value="4" class="clearBorder"> 
            </td>
            <td>FedEx Tube</td>
          </tr>
          <tr> 
            <td colspan="2" class="pcCPspacer"></td>
          </tr>
          <tr> 
            <th colspan="2">Default Package Size</th>
          </tr>
          <tr> 
            <td colspan="2" class="pcCPspacer"></td>
          </tr>
          <tr> 
            <td colspan="2"> <p> If you are using your own packaging to ship products, 
                please specify a default package size below. This size should 
                be the most common size for the majority of packages that you 
                ship.<br>
                <br>
                Packages are measured as package length plus girth:<br>
                (length + ((width x 2) + (height x 2)))</p>
              <ul>
                <li>A package weighing less than 30 lbs. and measuring more than 
                  84 inches and equal to or less than 108 inches in combined length 
                  and girth will be classified as an Oversize 1 (OS1) package. 
                  The transportation charges for an Oversize 1 (OS1) package will 
                  be the same as a 30-lb. package being transported under the 
                  same circumstances.<br>
                  <br>
                </li>
                <li>A package weighing less than 50 lbs. and measuring more than 
                  108 inches in combined length and girth will be classified as 
                  an Oversize 2 (OS2) package. The transportation charges for 
                  an Oversize 2 (OS2) package will be the same as a 50-lb. package 
                  being transported under the same circumstances. </li>
              </ul></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>Height: 
              <input name="FEDEX_HEIGHT" type="text" id="FEDEX_HEIGHT" value="10" size="4" maxlength="4"> 
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>Width: 
              <input name="FEDEX_WIDTH" type="text" id="FEDEX_WIDTH" value="10" size="4" maxlength="4"> 
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>Length: 
              <input name="FEDEX_LENGTH" type="text" id="FEDEX_LENGTH" value="10" size="4" maxlength="4"> 
              <font color="#FF0000"> * This is the measurement of the longest 
              side</font></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>Measurement Unit: 
              <input type="radio" name="FEDEX_DIM_UNIT" value="in" checked>
              Inches 
              <input type="radio" name="FEDEX_DIM_UNIT" value="cm">
              Centimeters</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> <input type="submit" name="Submit" value="Submit" class="submit2"></td>
          </tr>
        </table>
		</form>
		<% end if %>
<!--#include file="AdminFooter.asp"-->