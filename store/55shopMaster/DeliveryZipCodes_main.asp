<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Delivery Zip Codes" %>
<% section="shipOpt" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%PmAdmin="1*4*"%><!--#include file="adminv.asp"--> 
<%
Dim rs, connTemp, query
%>
<!--#include file="AdminHeader.asp"-->
    <table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                Use this feature <u>only if you wish to closely define the geographic area</u> where your orders can be shipped. To turn the <strong>Delivery Zip Codes</strong> validator on or off, use the <a href="checkoutoptions.asp#zipcodes">Checkout Options</a> page. This feature is currently <% if DeliveryZip="0" then Response.write("off") else Response.write("on") %>.&nbsp;<a href="http://wiki.earlyimpact.com/productcart/settings-checkout_options#limit_delivery_area_by_zip_code" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" width="16" height="16" border="0"></a>
                <div class="cpOtherLinks">
                    <form name="addnew" method="post" action="DeliveryZipCodes_add.asp?action=add" class="pcForms">
                    Add New Zip Code: <input type="text" name="zipcode" size="10">
                    &nbsp;<input type="submit" name="submit" value="Add New" class="submit2">
                    </form>
                </div>
          </td>
        </tr>
		<tr> 
        	<td colspan="2"><h2>ZIP Codes Currently Allowed</h2></td>
		</tr>

		<%
			call openDb()
			query="SELECT zipcode FROM ZipCodeValidation ORDER BY zipcode ASC"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=connTemp.execute(query)
			If rs.eof Then
		%>
              
        <tr> 
            <td colspan="2">No Zip Codes Found.</td>
        </tr>
              
		<%
			Else 
            	Do While NOT rs.EOF
            	zipcode=rs("zipcode")
		%>
            <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td width="20%" nowrap><%=zipcode%></td>
				<td width="80%" align="left" nowrap>
                	<a href="DeliveryZipCodes_edit.asp?zipcode=<%=zipcode%>"><img src="images/pcIconGo.jpg" alt="Edit" width="12" height="12" border="0"></a>&nbsp;<a href="javascript:if (confirm('You are about to remove this Zip Code from your database. Are you sure you want to complete this action?')) location='DeliveryZipCodes_delete.asp?zipcode=<%=zipcode%>'"><img src="images/pcIconDelete.jpg" alt="Delete" width="12" height="12" border="0"></a>
               	</td>
			</tr>
              
			<% 
				rs.MoveNext
            	Loop
            End If
			set rs = nothing
			call closeDb()
			%>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>
			<tr>
				<td colspan="2">
                    <div class="cpOtherLinks">
                        <form name="addnew" method="post" action="DeliveryZipCodes_add.asp?action=add" class="pcForms">
                        Add New Zip Code: <input type="text" name="zipcode" size="10">
                        &nbsp;<input type="submit" name="submit" value="Add New" class="submit2">
                        </form>
                    </div>
                </td>
            </tr>
        </table>
<!--#include file="AdminFooter.asp"-->