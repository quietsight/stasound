<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Manual Entry Method - Step 1" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<form name="form1" method="post" action="../includes/PageCreateTaxSettings.asp" class="pcForms">
    <input type="hidden" name="taxfile" value="0">
    <input type="hidden" name="Page_Name" value="taxsettings.asp">
    <input type="hidden" name="refpage" value="AdminTaxSettings_manual2.asp">
    <table class="pcCPcontent">	
        <tr> 
            <td colspan="2"><div>You can manually enter tax rates based on location and/or product. Consult your local tax authority to determine the tax laws that you need to adhere to.</div>
            <div style="padding-top: 8px;">The following settings apply to all tax rates that you will add:</div></td>
        </tr>
        <tr> 
            <td nowrap="nowrap">Include shipping charges?</td>
            <td width="80%">
            <input type="radio" name="TaxonCharges" value="0" checked class="clearBorder"> No 
            <input type="radio" name="TaxonCharges" value="1" <% If pTaxonCharges=1 then%>checked<% end if %> class="clearBorder"> Yes
            </td>
        </tr>
        <tr> 
            <td nowrap="nowrap">Include handling fees?</td>
            <td>
            <input type="radio" name="TaxonFees" value="0" checked class="clearBorder"> No 
            <input type="radio" name="TaxonFees" value="1" <% If pTaxonFees=1 then%>checked<% end if %> class="clearBorder"> Yes          
            </td>
        </tr>
        <tr> 
            <td nowrap="nowrap">Calculate tax based on:</td>
            <td>
            <input type="radio" name="taxshippingaddress" value="0" checked class="clearBorder"> Billing address 
            <input type="radio" name="taxshippingaddress" value="1" <% If ptaxshippingaddress="1" then%>checked<% end if %> class="clearBorder"> Shipping address</td>
        </tr>
        <tr> 
            <td nowrap="nowrap">Tax wholesale Customer?</td>
            <td>
            <input type="radio" name="taxwholesale" value="0" checked class="clearBorder"> No 
            <input type="radio" name="taxwholesale" value="1" <% If ptaxwholesale="1" then%>checked<% end if %> class="clearBorder"> Yes						
            </td>
        </tr>
        <tr> 
            <td nowrap="nowrap">&nbsp;</td>
        </tr>
        <tr> 
            <td colspan="2">Since you can enter more then one tax rule, ProductCart gives the ability to <strong>show different types of taxes separately</strong> to the customer. For example, this is useful for Canadian online stores (more on <a href="http://wiki.earlyimpact.com/productcart/tax_manual#tax_by_zone_for_canadian_online_stores" target="_blank">tax calculation for Canada-based stores</a>).</td>
        </tr>
        <tr> 
            <td nowrap="nowrap">Display taxes separately?</td>
            <td>
            <input type="radio" name="taxseparate" value="0" checked class="clearBorder"> No 
            <input type="radio" name="taxseparate" value="1" <% If ptaxseparate="1" then%>checked<% end if %> class="clearBorder"> Yes
            </td>
        </tr>
        <tr> 
            <td colspan="2">&nbsp;</td>
        </tr>
        <tr> 
            <td colspan="2" align="center">
            <input type="submit" name="Submit" value="Update & Continue" class="submit2">&nbsp;
            <input type="button" name="back" value="Back" onClick="javascript:history.back()">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->