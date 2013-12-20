<% Response.CacheControl = "no-cache" %>
<% Response.Expires = -1 %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Sales Tax Wizard" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<form name="taxwizard" method="get" action="AdminTaxWizard.asp" class="pcForms">
	<table class="pcCPcontent">
        <tr> 
            <td>
			<% 'Check to see if the Wizard has started
            	If request.QueryString("started") <> "yes" then %>
                <div>The ProductCart <strong>Tax Wizard</strong> will ask you a few questions to determine which sales tax calculation method should be used on your store.</div>
                <div style="margin-top: 10px;">Does your company <strong>pay taxes in the United States?</strong></div>
                <div style="margin-top: 10px;"><input name="us" type="radio" value="yes" checked> Yes <input name="us" type="radio" value="no"> No</div>
                <hr>
                <div style="margin-top: 10px;">
                	<input type="hidden" name="started" value="yes">
                    <input type="submit" name="submit1" value="Continue" class="submit2">
                    &nbsp;<input type="button" value="Back" name="back" onClick="JavaScript:history.go(-1);">
                </div>
            <% end if
			
			'If store is in the US, show states
			If request.QueryString("Submit1") <> "" then
				If request.QueryString("us")="yes" then %>
                    <div>Does your company have a <strong>physical presence</strong> in any of the <strong>following US states</strong>?<br>
        Typically, you have a &quot;physical presence&quot; in all states in which your company has employees and/or sales agents (e.g. where you have an office, a warehouse, a retail store, etc.). If you are unsure of whether or not you have a physical presence in a particular state, please contact the local tax authority.</div>
                    <table width="90%" border="0" align="center" cellpadding="3" cellspacing="0">
                        <tr><td>Alabama (AL) </td><td>Iowa (IA)</td><td>North Carolina (NC)</td><td>Virginia (VA) </td></tr>
                        <tr><td>Alaska (AK) </td><td>Kansas (KS) </td><td>North Dakota (ND) </td><td>Washington (WA) </td></tr>
                        <tr><td>Arizona (AZ) </td><td>	Louisiana (LA) </td><td>Ohio (OH) </td><td>Wisconsin (WI) </td></tr>
                        <tr><td>Arkansas (AR) </td><td>Minnesota (MN) </td><td>Oklahoma (OK) </td><td>Wyoming (WY) </td></tr>
                        <tr><td>California (CA) </td><td>Mississippi (MS) </td><td>Pennsylvania (PA) </td><td>&nbsp;</td></tr>
                        <tr><td>Colorado (CO) </td><td>Missouri (MO) </td><td>South Carolina (SC) </td><td>&nbsp;</td></tr>
                        <tr><td>Florida (FL) </td><td>Nebraska (NE) </td><td>South Dakota (SD) </td><td>&nbsp;</td></tr>
                        <tr><td>Georgia (GA) </td><td>Nevada (NV) </td><td>Tennessee (TN) </td><td>&nbsp;</td></tr>
                        <tr><td>Idaho (ID)</td><td>New Mexico (NM) </td><td>Texas (TX) </td><td>&nbsp;</td></tr>
                        <tr><td>Illinois (IL) </td><td>New York (NY) </td><td>Utah (UT) </td><td>&nbsp;</td></tr>
                    </table>
                    <div style="margin-top: 20px;"><input name="states" type="radio" value="yes" checked> Yes, in one or more states.</div>
                    <div style="margin-top: 5px;"><input name="states" type="radio" value="no"> No, we don't have a physical presence in any of the states listed above.</div>
                    <hr>
                    <div style="margin-top: 10px;">
                        <input type="hidden" name="started" value="yes">
                        <input type="submit" name="submit2" value="Continue" class="submit2">
                        &nbsp;<input type="button" value="Back" name="back" onClick="JavaScript:history.go(-1);">
                    </div>
				<% 'This is not a US store, show VAT vs. manual tax selection
				else %>
                    <div><strong>Do your prices include VAT (Value Added Tax)?</strong></div>
                	<div style="margin-top: 10px;">ProductCart allows you to display prices with and without VAT, and display the VAT included in the order total. If you need to show <u>multiple tax rates</u> separately (e.g. Canada), select &quot;No&quot; below.</div>
                	<div style="margin-top: 10px;"><input name="vat" type="radio" value="yes" checked> Yes (e.g. UK stores)</div>
                    <div style="margin-top: 5px;"><input name="vat" type="radio" value="no"> No, I don't use VAT or I need to show multiple tax rates (e.g. Canada).</div>
                    <hr>
                    <div style="margin-top: 10px;">
                        <input type="hidden" name="started" value="yes">
                        <input type="submit" name="submit3" value="Continue" class="submit2">
                        &nbsp;<input type="button" value="Back" name="back" onClick="JavaScript:history.go(-1);">
                    </div>
				<% end if 'End US vs. International store
			end if 'End submit1 

			if request.QueryString("submit2") <> "" then 'This is a US store
				if request.QueryString("states") = "yes" then 'The store is in one of the state that require a tax file. %>
                    <div><img src="images/taxdatasystems.gif" alt="Tax Data Systems" width="174" height="87" hspace="20" align="right">ProductCart makes it easy to properly calculate sales taxes by allowing you to use updated tax files that will correctly determine the tax rate to be applied to an order. We have partnered with <a href="http://www.earlyimpact.com/productcart/support/updates/taxlink.asp" target="_blank">Tax Data Systems</a> to provide you with a way to <u>easily</u> and <u>accurately</u> calculate sales taxes.</div>
        <div style="margin-top: 10px;">Let ProductCart automatically...</div>
                    <ul>
                        <li>... determine which orders are taxable.</li>
                        <li>... charge sales taxes based on the rate in effect in the location to which the order is shipped. There are cities and postal codes that reside across county lines, so there can be different tax rates in effect in the same city.</li>
                        <li> ... charge the correct tax rate. You are not allowed to overcharge customers by using a tax rate that is higher than the correct tax rate for the applicable to the order.</li>
                        <li>... automatically include or exclude shipping charges and/or shipping &amp; handling fees in the tax calculation based on the tax law applicable to the order.</li>
                    </ul>
                    <div style="margin-top: 10px;">How it works...</div>
                    <ul>
                        <li><a href="http://www.earlyimpact.com/productcart/support/updates/taxlink.asp" target="_blank">Buy a tax rate file</a> from Tax Data Systems</li>
                        <li>Upload the tax rate file to ProductCart</li>
                        <li>Upload a new file to ProductCart every time Tax Data Systems provides you with an updated file (tax rates change frequently).</li>
                    </ul>
                    <hr>
                    <div style="margin-top: 10px;">
                    	<input type="button" name="buy" value="Buy a tax rate file" onClick="window.open('http://www.earlyimpact.com/productcart/support/updates/taxlink.asp')" class="submit2">
                        &nbsp;<input type="button" name="haveit" value="I already purchased it" onClick="location.href='AdminTaxSettings_file.asp'">
                        &nbsp;<input type="button" name="nobuy" value="I'm not interested" onClick="location.href='AdminTaxSettings_manual.asp'">
                        &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey">
                        <input type="hidden" name="refpage" value="viewTax.asp">
                    </div>
				<%	else 'This store does not need to use a tax file %>
                    <div>The US state(s) in which you have a physical precense don't require complex sales tax calculation. You can comply with your local sales tax laws without using a tax file.</div>
                    <div style="margin-top: 10px;">
                    	<input type="button" value="Enter tax rates manually" onClick="document.location.href='AdminTaxSettings_manual.asp';" class="submit2">
                        &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back(-1)">
                </div>
				<% end if 'End store requires tax file
			end if 'This is US store

			if request.QueryString("submit3") <> "" then 'This is an international store
				if request.QueryString("vat") = "yes" then
					response.Redirect "AdminTaxSettings_VAT.asp"
					response.End()
					else
					response.Redirect "AdminTaxSettings_manual.asp"
				end if
			end if 'End this is an international store
			%>
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->