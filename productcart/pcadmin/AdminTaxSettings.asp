<% Response.CacheControl = "no-cache" %>
<% Response.Expires = -1 %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings" %>
<% Section="layout" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
function newWindow2(file,window) {
catWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
if (catWindow.opener == null) catWindow.opener = self;
}
//-->
</script>
<%
Dim connTemp,query,rs
%>
<form class="pcForms">
<table class="pcCPcontent">	
	<tr>
		<td colspan="2">ProductCart can calculate taxes in three ways: using a tax file (database), using rates that you manually enter, or assuming that a Value Added Tax is included in the prices (VAT). In all cases, make sure to consult your local tax authority for information about the tax laws that you need to adhere to. Here is a summary of your current settings.</td>
	</tr>
	<tr> 
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<%
	IF ptaxfile=1 THEN ' Store is using a tax file
	 %>
		<tr> 
			<td colspan="2">
			<% if request.QueryString("nofile")="0" then %>
			<div class="pcCPmessageSuccess">You are currently using a tax data file.</div>
			<% elseif request.QueryString("nofile")="1" then %>
			<div class="pcCPmessage">The system was not able to locate the tax file that you specified in your 'tax' folder. Please check that you have uploaded the file and that you have typed in the file name correctly, including the file extension. <a href="#" onClick="window.open('taxuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><strong>Upload the file now</strong></a>. For more information about obtaining a <u>properly formatted</u> tax data file, <a href="http://www.earlyimpact.com/productcart/support/updates/taxes.asp" target="_blank">click here</a>.</div>
			<% end if %>
            </td>
		</tr>
        <tr>
        	<td width="20%" nowrap>Tax file name:</td>
			<td><strong><%=ptaxfilename%></strong></td>
        </tr>
        <tr>
        	<td>Tax Wholesale Customers:</td>
            <td>
			<% If ptaxwholesale="1" then
				response.write "Yes"
			else
				response.write "No"
			end if %>
            </td>
		<tr>
        	<td nowrap valign="top">Fallback States Tax Rates</td>
            <td>
            	<table class="pcCPcontent">
                  <tr bgcolor="#FFFF99"> 
                    <td>State</td>
                    <td>Tax Rate</td>
                    <td><div align="center">Tax Shipping</div></td>
                    <td><div align="center">Tax Shipping and Handling Together</div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% stateArray=split(ptaxRateState,", ")
                                rateArray=split(ptaxRateDefault,", ")
                                if ptaxSNH<>"" then
                                    taxSNHArray=split(ptaxSNH,", ")
                                end if
                                if ubound(stateArray)=0 then %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td width="9%"><%=stateArray(0)%> <input type="hidden" name="taxRateState" value="<%=stateArray(0)%>"> 
                    </td>
                    <td><%=rateArray(0)%>%</td>
                    <% if ptaxSNH<>"" then
                                        select case taxSNHArray(0)
                                        case "YY"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether="Yes"
                                        case "YN"
                                            taxShippingAlone="Yes"
                                            taxShippingAndHandlingTogether=""
                                        case "NN"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether=""
                                        end select
                                    else
                                        taxShippingAlone=""
                                        taxShippingAndHandlingTogether=""
                                    end if %>
                    <td><div align="center"><%=taxShippingAlone%></div></td>
                    <td><div align="center"><%=taxShippingAndHandlingTogether%></div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% else
                                for i=0 to ubound(stateArray)-1 %>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    <td><%=stateArray(i)%> <input type="hidden" name="taxRateState" value="<%=stateArray(i)%>"> 
                    </td>
                    <td><%=rateArray(i)%>%</td>
                    <%if ptaxSNH<>"" then
                                        select case taxSNHArray(i)
                                            case "YY"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether="Yes"
                                        case "YN"
                                            taxShippingAlone="Yes"
                                            taxShippingAndHandlingTogether=""
                                        case "NN"
                                            taxShippingAlone=""
                                            taxShippingAndHandlingTogether=""
                                        end select
                                    else
                                        taxShippingAlone=""
                                        taxShippingAndHandlingTogether=""
                                    end if %>
                    <td><div align="center"><%=taxShippingAlone%></div></td>
                    <td><div align="center"><%=taxShippingAndHandlingTogether%></div></td>
                    <td>&nbsp;</td>
                  </tr>
                  <% next
                    end if %>
                </table>
            </td>
        </tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
        <tr>
        	<td colspan="2">
            <div style="margin-bottom: 20px;">
			<input type="button" value="Edit Settings" onClick="location.href='AdminTaxSettings_file.asp'" class="submit2">
            &nbsp;<input type="button" onClick="location.href='manageTaxEpt.asp'" value="Set tax exemptions for US states">
            </div>
            <hr>
            <div style="margin-top: 10px">
            Select 'Switch to Manual Tax Calculation' or 'Switch to VAT' if you no longer wish to use this tax file, and would like to swith to an alternative tax calculation method.
            </div>
            <div style="margin-top: 10px">
			<input type="button" onClick="location.href='AdminTaxSettings_VAT.asp'" value="Switch to VAT">
            &nbsp;
			<input type="button" onClick="location.href='AdminTaxSettings_manual.asp'" value="Switch to Manual Tax Calculation">
            </div>
            </td>
		</tr>
	<% 
	END IF
	' End store using a tax file 
	%>
										
	<% if ptaxfile=0 AND ptaxsetup=1 then 
		if ptaxVAT="1" then
		' Store is using VAT		
		%>
			<tr> 
				<th colspan="2">VAT (Value Added Tax)</th>
			</tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
			<tr> 
			<td colspan="2">
				<div class="pcCPmessageSuccess">You are currently setup to use the Value Added Tax (prices include taxes).</div>
            </td>
            </tr>
            <tr>
            	<td nowrap width="20%">Default VAT Rate:</td>
                <td><strong><%=ptaxVATrate%></strong></td>
            </tr>
            <tr>
            	<td nowrap>EU Member State:</td>
                <td>
				<%
				ttaxVATRate_State = "Not Selected - Use Default Rate"
				call openDB()
				query="SELECT pcVATCountries.pcVATCountry_State From pcVATCountries WHERE pcVATCountries.pcVATCountry_Code = '"& ptaxVATRate_Code &"' Order By pcVATCountry_State ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				if not rs.eof then
					ttaxVATRate_State=rs("pcVATCountry_State")
				end if
				set rs = nothing
				call closeDB()
				%>			
				<%=ttaxVATRate_State%>
                </td>
            </tr>
            <tr>
            	<td nowrap>Show VAT on product details page:</td>
                <td><% If ptaxdisplayVAT="1" then response.write "Yes" else response.write "No"	end if %></td>
            </tr>
            <tr>
            	<td nowrap>Include shipping charges:</td>
                <td><% If pTaxonCharges="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
            <tr>
            	<td nowrap>Include handling fees:</td>
				<td><% If pTaxonFees="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
            <tr>
            	<td nowrap>Tax Wholesale Customers:</td>
                <td><% If ptaxwholesale="1" then response.write "Yes" else response.write "No" end if %></td>
            </tr>
			<tr> 
				<td class="pcCPspacer" colspan="2"></td>
			</tr>
			<tr> 
				<td colspan="2"><input type="button" value="Edit Settings" onClick="location.href='AdminTaxSettings_VAT.asp'" class="submit2"></td>
			</tr>
            <tr>
            	<td colspan="2">
                <hr>
				Select 'Switch to Manual Tax Calculation' or 'Switch to Using a Tax File' if you no longer wish to use this option.
                <div style="margin-top: 10px;">
                	<input name="button" type="button" onClick="location.href='AdminTaxSettings_manual.asp'" value="Switch To Manual Tax Calculation">		
                    &nbsp;<input name="button2" type="button" onClick="location.href='AdminTaxSettings_file.asp'" value="Switch To Using a Tax File">
                </div>
                </td>
			</tr>
		<% else
		' End store using VAT
		' The store is using manual tax calculation method: redirect to that page
			response.redirect "viewTax.asp"
			response.end
		end if
	end if
	
	if ptaxsetup=0 then
		' Tax settings have not been configured yet: redirect to Tax Wizard
		Response.Redirect "AdminTaxWizard.asp"
		Response.End()
	end if %>
    <tr> 
        <td class="pcCPspacer" colspan="2"></td>
    </tr>
	<tr> 
		<td colspan="2" align="right">
        <hr>
		<input type="button" name="back" value="Finished" onClick="location.href='menu.asp'">
		</td>
	</tr>
    <tr> 
        <td class="pcCPspacer" colspan="2"></td>
    </tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->