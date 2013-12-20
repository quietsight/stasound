<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="View/Edit Tax Settings - Tax File Lookup Method" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
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
if request.QueryString("nofilename")="1" then        
	msg="File name is a required field. The <strong>file name</strong> to be entered is the name of the tax data file that you uploaded in <strong>Step 2</strong> below (e.g. &quot;CA_complete.csv&quot;)."     
end if 
%>

<form name="form1" method="post" action="../includes/PageCreateTaxSettings.asp" class="pcForms">
	<input type="hidden" name="taxfile" value="1" checked>
	<input type="hidden" name="refpage" value="AdminTaxsettings.asp">
    
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
                	How do you know which sales tax rate applies to an order? Let ProductCart automatically locate the correct tax rate. <a href="http://www.earlyimpact.com/productcart/support/updates/taxlink.asp" target="_blank">Order an updated sales tax rates file</a>, which can help you apply the correct sales tax to your customers' orders.
                <a href="http://wiki.earlyimpact.com/productcart/tax_data_file" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="More information on this topic" border="0"></a></td>
            </tr>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr>
                <td align="right" nowrap valign="top">Step 1:</td>
                <td><a href="http://www.earlyimpact.com/productcart/support/updates/taxlink.asp" target="_blank"><strong>Order</strong> an updated tax file (or find out whether you need one)</a>
                </td>
            </tr>
            <tr>
                <td align="right" nowrap valign="top">Step 2:</td>
                <td><a href="#" onClick="window.open('taxuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><strong>Upload</strong> the file to your store</a></td>
            </tr>
            <tr>
                <td align="right" nowrap valign="top">Step 3:</td>
                <td>File name: <input type="text" name="taxfilename" value="<%=ptaxfilename%>" size="30">
                    <input type="hidden" name="Page_Name" value="taxsettings.asp">
                    <input type="hidden" name="TaxonCharges" value="0">
                    <input type="hidden" name="TaxonFees" value="0">
                </td>
            </tr>
            <tr>
                <td align="right" nowrap valign="top">Step 4:</td>
                <td>Tax wholesale customers?
                <input type="radio" name="taxwholesale" value="0" checked>No 
                <input type="radio" name="taxwholesale" value="1" <% If ptaxwholesale="1" then%>checked<% end if %>>Yes</td>
            </tr>          
            <tr>
                <td align="right" nowrap valign="top">Step 5:</td>
                <td>
                	<table width="100%" border="0" cellspacing="0" cellpadding="4">
						<tr> 
                          <td colspan="6">
                          Fallback Tax Rate <a href="http://wiki.earlyimpact.com/productcart/tax_data_file#configuring_productcart_s_tax_module" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="Fallback tax rates" width="16" height="16" border="0"></a></td>
                		</tr>
							<% 
								stateArray=split(ptaxRateState,", ")
								if ptaxSNH<>"" then
									taxSNHArray=split(ptaxSNH,", ")
								end if
								rateArray=split(ptaxRateDefault,", ")
							%>
                            <tr> 
                              <th>State</th>
                              <th>Tax Rate</th>
                              <th><div align="center">Tax Shipping</div></th>
                              <th><div align="center">Tax Shipping &amp; Handling Together</div></th>
                              <th colspan="2"><div align="center">Do Not Tax Shipping and Handling</div></th>
                            </tr>
                            <tr>
                            	<td colspan="6" class="pcCPspacer"></td>
                            </tr>
								<% 				
								if ubound(stateArray)=0 then
									if ptaxSNH<>"" then
										strTaxSNH=taxSNHArray(0)
									else
										strTaxSNH="NN"
									end if
								%>
                                    <tr> 
                                      <td width="11%"><%=stateArray(0)%> <input type="hidden" name="taxRateState" value="<%=stateArray(0)%>"></td>
                                      <td width="15%" nowrap><input name="taxRateDefault" type="text" value="<%=rateArray(0)%>" size="6" maxlength="20"> %</td>
                                      <td><div align="center">
                                          <input type="radio" name="taxSNH<%=stateArray(0)%>" value="YN" <% if strTaxSNH="YN" then%>checked<% end if %>>
                                        </div></td>
                                      <td><div align="center">
                                          <input type="radio" name="taxSNH<%=stateArray(0)%>" value="YY" <% if strTaxSNH="YY" then%>checked<% end if %>>
                                        </div></td>
                                      <td><div align="center">
                                          <input type="radio" name="taxSNH<%=stateArray(0)%>" value="NN" <% if strTaxSNH="NN" then%>checked<% end if %>>
                                        </div></td>
                                      <td width="7%" align="right"><a href="../includes/PageCreateTaxSettings.asp?sa=<%=stateArray(0)%>" title="Delete Fallback Tax Rate"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete Fallback Tax Rate" border="0"></a></td>
                                    </tr>
								<% else
                                    for i=0 to ubound(stateArray)-1
                                        tShowInfo=0
                                        if stateArray(i)<>"" then
                                            if ptaxSNH<>"" then
                                                strTaxSNH=taxSNHArray(i)
                                            else
                                                strTaxSNH="NN"
                                            end if %>
                                            <tr> 
                                              <td width="11%"><%=stateArray(i)%><input type="hidden" name="taxRateState" value="<%=stateArray(i)%>"></td>
                                              <td width="15%" nowrap><input name="taxRateDefault" type="text" value="<%=rateArray(i)%>" size="6" maxlength="20"> %</td>
                                              <td><div align="center"> 
                                                    <% if strTaxSNH="YN" then%>
                                                    <input type="radio" name="taxSNH<%=stateArray(i)%>" value="YN" checked>
                                                    <% else %>
                                                    <input type="radio" name="taxSNH<%=stateArray(i)%>" value="YN">
                                                  <% end if %>
                                                  </div>
                                              </td>
                                              <td><div align="center"> 
                                                  <% if strTaxSNH="YY" then%>
												<input type="radio" name="taxSNH<%=stateArray(i)%>" value="YY" checked>
											<% else %>
												<input type="radio" name="taxSNH<%=stateArray(i)%>" value="YY">
											<% end if %>
                                            </div></td>
                                              <td><div align="center">
                                                  <div align="center"> 
                                                    <% if strTaxSNH="NN" then%>
                                                    <input type="radio" name="taxSNH<%=stateArray(i)%>" value="NN" checked>
                                                    <% else %>
                                                    <input type="radio" name="taxSNH<%=stateArray(i)%>" value="NN">
                                                    <% end if %>
                                                  </div>
                                                </div></td>
                                              <td width="7%" align="right"><a href="../includes/PageCreateTaxSettings.asp?sa=<%=stateArray(i)%>" title="Delete Fallback Tax Rate"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Delete" border="0"></a></td>
                                            </tr>
									<%
										else 
											tShowInfo=1
										end if
									next 
									if tShowInfo=1 then %>
                                    <tr>
                                        <td colspan="6" class="pcCPspacer"></td>
                                    </tr>
                                    <tr> 
                                      <td colspan="6">No fallback tax rates have been set.</td>
                                    </tr>
                                    <%
									end if
								end if %>
                                <tr>
                                    <td colspan="6" class="pcCPspacer"></td>
                                </tr>
                                <tr> 
                                  <td colspan="6">
								  	<% if tShowInfo=1 then %>
                                    	<input type="button" name="Update" value="Add Fallback State Tax Rate" onClick="newWindow2('AddTaxRateState_popup.asp','window2')" class="submit2"> 
                                    <% else %>
                                    	<input type="button" name="Update" value="Add New Fallback State Tax Rate" onClick="newWindow2('AddTaxRateState_popup.asp','window2')" class="submit2"> 
                                    <% end if %>
                                   </td>
                                </tr>
                              </table>
                          </td>
                    </tr>
                    <tr>
                    <td colspan="2"><hr></td>
                    </tr>          
                    <tr> 
                    <td colspan="2" align="center">
                    	<input type="submit" name="Submit" value="Save Settings" class="submit2"> 
                        <input type="hidden" name="taxshippingaddress" value="1">
                    </td>
                    </tr>       
                    <tr> 
                    <td colspan="2">&nbsp;</td>
                    </tr>
				</table>
	</form>
<!--#include file="AdminFooter.asp"-->