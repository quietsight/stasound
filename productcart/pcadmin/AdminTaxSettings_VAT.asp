<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage VAT Settings" %>
<% Section="taxmenu" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%
Dim connTemp,query,rs
%>
<form name="form1" method="post" action="../includes/PageCreateTaxSettings.asp" class="pcForms">
	<input type="hidden" name="taxfile" value="0" checked>
	<input type="hidden" id="refPage" name="refpage" value="AdminTaxsettings.asp">
	<table class="pcCPcontent">	 
		<tr>
			<td colspan="2">
			<p>This feature allows online stores located in countries that use the Value Added Tax (VAT) to correctly display retail taxes (e.g. Europe). The use of VAT is based on a number of <a href="JavaScript:win('helpOnline.asp?ref=451')">assumptions</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=451')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></p>
			</td>
		</tr>	
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">Settings</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td width="20%" valign="top"><p>Default VAT Rate:</p></td>
			<td width="80%"><input name="taxVATrate" type="text" value="<%=ptaxVATrate%>" size="4" maxlength="10"> 
			% (e.g. 20) 
			  <input type="hidden" name="Page_Name" value="taxsettings.asp"> <input type="hidden" name="taxVAT" value="1">			</td>
		</tr>	
		<tr> 
			<td valign="top"><p>Show VAT on product details page?</p></td>
			<td><input type="radio" name="taxdisplayVAT" value="0" checked class="clearBorder"> 
			No <input type="radio" name="taxdisplayVAT" value="1" <% If ptaxdisplayVAT="1" then%>checked<% end if %> class="clearBorder">
			Yes
			<div style="padding-bottom:4px">If set to &quot;Yes&quot;, ProductCart calculates the price without VAT and displays it.</div></td>
		</tr>		  
		<tr>					
			<td valign="top"><p>Include shipping charges?</p></td>
			<td><input type="radio" name="TaxonCharges" value="0" checked class="clearBorder"> 
			No <input type="radio" name="TaxonCharges" value="1" <% If pTaxonCharges=1 then%>checked<% end if %> class="clearBorder">
			Yes
			<div style="padding-bottom:4px">If set to &quot;Yes&quot;, ProductCart assumes that shipping charges include VAT. So when it calculates the total VAT applied to the order, it includes these charges in the equation. If you are using Google Checkout this should be set to &quot;No&quot;.</div> </td>
		</tr>				  
		<tr> 
			<td valign="top"><p>Include handling fees?</p></td>
			<td><input type="radio" name="TaxonFees" value="0" checked class="clearBorder"> 
			No <input type="radio" name="TaxonFees" value="1" <% If pTaxonFees=1 then%>checked<% end if %> class="clearBorder">
			Yes
			<div style="padding-bottom:4px">If set to &quot;Yes&quot;, ProductCart assumes that handling fees include VAT. So when it calculates the total VAT applied to the order, it includes these fees in the equation. If you are using Google Checkout this should be set to &quot;No&quot;. </div></td>
		</tr>				  
		<tr> 
			<td valign="top"><p>Tax wholesale customers?</p></td>
			<td>
			<input type="radio" name="taxwholesale" value="0" checked class="clearBorder">No
			<input type="radio" name="taxwholesale" value="1" <% If ptaxwholesale="1" then%>checked<% end if %> class="clearBorder">Yes<br>
			<div style="padding-bottom:4px">Choose whether or not VAT should be shown when wholesale customers are checking out (i.e. if VAT is not charged to wholesale customers, then wholesale prices don't include it, so there is nothing to show)</div> 
			<input type="hidden" name="taxshippingaddress" value="0" class="clearBorder">
			</td>
		</tr>
		<tr> 
			<td valign="top"><p>Display VAT ID?</p></td>
			<td>
            <input type="radio" name="showVatID" value="0" checked class="clearBorder">No
			<input type="radio" name="showVatID" value="1" <% If pShowVatID="1" then%>checked<% end if %> class="clearBorder">Yes<br>
			<div style="padding-bottom:4px">VAT ID Required? <input name="VatIdReq" type="checkbox" value="1" <% If pVatIDReq="1" then%>checked<% end if %>> (Check this box to make VAT ID a required field)  </div>
            
			</td>
		</tr>
		<tr> 
			<td valign="top"><p>Display National Indentification Number?</p></td>
			<td>
            <input type="radio" name="showSSN" value="0" checked class="clearBorder">No
			<input type="radio" name="showSSN" value="1" <% If pShowSSN="1" then%>checked<% end if %> class="clearBorder">Yes<br>
			<div style="padding-bottom:4px">ID Number Required? <input name="SSNReq" type="checkbox" value="1"  <% If pSSNReq="1" then%>checked<% end if %>> (Check this box to make National ID a required field)  </div>
            
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>
				<input type="submit" name="Submit" value="Save Settings" class="submit2" onClick="javascript:setRefPage('AdminTaxSettings.asp');">&nbsp;
				<input type="button" name="Back" value="Back" onClick="javascript:history.back()" class="ibtnGrey"> 
			</td>
		</tr>		  
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<th colspan="2">VAT Categories</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<% if ptaxVATRate_Code="" then %>
		<tr> 
			<td colspan="2">
			<p>VAT categories allow you to use a different rate on different products&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=452')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>. </p>
			<p><strong>To enable VAT Categories</strong> you need to select your European Union Member State from the drop-down menu and click the &quot;Save&quot; button.</p></td>
		</tr>
		<% end if %>
		<tr> 
			<td valign="top"><p>EU Member State:</p></td>
			<td>
				<%
				call openDB()
				ttaxVATRate_State=""
				query="SELECT pcVATCountries.pcVATCountry_ID, pcVATCountries.pcVATCountry_State, pcVATCountries.pcVATCountry_Code From pcVATCountries Order By pcVATCountry_State ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				%>
				<select name="taxVATRate_Code">
				<option value="">Select an option.</option>
				<%
				if not rs.eof then
					pcArr=rs.getRows()
					set rs=nothing
					intCount=ubound(pcArr,2)
					For i=0 to intCount
						if UCASE(ptaxVATRate_Code)=UCASE(pcArr(2,i)) then
							ttaxVATRate_State=pcArr(1,i)
						end if
						%>
						<option value="<%=pcArr(2,i)%>" <%if UCASE(ptaxVATRate_Code)=UCASE(pcArr(2,i)) then response.write "selected"%>><%=pcArr(1,i) & " (" & pcArr(2,i) & ") "%></option>
					<%Next
				end if
				set rs = nothing
				call closeDB()
				%>
				</select>&nbsp;&nbsp;&nbsp;<input type="button" name="Update" value="Manage EU States" onclick="location='ManageEUStates.asp'">
				<div style="padding-top: 8px;"><span class="pcSmallText">This is the country in which the store is located</span>.</div>
			</td>
		</tr>
		<% 
		if ptaxVATRate_Code<>"" then
			call openDB()
			intCount=0
			query="SELECT pcVATRates.pcVATRate_Category, pcVATRates.pcVATRate_Rate, pcVATRates.pcVATRate_ID "
			query=query&"From pcVATRates "
			query=query&"WHERE pcVATRates.pcVATCountry_Code = '"& ptaxVATRate_Code &"';"
			set rs=Server.CreateObject("ADODB.Recordset")
			set rs=conntemp.execute(query)
			if not rs.eof then %>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr> 
					<td colspan="2">
						<p>To assign products to VAT Rate Categories select the &quot;Add/Remove&quot; link next to a category</p>
					</td>
				</tr>
				<tr>					
					<td colspan="2">
						<%
						pcArr=rs.getRows()
						set rs=nothing
						intCount=ubound(pcArr,2)
						%>
						<table class="pcCPcontent" style="width:100%;">
							<tr>
								<th><%=ttaxVATRate_State%></th><th>Rate</th><th>Products</th>
							</tr>
							<% For x=0 to intCount %>
							<tr>
								<td><%=pcArr(0,x)%></td>
								<td><%=pcArr(1,x)%>%</td>
								<td><a href="manageVATCategories.asp?VATID=<%=pcArr(2,x)%>">Add/ Remove</a></td>
							</tr>
							<% Next %>
						</table>	
					</td>
				</tr>			
			<% else %>
				<tr> 
					<td colspan="2">
						<p><span class="pcCPnotes"><strong><%=ttaxVATRate_State%></strong> has no VAT Categories.</span></p>
						<ul>
							<li><a href="AddVATCategory.asp">Add VAT Category</a></li>
						</ul>			
					</td>
				</tr>
			<% end if %>
			<%							
			set rs = nothing
			call closeDB()
			%>	
		<% end if %>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td></td>
			<td>
				<% 
				'// Determine the Redirect Page
				if ptaxVATRate_Code="" then
					pcv_strRefPage="AddVATCategory.asp"
				else
					pcv_strRefPage="AdminTaxSettings_VAT.asp"
				end if
				%>
				<input type="submit" name="Submit" value="Save Settings" class="submit2" onClick="javascript:setRefPage('<%=pcv_strRefPage%>');">&nbsp;
				<% if ptaxVATRate_Code<>"" then %>
				<input type="button" name="ManageVATCategories" value="Manage VAT Categories" onclick="location='viewVAT.asp';" class="ibtnGrey">&nbsp;
				<% end if %>
				<input type="button" name="back" value="Finished" onClick="location.href='menu.asp'" class="ibtnGrey">
				<script language="javascript">
				function setRefPage(pagename) {					
					document.getElementById('refPage').value=pagename;
				}				
				</script>
			</td>
		</tr>	
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->