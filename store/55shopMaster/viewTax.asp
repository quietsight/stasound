<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
Response.Buffer = true
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.CacheControl = "No-Store"
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
%>
<% pageTitle="View/Edit Tax Settings - Manual Entry Method - Summary" %>
<% section="misc" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rsTaxLoc, connTemp, strSQL, pid
%>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message

call openDb()

'//////////////////////////////////////////////////
'// TAX Rules by Product
'//////////////////////////////////////////////////

query="SELECT TaxPrd.idTaxPerProduct, TaxPrd.idproduct, TaxPrd.taxPerProduct, TaxPrd.CountryCode, TaxPrd.stateCode, TaxPrd.zip, products.description FROM TaxPrd, products WHERE TaxPrd.idproduct = products.idproduct ORDER BY products.description ASC"
Set rsTaxPrd=Server.CreateObject("ADODB.Recordset")   
rsTaxPrd.Open query, connTemp, adOpenStatic, adLockReadOnly

'FormatPercent(rsTaxPrd("taxPerProduct"),2,0,0,0)

%>
<form method="POST" action="product_action.asp" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="8"><div style="float: right; padding-top: 8px;"><a href="AddTaxPerPrd.asp">Add new</a></div><h2>Tax Rules by Product</h2></td>
		</tr>
		<% If rsTaxPrd.eof Then %>
			<tr> 
				<td colspan="8">No tax rules have been set.</td>
			</tr>
		<% Else %>
		<tr> 
			<th align="center" nowrap>ID</th>
			<th nowrap>Tax</th>
			<th nowrap>Country</th>
			<th nowrap>State</th>
			<th nowrap>Zip</th>
			<th nowrap colspan="3">Product</th>
		</tr>
	
		<% 
				Do While NOT rsTaxPrd.EOF
		%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td width="2%"><%=rsTaxPrd("idTaxPerProduct") %></td>
				<td width="5%"><%=FormatPercent(rsTaxPrd("taxPerProduct"),2,0,0,0) %></td>
				<td><%=rsTaxPrd("CountryCode") %></td>
				<td><%=rsTaxPrd("stateCode") %></td>
				<td><%=rsTaxPrd("zip") %></td>
				<td width="50%"><%=rsTaxPrd("description")%></td>
				<td width="2%"><a href="modtaxPrd.asp?idTaxPerProduct=<%= rsTaxPrd("idTaxPerProduct") %>&mode=MOD"><img src="images/pcIconGo.jpg" border="0"></a></td>
				<td width="2%"><a href="javascript:if (confirm('You are about to permanantly delete this tax rule from the database. Are you sure you want to complete this action?')) location='modtaxPrd.asp?idTaxPerProduct=<%= rsTaxPrd("idTaxPerProduct") %>&mode=DEL'"><img src="images/pcIconDelete.jpg" border="0"></a>
				</td>
			</tr>
	
			<%
			rsTaxPrd.MoveNext
			Loop
			set rsTaxPrd=nothing
			End If
			%>    
	</table>
</form>
<%
'//////////////////////////////////////////////////
'// End TAX Rules by Product
'//////////////////////////////////////////////////

'//////////////////////////////////////////////////
'// TAX Rules by Location
'//////////////////////////////////////////////////

query="SELECT idTaxPerPlace,taxLoc,CountryCode,stateCode,zip,taxDesc FROM TaxLoc"
Set rsTaxLoc=Server.CreateObject("ADODB.Recordset")   
rsTaxLoc.Open query, connTemp, adOpenStatic, adLockReadOnly
%>
<form method="POST" action="product_action.asp" class="pcForms">
	<table class="pcCPcontent">
			<tr>
				<td colspan="8" class="pcCPspacer"></td>
			</tr>
		<tr> 
			<td colspan="8"><div style="float: right; padding-top: 8px;"><a href="AddTaxPerPlace.asp">Add new</a></div><h2>Tax Rules by Location</h2></td>
		</tr>
		
		<% If rsTaxLoc.eof Then %>
		<tr> 
			<td colspan="8">No tax rules have been set.</td>
		</tr>
		
		<% Else %>	
		<tr> 
			<th align="center" nowrap>ID</th>
			<th nowrap>Tax</th>
			<th nowrap>Country</th>
			<th nowrap>State</th>
			<th nowrap>Zip</th>
			<th nowrap colspan="3">Description</th>
		</tr>

		<%
			Do While NOT rsTaxLoc.EOF
				idTaxPerPlace=rsTaxLoc("idTaxPerPlace")
				taxLoc=rsTaxLoc("taxLoc")
				CountryCode=rsTaxLoc("CountryCode")
				stateCode=rsTaxLoc("stateCode")
				zip=rsTaxLoc("zip")
				taxDesc=rsTaxLoc("taxDesc")  
				%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td width="2%" align="center"><%= idTaxPerPlace %></td>
					<td width="5%"><%= FormatPercent( taxLoc,2,0,0,0) %></td>
					<td><%= CountryCode %></td>
					<td><%= stateCode %></td>
					<td><%= zip %></td>
					<td width="50%"><%=taxDesc%></td>
					<td width="2%"><a href="modtaxloc.asp?idTaxPerPlace=<%= idTaxPerPlace %>&mode=MOD"><img src="images/pcIconGo.jpg" border="0"></a></td>
					<td width="2%"><a href="javascript:if (confirm('You are about to permanantly delete this tax rule from the database. Are you sure you want to complete this action?')) location='modtaxloc.asp?idTaxPerPlace=<%= idTaxPerPlace %>&mode=DEL'"><img src="images/pcIconDelete.jpg" border="0"></a></td>
				</tr>
				<% rsTaxLoc.MoveNext
			Loop
		End If %>         
	</table>
</form>

<%
'//////////////////////////////////////////////////
'// End TAX Rules by Location
'//////////////////////////////////////////////////

'//////////////////////////////////////////////////
'// TAX Rules by Zone
'//////////////////////////////////////////////////

query="SELECT pcTaxZoneRates.pcTaxZoneRate_ID, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneRates.pcTaxZoneRate_Type, pcTaxZoneRates.pcTaxZoneRate_Order, pcTaxZoneRates.pcTaxZoneRate_Rate, pcTaxZoneRates.pcTaxZoneRate_ApplyToSH, pcTaxZoneRates.pcTaxZoneRate_Taxable, pcTaxZoneDescriptions.pcTaxZoneDesc FROM (pcTaxZoneRates INNER JOIN pcTaxZonesGroups ON pcTaxZoneRates.pcTaxZoneRate_ID = pcTaxZonesGroups.pcTaxZoneRate_ID) INNER JOIN pcTaxZoneDescriptions ON pcTaxZonesGroups.pcTaxZoneDesc_ID = pcTaxZoneDescriptions.pcTaxZoneDesc_ID ORDER BY pcTaxZoneRates.pcTaxZoneRate_Name;"
Set rsTaxZone=Server.CreateObject("ADODB.Recordset")   
rsTaxZone.Open query, connTemp, adOpenStatic, adLockReadOnly
%>
	<form method="POST" action="product_action.asp" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="8" class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td colspan="8"><div style="float: right; padding-top: 8px;"><a href="AddTaxPerZone.asp">Add new</a></div><h2>Tax Rules by Zone&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=450')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></h2></td>
			</tr>
			<% If rsTaxZone.eof Then %>
				<tr> 
					<td colspan="8">No tax rules have been set. <br />A <strong>zone</strong> is a &quot;group&quot; of locations. For example, a group of states or provinces might share the same tax.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=450')">More details</a>.</td>
				</tr>
				
			<% Else %>
			<tr> 
				<th width="30%" nowrap>Description</th>
				<th nowrap>Type</th>
				<th nowrap>Tax Rate</th>
				<th nowrap>Tax S &amp; H</th>
				<th align="center" nowrap>Taxable</th>
				<th align="center" nowrap>Order</th>
				<th width="5%" align="center" nowrap colspan="2"></th>
			</tr>
			<%
				strCol="#E1E1E1"
				Do While NOT rsTaxZone.EOF
					intTaxZonesGroupID=rsTaxZone("pcTaxZoneRate_ID")
					strTaxZoneRateName=rsTaxZone("pcTaxZoneRate_Name")
					strTaxZoneRateType=rsTaxZone("pcTaxZoneRate_Type")
					intTaxZoneRateOrder=rsTaxZone("pcTaxZoneRate_Order")
					dblTaxZoneRateRate=rsTaxZone("pcTaxZoneRate_Rate")
					intTaxZoneRateApplyToSH=rsTaxZone("pcTaxZoneRate_ApplyToSH")
					intTaxZoneRateTaxable=rsTaxZone("pcTaxZoneRate_Taxable")  
					strTaxZoneDesc=rsTaxZone("pcTaxZoneDesc")
					If strCol <> "#FFFFFF" Then
						strCol="#FFFFFF"
					Else 
						strCol="#E1E1E1"
					End If %>
					<tr bgcolor="<%= strCol %>"> 
						<td><a href="AddTaxPerZone.asp?idTaxZonesGroupID=<%= intTaxZonesGroupID %>&mode=MOD"><%= strTaxZoneRateName&" ("&strTaxZoneDesc&")" %></a></td>
						<td><%= strTaxZoneRateType %></td>
						<td><%= FormatPercent( dblTaxZoneRateRate,4,0,0,0) %></td>
						<td><% if trim(intTaxZoneRateApplyToSH)=1 then %>Yes<% else %>No<% end if %></td>
						<td><% if trim(intTaxZoneRateTaxable)=1 then %>Yes<% else %>No<% end if %></td>
						<td align="center"><%= intTaxZoneRateOrder %></td>
						<td align="right"><a href="AddTaxPerZone.asp?idTaxZonesGroupID=<%= intTaxZonesGroupID %>&mode=MOD"><img src="images/pcIconGo.jpg" border="0" alt="Modify This Tax Rule"></a>&nbsp;<a href="javascript:if (confirm('You are about to permanantly delete this tax rule from the database. Are you sure you want to complete this action?')) location='modtaxloc.asp?idTaxPerPlace=<%= intTaxZonesGroupID %>&mode=DELZONE'"><img src="images/pcIconDelete.jpg" border="0" alt="Delete This Tax Rule"></a></td>
					</tr>
					<% rsTaxZone.MoveNext
				Loop
			End If %>         
		</table>
	</form>

	<form name="form1" method="post" action="../includes/PageCreateTaxSettings.asp" class="pcForms">
		<table class="pcCPcontent" style="margin-top: 20px;">	
			<tr>
				<td colspan="2"><h2>Other Tax Calculation Settings</h2></td>
			</tr>
			<tr> 
				<td colspan="2">The following settings apply to all tax rates listed on this page:</td>
			</tr>
			<tr> 
				<td nowrap="nowrap">Include shipping charges?</td>
				<td width="80%">
					<% if pTaxonCharges=1 then %>
					<input type="radio" name="TaxonCharges" value="0" class="clearBorder"> No 
					<input type="radio" name="TaxonCharges" value="1" checked class="clearBorder"> Yes 
					<% else %>
					<input type="radio" name="TaxonCharges" value="0" checked class="clearBorder"> No 
					<input type="radio" name="TaxonCharges" value="1" class="clearBorder"> Yes 
					<% end if %>
				</td>
			</tr>
			<tr> 
				<td nowrap="nowrap">Include handling fees?</td>
				<td>
					<% if pTaxonFees=1 then %>
						<input type="radio" name="TaxonFees" value="0" class="clearBorder"> No 
						<input type="radio" name="TaxonFees" value="1" checked class="clearBorder"> Yes 
					<% else %>
						<input type="radio" name="TaxonFees" value="0" checked class="clearBorder"> No 
						<input type="radio" name="TaxonFees" value="1" class="clearBorder"> Yes 
					<% end if %>
					<input type="hidden" name="Page_Name" value="taxsettings.asp">
					<input type="hidden" name="refpage" value="viewTax.asp">
				</td>
			</tr>
			<tr> 
				<td nowrap="nowrap">Tax is calculated on</td>
				<td>
					<% if ptaxshippingaddress=1 then %>
					<input type="radio" name="taxshippingaddress" value="0" class="clearBorder"> Billing address 
					<input type="radio" name="taxshippingaddress" value="1" checked class="clearBorder"> Shipping address 
					<% else %>
					<input type="radio" name="taxshippingaddress" value="0" checked class="clearBorder"> Billing address 
					<input type="radio" name="taxshippingaddress" value="1" class="clearBorder"> Shipping address 
					<% end if %>
				</td>
			</tr>
			<tr> 
				<td nowrap="nowrap">Display taxes Separately</td>
				<td>
					<input type="radio" name="taxseparate" value="0" checked class="clearBorder"> No 
					<input type="radio" name="taxseparate" value="1" <% If ptaxseparate="1" then%>checked<% end if %> class="clearBorder"> Yes
				</td>
			</tr>
			<tr> 
				<td nowrap="nowrap">Tax Wholesale Customers?</td>
				<td>
					<input type="radio" name="taxwholesale" value="0" checked class="clearBorder"> No 
					<input type="radio" name="taxwholesale" value="1" <% If ptaxwholesale="1" then%>checked<% end if %> class="clearBorder"> Yes
				</td>
			</tr>
			<tr> 
				<td colspan="2"><hr></td>
			</tr>
			<tr> 
				<td colspan="2">
					<input type="submit" name="Submit" value="Update" class="submit2">
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="2"><h2>Tax Exemptions</h2></td>
			</tr>
			<tr> 
				<td colspan="2">You can create an exemption list that include specific products or categories:</td>
			</tr>
			<tr> 
				<td colspan="2">
					<ul>
						<li>Set state/province-specific product/category tax exemptions for <a href="manageTaxEpt.asp">tax by location</a> rules</li>
						<li>Set zone-specific product/category tax exemptions for <a href="manageTaxEptZone.asp">tax by zone</a> rules</li>
					</ul>
				</td>
			</tr>

			<tr>
				<td colspan="2"><hr></td>
			</tr>
			<tr> 
				<td colspan="2">Select 'Switch to Tax File' if you no longer wish to enter tax rates manually, and want to switch to using a tax rate database (US customers only). Select 'Switch to VAT' if your prices now include the Value Added Tax.</td>
			</tr>
			<tr> 
				<td colspan="2">
					<input type="button" value="Switch to Tax File" onClick="location.href='AdminTaxSettings_file.asp'">&nbsp;
					<input type="button" value="Switch to VAT" onClick="location.href='AdminTaxSettings_VAT.asp'">
					<input type="hidden" name="taxfile" value="0">
				</td>
			</tr>
		</table>
	</form>
<!--#include file="AdminFooter.asp"-->