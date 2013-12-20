<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Tax Settings - Manual Entry Method - Tax by Zone" %>
<% section="misc" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/taxsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<%
dim query, conntemp, rs

sMode=Request.Form("Submit")

If sMode <> "" Then
	If sMode="Add" OR sMode="Update" Then
		TaxZoneID=request.Form("TaxZoneID")
		TaxDesc=request.Form("TaxDesc")
		TaxType=request.Form("TaxType")
		TaxOrder=request.Form("TaxOrder")
		if TaxOrder="" then
			TaxOrder=1
		end if
		TaxRate=request.Form("TaxRate")
		TaxRate=((TaxRate)/100)
		TaxApplySH=request.Form("TaxApplySH")
		if TaxApplySH="" then
			TaxApplySH=0
		end if
		TaxTaxable=request.Form("TaxTaxable")
		if TaxTaxable="" then
			TaxTaxable=0
		end if
		TaxLocalZone=request("TaxLocalZone")
		if TaxLocalZone="" then
			TaxLocalZone=0
		end if
		
		call opendb()
		
		if sMode="Update" then
			intUpdateID=request("UpdateID")

			query="UPDATE pcTaxZoneRates SET pcTaxZoneRate_Name='"&TaxDesc&"', pcTaxZoneRate_Type='"&TaxType&"', pcTaxZoneRate_Order="&TaxOrder&", pcTaxZoneRate_Rate="&TaxRate&", pcTaxZoneRate_ApplyToSH="&TaxApplySH&", pcTaxZoneRate_Taxable="&TaxTaxable&", pcTaxZoneRate_LocalZone="&TaxLocalZone&" WHERE pcTaxZoneRate_ID="&intUpdateID&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			
			'if orphaned, insert new group
			if request("orphaned")="1" then
				query="INSERT INTO pcTaxZonesGroups (pcTaxZoneRate_ID, pcTaxZoneDesc_ID) VALUES ("&intUpdateID&","&TaxZoneID&");"
			else
				query="UPDATE pcTaxZonesGroups SET pcTaxZoneDesc_ID="&TaxZoneID&" WHERE pcTaxZoneRate_ID="&intUpdateID&";"
			end if

			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
		else
			query="INSERT INTO pcTaxZoneRates (pcTaxZoneRate_Name, pcTaxZoneRate_Type, pcTaxZoneRate_Order, pcTaxZoneRate_Rate, pcTaxZoneRate_ApplyToSH, pcTaxZoneRate_Taxable, pcTaxZoneRate_LocalZone) VALUES ('"&TaxDesc&"', '"&TaxType&"', "&TaxOrder&", "&TaxRate&", "&TaxApplySH&", "&TaxTaxable&", "&TaxLocalZone&");"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			'//Get last enter
			query="SELECT pcTaxZoneRates.pcTaxZoneRate_ID FROM pcTaxZoneRates ORDER BY pcTaxZoneRates.pcTaxZoneRate_ID DESC;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			intTaxZoneRateID=rs("pcTaxZoneRate_ID")
			query="INSERT INTO pcTaxZonesGroups (pcTaxZoneRate_ID, pcTaxZoneDesc_ID) VALUES ("&intTaxZoneRateID&","&TaxZoneID&");"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		end if
	
	set rs=nothing
	call closedb()
	response.redirect "viewTax.asp"
	End If
	
End If

pcv_IntTaxZoneRate=0
pcv_IntTaxZoneRateOrder=1
If request("mode")="MOD" Then
	idTaxZonesGroupID=Request("idTaxZonesGroupID")

	'//GET database information for this group
	call opendb()
	query="SELECT pcTaxZoneRates.pcTaxZoneRate_ID, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneRates.pcTaxZoneRate_Type, pcTaxZoneRates.pcTaxZoneRate_Order, pcTaxZoneRates.pcTaxZoneRate_Rate, pcTaxZoneRates.pcTaxZoneRate_ApplyToSH, pcTaxZoneRates.pcTaxZoneRate_Taxable, pcTaxZoneRates.pcTaxZoneRate_LocalZone, pcTaxZonesGroups.pcTaxZoneDesc_ID FROM pcTaxZoneRates INNER JOIN pcTaxZonesGroups ON pcTaxZoneRates.pcTaxZoneRate_ID = pcTaxZonesGroups.pcTaxZoneRate_ID WHERE (((pcTaxZoneRates.pcTaxZoneRate_ID)="&idTaxZonesGroupID&"));"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if NOT rs.eof then
		pcv_IntTaxZoneRateID=rs("pcTaxZoneRate_ID")
		pcv_StrTaxZoneRateName=rs("pcTaxZoneRate_Name")
		pcv_StrTaxZoneRateType=rs("pcTaxZoneRate_Type")
		pcv_IntTaxZoneRateOrder=rs("pcTaxZoneRate_Order")
		pcv_IntTaxZoneRate=rs("pcTaxZoneRate_Rate")
		pcv_IntTaxZoneRate=(pcv_IntTaxZoneRate*100)
		pcv_IntTaxZoneRateApplyToSH=rs("pcTaxZoneRate_ApplyToSH")
		pcv_IntTaxZoneRateTaxable=rs("pcTaxZoneRate_Taxable")
		pcv_IntTaxZoneRateLocalZone=rs("pcTaxZoneRate_LocalZone")
		pcv_TaxZoneDescID=rs("pcTaxZoneDesc_ID")
	else
		'//NO Zones!
		pcv_TaxZoneDescID=0
		pcv_IntTaxZoneRateID=idTaxZonesGroupID
	end if
	set rs=nothing
	call closedb()
End If

%>
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript"><!--
function newWin(file,window) {
		catWindow=open(file,window,'resizable=no,width=500,height=600,scrollbars=1');
		if (catWindow.opener == null) catWindow.opener = self;
}

function Form1_Validator(theForm)
{

	if (theForm.TaxZoneID.value == "")
		{
		alert("Please select a tax zone");
		theForm.TaxZoneID.focus();
		return (false);
	}

	if (theForm.TaxDesc.value == "")
		{
		alert("Please enter a tax description");
		theForm.TaxDesc.focus();
		return (false);
	}
	
	if (theForm.TaxRate.value == "")
		{
		alert("Please enter a tax rate");
		theForm.TaxRate.focus();
		return (false);
	}
return (true);
}
//--></script>
<form method="post" name="AddTaxZoneForm" action="AddTaxPerZone.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<% if taxErrMsg<>"" then %>
				<tr>
					<td colspan="2">
						<div class="pcCPmessage"><%=msg%></div>
					</td>
				</tr>
			<% end if %>
		<% 
		'If Zones exist, shown drop down 
		call opendb()
		query="SELECT pcTaxZoneDescriptions.pcTaxZoneDesc_ID, pcTaxZoneDescriptions.pcTaxZoneDesc FROM pcTaxZoneDescriptions;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rs.eof then %>								
			<tr>
			  <th colspan="2">Add/Edit Tax Rate by Zone</th>
			  </tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
			  <td colspan="2">
				<p>Create a new tax rate by selecting a predefined zone. You can add more zones by clicking on the &quot;Add/view Zones&quot; button.</p>
			  </td>
			</tr>
			<tr> 
				<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<% if pcv_TaxZoneDescID=0 AND request("mode")="MOD" then %>
				<tr> 
					<td colspan="2" align="center">
						<div class="pcCPmessage">This tax rule is not assigned to a zone. Assign a zone by filling in the form below or you can click on the back below the form to return to the previous screen and &quot;delete&quot; this tax rule. </div> 
						<input type="hidden" name="orphaned" value="1">
					</td>
				</tr>   
			<% end if %>
			<tr>
				<td valign="top"><div align="right">Zone:</div></td>
				<td><p>
				 <select name="TaxZoneID">
						<option value="">Select Zone</option>
						<% do until rs.eof
							intTaxZoneID=rs("pcTaxZoneDesc_ID")
							strTaxZoneName=rs("pcTaxZoneDesc")
							if pcv_TaxZoneDescID=intTaxZoneID then %>
							<option value="<%=intTaxZoneID%>" selected><%=strTaxZoneName%></option>
							<% else %>
							<option value="<%=intTaxZoneID%>"><%=strTaxZoneName%></option>
							<% end if %>
							<% rs.moveNext
						loop 
						%>
					</select>
					<img src="images/pc_required.gif" alt="required" width="9" height="9">&nbsp;&nbsp;&nbsp;<input type="button" name="Update" value="Add/View Zones" onClick="newWin('pcTaxViewZones.asp','window2')"></p></td>
			</tr>
			<%
			set rs=nothing
			call closedb() 
			%>
			<tr>
				<td width="20%" valign="top"><div align="right">Description:</div></td>
				<td width="80%"><p><input name="TaxDesc" type="text" id="TaxDesc" size="20" maxlength="50" value="<%=pcv_StrTaxZoneRateName%>"><img src="images/pc_required.gif" alt="required" width="9" height="9"></p>
				<p style="padding-top: 5px;">E.g. "GST". This is the description that is shown in the storefront to your customers. It will be displayed on the order summaries and all invoices.</p>
				</td>
				</tr>
			<tr>
				<td valign="top"><div align="right">Type:</div></td>
				<td>
					<p><input name="TaxType" type="text" id="TaxType" size="20" maxlength="50" value="<%=pcv_StrTaxZoneRateType%>">
					<img src="images/pc_required.gif" alt="required" width="9" height="9"></p>
					<p style="padding-top: 5px;">This is for your own reference for this tax rate and isn't displayed to the customer.</p>
				</td>
				</tr>
			<tr> 
				<td valign="top"><div align="right">Rate:</div></td>
				<td>
					<p><input name="TaxRate" size="6" value="<%=pcv_IntTaxZoneRate%>"> <img src="images/pc_required.gif" alt="required" width="9" height="9"> (5=5%)</p>
					<p style="padding-top: 5px;">This is the actual tax rate. If your tax rate is 7.25%, you would simply put 7.25 without the &quot;%&quot; symbol</p>
				</td>
			</tr>
			<tr>
				<td valign="top"><div align="right">Order:</div></td>
				<td><p><input name="TaxOrder" size="6" value="<%=pcv_IntTaxZoneRateOrder%>"></p>
					<p style="padding-top: 5px; padding-bottom: 10px;">If your zone has more then one tax rate, you can set the order in which you want the taxes applied on them.</p>
				</td>
				</tr>
			<tr>
				<td valign="top"><div align="right"><input type="checkbox" name="TaxApplySH" value="1" <% if pcv_IntTaxZoneRateApplyToSH="1" then%> checked<% end if %> class="clearBorder"></div></td>
				<td>
					<p>Apply tax to Shipping &amp; Handling</p>
					<p style="padding-top: 5px; padding-bottom: 10px;">If you want this rate to tax shipping and handling, you would check this box to enable that feature.</p>            
				</td>
			</tr>
			<tr>
        <td valign="top"><div align="right"><input type="checkbox" name="TaxLocalZone" value="1" <% if pcv_IntTaxZoneRateLocalZone="1" then%> checked<% end if %> class="clearBorder">
        </div></td>
			  <td>
					<p>Apply this tax only to customers in the same state/province as your store</p>
					<p style="padding-top: 5px; padding-bottom: 10px;">For example, this feature can be used to apply a local tax only to local sales.</p>
				</td>
			  </tr>
			<tr> 
				<td valign="top"><div align="right"><input type="checkbox" name="TaxTaxable" value="1" <% if pcv_IntTaxZoneRateTaxable="1" then%> checked<% end if %> class="clearBorder"></div></td>
				<td>
					<p>This tax is taxable</p>
					<p style="padding-top: 5px; padding-bottom: 10px;">If this tax is to be taxed by a second tax rule, check this box. Tax rules that are subject to additional taxation must be ordered so that they appear before the additional tax rule is applied. For example:  In Quebec the local tax is applied to the order total + federal tax (<a href="http://www.revenu.gouv.qc.ca/eng/entreprise/taxes/tvq_tps/info.asp" target="_blank">see example</a>), whereas in Ontario both taxes are calculated only on the order total. </p>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="pcCPspacer"></td>
			</tr>			
			<tr> 
				<td></td>
				<td> 
				    <% if request("mode")="MOD" then %>
				        <input type="hidden" name="UpdateID" value="<%=pcv_IntTaxZoneRateID%>">                
				        <input name="Submit" type="Submit" value="Update" class="submit2">&nbsp;
				        <% else %>                   
				        <input type="submit" name="Submit" value="Add" class="submit2">&nbsp;
				        <% end if %>
                       	<input name="button" type="button" onClick="location.href='manageTaxEptCust.asp?ZoneRateID=<%=idTaxZonesGroupID%>&mode=view '" value="View/Edit Customer Exemptions for This Tax">&nbsp;
				        <input type="button" name="back" value="Back" onClick="location.href='viewTax.asp'">
				  </td>
			</tr>
		<% else
			if ptaxCanada="1" then %>
				<tr>
					<td colspan="2">&nbsp;</td>
					</tr>
				<tr> 
					<td colspan="2">
						<p>You have not defined any zones for your store. Please create at least one zone before creating a tax rule.<br>
						<br>
					<input type="button" name="Update" value="Define A Zone" onClick="newWin('AdminTaxZones.asp','window2')" class="submit2"></p>	</td></tr>
			<% else %>
				<tr>
					<td colspan="2">&nbsp;</td>
					</tr>
				<tr> 
					<td colspan="2"><p>Activate Tax Zone Feature<br>
						<br>
							<input type="button" name="New" value="Activate Tax Zones" class="submit2" onClick="location.href='../includes/PageCreateTaxSettings.asp?ActivateZone=1'">
					</p>	</td>		</tr>
			<% end if %>
		<% end if %>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->