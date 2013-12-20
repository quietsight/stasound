<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>
<% pageTitle="Manage Customers: Pricing Categories" %>
<% section="mngAcc" %>
<%PmAdmin=19%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="adminHeader.asp"-->

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<% 
Dim pcCC_EXIST 'Flag - if there are current category
Dim pcCC_INPROCESS 'Flag - if there is a current category being added with errors
Dim strpcCC_Name, strpcCC_Description, intpcCC_WholesalePriv, strpcCC_CategoryType, dblpcCC_ATB_Percentage, strpcCC_ATB_Off, strpcCC_NFSO 'for field variables
dim conntemp, query, rs 'database and query variables
dim pcArray, intCount

pcCC_INPROCESS=0
pcCC_EXIST=0

'Check mode
'mode=1 - Add
'mode=2 - Modify
'mode=3 - Delete
dim intMode, intModeID
intMode=getuserinput(request("mode"),1)
intModeID=0

if NOT isNumeric(intMode) or intMode="" then
	intMode=0
end if

if intMode<>0 then
	intModeID=getuserinput(request("id"),6)
end if
if NOT isNumeric(intModeID) or intModeID="" then
	intModeID=0
end if

call opendb()

'// DELETE - Start
If intMode=3 then
	'del
	query="DELETE FROM pcCustomerCategories WHERE idcustomerCategory="&intModeID&";"
	SET rs=server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	
	query="UPDATE customers SET customerType=0, idcustomerCategory=0 WHERE idcustomerCategory="&intModeID&";"
	SET rs=server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)

	SET rs=nothing
	call closedb()
	response.redirect("AdminCustomerCategory.asp?s=1&msg=" & Server.Urlencode("Pricing category deleted successfully."))
end if
'// DELETE - End

'// MODIFY - Get Info - Start
If request.Form("submitflag")="" AND intMode<>0 and intModeID<>0 Then
	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_off, pcCC_NFSoverride FROM pcCustomerCategories WHERE idcustomerCategory="&intModeID&";"
	SET rs=server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	strpcCC_Name=rs("pcCC_Name")
	strpcCC_Description=rs("pcCC_Description")
	intpcCC_WholesalePriv=rs("pcCC_WholesalePriv")
	if isNULL(intpcCC_WholesalePriv) OR intpcCC_WholesalePriv="" then
		intpcCC_WholesalePriv=0
	end if
	strpcCC_CategoryType=rs("pcCC_CategoryType")
	dblpcCC_ATB_Percentage=rs("pcCC_ATB_Percentage")
	strpcCC_ATB_Off=rs("pcCC_ATB_OFF")
	strpcCC_NFSO=rs("pcCC_NFSoverride")
'// MODIFY - End
Else
'// DISPLAY - Start
	'check for existing customer categories
	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_off, pcCC_NFSoverride FROM pcCustomerCategories ORDER BY pcCC_Name ASC;"
	SET rs=server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	if NOT rs.eof then
		pcCC_EXIST=1
		'getrows
		pcArray=rs.getRows()
	end if
End If

SET rs=nothing
call closedb()

if request.Form("submitflag")="1" then
	pcCC_INPROCESS=1 
	strpcCC_Name=getuserinput(request.Form("pcCC_Name"),250)
	strpcCC_Description=getuserinput(request.Form("pcCC_Description"),250)
	intpcCC_WholesalePriv=getuserinput(request.Form("pcCC_WholesalePriv"),4)
	if intpcCC_WholesalePriv="" then
		intpcCC_WholesalePriv=0
	end if
	strpcCC_CategoryType=getuserinput(request.Form("pcCC_CategoryType"),4)
	dblpcCC_ATB_Percentage=getuserinput(request.Form("pcCC_ATB_Percentage"),6)
	dblpcCC_ATB_Percentage=replace(dblpcCC_ATB_Percentage,"%","")
	if dblpcCC_ATB_Percentage="" or not validNum(dblpcCC_ATB_Percentage) then
		dblpcCC_ATB_Percentage=0
	end if
	strpcCC_ATB_Off=getuserinput(request.Form("pcCC_ATB_OFF"),10)
	if strpcCC_ATB_Off="" then
		strpcCC_ATB_Off=0
	end if
	GlobalPercentOff=getuserinput(request.Form("GlobalPercentOff"),4)
	GlobalPercentOff=replace(GlobalPercentOff,"%","")
	if NOT validNum(GlobalPercentOff) OR GlobalPercentOff="" then
		GlobalPercentOff=0
	end if
	strpcCC_NFSO=getuserinput(request.Form("NFSO"),1)
	if NOT validNum(strpcCC_NFSO) OR strpcCC_NFSO="" then
		strpcCC_NFSO=0
	end if
	call opendb()
	
	if intMode="2" then
	'// MODIFY - Start
		query="UPDATE pcCustomerCategories SET pcCC_Name='"&strpcCC_Name&"', pcCC_Description='"&strpcCC_Description&"', pcCC_WholesalePriv="&intpcCC_WholesalePriv&" ,pcCC_CategoryType='"& strpcCC_CategoryType&"', pcCC_ATB_Percentage="&dblpcCC_ATB_Percentage&", pcCC_ATB_off='"&strpcCC_ATB_off&"', pcCC_NFSoverride="&strpcCC_NFSO&" WHERE idcustomerCategory="&intModeID&";"
		SET rs=server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		'update all customers that are assigned to this category
		query="UPDATE customers SET customerType="&intpcCC_WholesalePriv&" WHERE idcustomerCategory="&intModeID&";"
		SET rs=server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		intMode=0
		intModeID=0
		pcCC_INPROCESS=0
		pcCC_EXIST=1
	'// MODIFY - End
	else
	'// ADD NEW - Start
		'ENSURE THERE ISN'T ANOTHER CUSTOMER CATEGORY WITH THE SAME NAME
		query="SELECT idcustomerCategory FROM pcCustomerCategories WHERE pcCC_NAME='"&strpcCC_Name&"';"
		SET rs=server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		if rs.eof then
			query="INSERT INTO pcCustomerCategories (pcCC_Name, pcCC_Description, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_off, pcCC_NFSoverride) VALUES ('"&strpcCC_Name&"', '"&strpcCC_Description&"', "&intpcCC_WholesalePriv&", '"& strpcCC_CategoryType&"', "&dblpcCC_ATB_Percentage&", '"&strpcCC_ATB_off&"',"&strpcCC_NFSO&");"
			SET rsCTObj=server.CreateObject("ADODB.RecordSet")
			SET rsCTObj=conntemp.execute(query)
			if err.number<>0 then
				response.write "Error adding new pricing category to database. The error is: " & err.description
				response.end
			else
				query="SELECT idcustomerCategory FROM pcCustomerCategories WHERE pcCC_NAME='"&strpcCC_Name&"';"
				SET rsCTObj=server.CreateObject("ADODB.RecordSet")
				SET rsCTObj=conntemp.execute(query)
				if not rsCTObj.eof then
					intModeID=rsCTObj("idcustomerCategory")
				end if
				set rsCTObj=nothing
			end if
			set rsCTObj=nothing
			
		else
			set rs=nothing
			call closeDb()
			response.redirect("AdminCustomerCategory.asp?mode=1&msg=" & Server.Urlencode("The Pricing Category could not be added because there is already one with the same name."))
			response.End()
		end if
		
		SET rs=nothing
		call closedb()
		if request("Submit1")<>"" then
			response.redirect "editCustCategories.asp?OrderBy=&SortOrd=&CCID=" & intModeID & "&mode=edit"
		else
			response.redirect("AdminCustomerCategory.asp?s=1&msg=" & Server.Urlencode("New Pricing Category added successfully."))
		end if
		response.End()
	'// ADD NEW - End
	end if
	
	'see if Global Product Price Change is needed

	if GlobalPercentOff>0 then
		'Get idCustomerCategory from DATABASE
		query="SELECT idcustomerCategory FROM pcCustomerCategories WHERE pcCC_NAME='"&strpcCC_Name&"';"
		SET rs=server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		if NOT rs.eof then
			intIdCustomerCategory=rs("idcustomerCategory")
			query="SELECT idProduct, price FROM products;"
			SET rsPrdObj=Server.CreateObject("ADODB.RecordSet")
			SET rsPrdObj=conntemp.execute(query)
			if NOT rsPrdObj.eof then
				pcPrdArray=rsPrdObj.GetRows()
				intPrdCount=ubound(pcPrdArray,2)
				For pcDoCnt=0 to intPrdCount
					tIdProduct=pcPrdArray(0,pcDoCnt)
					tPrice=pcPrdArray(1,pcDoCnt)
					intChangePrice=tPrice-(tPrice*(GlobalPercentOff/100))
					'insert into database
					query="SELECT idCC_Price FROM pcCC_Pricing WHERE idcustomerCategory="&intIdCustomerCategory&" AND idProduct="&tIdProduct&";"
					SET rsPBPObj=Server.CreateObject("ADODB.RecordSet")
					SET rsPBPObj=conntemp.execute(query)
					if rsPBPObj.eof then
						query="INSERT INTO pcCC_Pricing (idcustomerCategory, idProduct, pcCC_Price) VALUES ("&intIdcustomerCategory&","&tIdProduct&","&intChangePrice&");"
					else
						intIdCC_Price=rsPBPObj("idCC_Price")
						query="UPDATE pcCC_Pricing SET pcCC_Price="&intChangePrice&" WHERE idCC_Price="&intIdCC_Price&";"
					end if
					SET rsIObj=Server.CreateObject("ADODB.RecordSet")
					SET rsIObj=conntemp.execute(query)
		
					SET rsIObj=nothing
					SET rsPBPObj=nothing
				Next
			end if
			SET rsPrdObj=nothing
		end if
		SET rs=nothing
	end if

	'check for existing customer categories
	query="SELECT idcustomerCategory, pcCC_Name, pcCC_Description, pcCC_WholesalePriv, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_off, pcCC_NFSoverride FROM pcCustomerCategories ORDER BY pcCC_Name ASC;"
	SET rs=server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
	'getrows
	pcArray=rs.getRows()

	SET rs=nothing
	
	call closedb()
end if

if intMode=1 AND intModeID=0 then
	pcCC_INPROCESS=1
end if

if pcCC_EXIST=0 OR pcCC_INPROCESS=1 OR intMode=2 then %>
<form name="CustTypeForm" action="AdminCustomerCategory.asp" method="post" class="pcForms">
  <input type="hidden" name="submitflag" value="1">
  <input type="hidden" name="mode" value="<%=intMode%>">
  <input type="hidden" name="id" value="<%=intModeID%>">
	<% if intMode<2 then
		strPageAction="Add Customer Category Type"
	else
		strPageAction="Modify Customer Category Type"
	end if %>
  	<table class="pcCPcontent">
		<tr>
		  <td colspan="2">
          <div style="float: right; margin: 5px 0 0 0; background-color: #FFC; padding: 5px;">Feature limitation: Pricing categories do not affect option prices.</div>
          <h2><%=strPageAction%></h2></td>
		</tr>
		<tr> 
			<td width="20%" align="right">Name:</td>
			<td width="80%"><input name="pcCC_Name" type="text" size="50" maxlength="250" value="<%=strpcCC_Name%>"></td>
		</tr>
		<tr>
		  <td align="right">Description:</td>
		  <td><input name="pcCC_Description" type="text" size="50" maxlength="250" value="<%=strpcCC_Description%>"></td>
		  </tr>
		<tr>
			<td align="right"></td>
			<td>Apply Wholesale Customer Privileges: <input type="checkbox" name="pcCC_WholesalePriv" value="1" <% if intpcCC_WholesalePriv=1 then %>checked<% end if %> class="clearBorder">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=305')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td align="right"></td>
			<td>Override &quot;Not For Sale&quot; setting: <input type="checkbox" name="NFSO" value="1" <% if strpcCC_NFSO=1 then %>checked<% end if %> class="clearBorder">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=318')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>Apply custom pricing: <input name="pcCC_CategoryType" type="radio" value="ATB" <% if strpcCC_CategoryType="ATB" then %>checked<% end if %> class="clearBorder"> Across the board &nbsp;<input name="pcCC_CategoryType" type="radio" value="PBP" <% if strpcCC_CategoryType="PBP" then %>checked<% end if %> class="clearBorder"> On a product by product basis &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=306')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
		  <td colspan="2" bgcolor="#e5e5e5">If you chose &quot;across the board&quot;, specify these additional settings</td>
		  </tr>
		<tr>
		  <td align="right"><input name="pcCC_ATB_Percentage" type="text" size="10" maxlength="10" value="<%=dblpcCC_ATB_Percentage%>"></td>
			 <td>Percentage discount (e.g. 10% off: enter 10)</td>
		 </tr>
		<tr>
		  <td align="right"><input name="pcCC_ATB_Off" type="radio" value="Wholesale" <% if strpcCC_ATB_Off="Wholesale" then %>checked<% end if %> class="clearBorder"></td>
		   <td>Calculated off the wholesale price.</td>
		  </tr>
		<tr>
		  <td align="right"><input name="pcCC_ATB_Off" type="radio" value="Retail" <% if strpcCC_ATB_Off="Retail" then %>checked<% end if %> class="clearBorder"></td>
			<td>Calculated off the retail price.</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
		  <td colspan="2" bgcolor="#e5e5e5">If you chose &quot;product by product&quot;, specify these additional settings</td>
		</tr>
		<% if intMode>1 then %>
			<tr>
				<td colspan="2">
					<div class="pcCPnotes" style="margin: 10px;">WARNING: If you have already assigned pricing on a Product by Product basis for this customer category, then this setting will overwrite all your previous pricing. Please set to &quot;0&quot; if you do not want to overwrite current prices.</div>
				</td>
			</tr>
		<% end if %>
		<tr>
		  <td colspan="2">Assign price based on <input name="GlobalPercentOff" type="text" value="0" size="4"> % off the current online price.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=307')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr>
		  <td colspan="2" align="center">
				<input type="submit" name="Submit" value="Save" class="submit2">
				<%if intMode="1" then%>
				&nbsp;<input type="submit" name="Submit1" value="Assign/Remove Customers">
				<%else%>
				&nbsp;<input type="button" value="Assign/Remove Customers" onClick="location.href='EditCustCategories.asp?CCID=<%=intModeID%>&move=view'">
				<%end if%>
				&nbsp;<input type="button" value="Reset Prices" onClick="location.href='resetPricingCatPrices.asp'">
				&nbsp;<input type="button" value="Other Pricing Categories" onClick="location.href='AdminCustomerCategory.asp'">
			</td>
		</tr>
	</table>
</form>
<% End If

'first time on page, show only if there are existing categories
if pcCC_EXIST=1 and pcCC_INPROCESS=0 then %>
	<!--#include file="pcCharts.asp"-->
	<table class="pcCPcontent">
		<tr>
			<td>
				<div id="chartPricingCats" style="height:330px; "></div>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	</table>
	<%
	call pcs_PricingCatsChart("chartPricingCats")
	%>
	<table class="pcCPcontent">
		<tr> 
			<td colspan="3">
                <div class="cpOtherLinks"><a href="AdminCustomerCategory.asp?mode=1">Add New Pricing Category</a> | <a href="resetPricingCatPrices.asp">Reset Prices for a Pricing Category</a> | <a href="http://wiki.earlyimpact.com/productcart/customer-prices" target="_blank">Help</a></div>
		  </td>
		</tr>

		<%
		intCount=ubound(pcArray,2)
		For pcCT=0 to intCount
		%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
        	<td width="30%" align="left" valign="top"><a href="AdminCustomerCategory.asp?mode=2&id=<%=pcArray(0,pcCT)%>"><strong><%=pcArray(1,pcCT)%></strong></a></td>
			<td width="65%" align="left" valign="top"><%=pcArray(2,pcCT)%></td>
			<td width="5%" align="right" valign="top" nowrap class="cpLinksList">
				<a href="editCustCategories.asp?CCID=<%=pcArray(0,pcCT)%>&mode=view"><img src="images/pcIconList.jpg" alt="View / Add / Remove" title="List, add, and remove customers in, to, and from the selected pricing category." width="12" height="12" border="0"></a> 
                <a href="AdminCustomerCategory.asp?mode=2&id=<%=pcArray(0,pcCT)%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit settings" title="Edit settings for the selected pricing category"></a> 
                <a href="javascript:if (confirm('You are about to remove this pricing category. Customers assigned to the category will lose the corresponding privileges. Are you sure you want to complete this action?')) location='AdminCustomerCategory.asp?mode=3&id=<%=pcArray(0,pcCT)%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" title="Delete Pricing Category" alt="Delete"></a>
			</td>
		</tr>
		<% next %>
	</table>
<% End If %>
<br /><br />
<!--#include file="adminFooter.asp"-->