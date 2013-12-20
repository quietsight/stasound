<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/currencyformatinc.asp" --> 
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<% dim f, query, conntemp, rstemp, pIdProduct

call openDB()

'***********************************************************************
' START: ON POST BACK
'***********************************************************************
If request.form("Submit2")<>"" then
	pCnt=request.Form("oCnt")
	for i=1 to pCnt
		Uprice=request.form("price"&i)
		If Uprice="" then
		  Uprice="0"
		End If
		UWprice=request.form("Wprice"&i)
		If UWprice="" then
			UWprice="0"
		End If
		if scDecSign="," then
			Uprice=replacecomma(Uprice)
			UWprice=replacecomma(UWprice)
		else
			Uprice=replace(Uprice,",","")
			UWprice=replace(UWprice,",","")
		end if
		Uid=request.form("id"&i)
		USortOrder=request.form("sortOrder"&i)
		If USortOrder="" then
			USortOrder="0"
		End If
		OptActive=request.form("OptActive"&i)
		If OptActive="" then
			OptInActive="1"
		else
			OptInActive="0"	
		End If		
		query="UPDATE options_optionsGroups SET price="& Uprice &", Wprice="& UWprice &", SortOrder="& USortOrder &",InActive=" & OptInActive & " WHERE idoptoptgrp="& Uid
		set rstemp=conntemp.execute(query)				
	next
	
	
	pCnt2=request.Form("yCnt")
	for i=1 to pCnt2
		pRequired=request.Form("Required"&i)
		If pRequired<>"1" then
			pRequired="0"
		End If	
		catSort=request.Form("catSort"&i)
		If catSort="" then
			catSort="0"
		End If
		query="UPDATE pcProductsOptions SET pcProdOpt_Required="& pRequired &", pcProdOpt_Order="& catSort &" WHERE pcProdOpt_ID="& request.form("OptionGroupID"&i)
		set rstemp=conntemp.execute(query)			
	next
	set rstemp = nothing
	
	call updPrdEditedDate(request.form("idProduct"))
	
	call closeDB()
	'response.end
	response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode("You have successfully updated your product attributes.")&"&idProduct="& request.form("idProduct")
	
End If
'***********************************************************************
' END: ON POST BACK
'***********************************************************************



'***********************************************************************
' START: ON LOAD
'***********************************************************************
'// Form parameter 
pIdProduct=Request("idProduct")
if not validNum(pidproduct) then
   response.redirect "msg.asp?message=2"
end if

'// Get item details from db
query="SELECT idProduct, description FROM products WHERE products.idProduct=" & pIdProduct
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)
if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
    response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta.asp: "&Err.Description) 
end if

'// set data into local variables
pIdProduct=rstemp("idProduct")
pDescription=rstemp("description")

' SELECT DATA SET
' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
query = 		"SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
query = query & "FROM products "
query = query & "INNER JOIN ( "
query = query & "pcProductsOptions INNER JOIN ( "
query = query & "optionsgroups "
query = query & "INNER JOIN options_optionsGroups "
query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
query = query & "WHERE products.idProduct=" & pidProduct &" "
query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order, optionsGroups.OptionGroupDesc;"
set rs=server.createobject("adodb.recordset")
set rs=conntemp.execute(query)	
if err.number<>0 then
	'//Logs error to the database
	'call LogErrorToDatabase()
	'//clear any objects
	'set rs=nothing
	'//close any connections
	'call closedb()
	'//redirect to error page
	'response.redirect "techErr.asp?err="&pcStrCustRefID
	set rs=nothing
	call closeDb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in modPrdOpta.asp: "&Err.Description)
end if
	
'***********************************************************************
' END: ON LOAD
'***********************************************************************


'***********************************************************************
' START: MODE DELETE
'***********************************************************************
If Request("mode")="DEL" then

	idoptoptgrp=Request("id")
	
	'// Check the Option Group Number
	strSQL="SELECT idOptionGroup FROM options_optionsGroups WHERE idoptoptgrp="& idoptoptgrp &";"
	set rstemp=conntemp.execute(strSQL)	
	pIdOptionGroup = rstemp("idOptionGroup")
	
	'// Delete this option
	query="DELETE FROM options_optionsGroups WHERE idoptoptgrp="& idoptoptgrp
	set rstemp=conntemp.execute(query)
	
	'// Check if all options have been removed.
	strSQL="SELECT * FROM options_optionsGroups WHERE idproduct="& pIdProduct &" AND idoptionGroup="& pIdOptionGroup &";"
	set rstemp=conntemp.execute(strSQL)							
	if rstemp.eof then
		'// It is NOT related
		contgo=1
	end if	
	
	'// If all Options have been removed then delete the corrisponding record in pcProductOptions
	if contgo=1 then				
		strSQL="DELETE FROM pcProductsOptions WHERE idproduct="& pIdProduct &" AND idoptionGroup="& pIdOptionGroup &";"
		set rstemp=conntemp.execute(strSQL)
	end if	
	
	set rstemp=nothing
	
	call updPrdEditedDate(pIdProduct)
	
	call closedb()
	response.redirect "modPrdOpta.asp?s=1&msg="&Server.Urlencode("Your deletion was successful.")&"&idProduct="& pIdProduct
	response.end
End If
'***********************************************************************
' END: MODE DELETE
'***********************************************************************

pageTitle="Modify Product Options for: <strong>" & pDescription & "</strong>"
%>
<!--#include file="AdminHeader.asp"-->
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<form method="post" name="modifyProduct" action="modPrdOpta.asp" class="pcForms">                 
<input type="hidden" name="idproduct" value="<%=pidProduct%>">
		<table class="pcCPcontent">
			<tr>
				<td colspan="6"><div class="cpOtherLinks"><a href="modPrdOpta1.asp?idproduct=<%=pIdProduct%>">Add New Option Group</a> | <a href="FindProductType.asp?id=<%=pIdProduct%>">Edit Product</a> | <a href="../pc/viewPrd.asp?idProduct=<%=pIdProduct%>&adminPreview=1" target="_blank">Preview</a></div></td>				
			</tr>
			<%									
			' If we have data	
			If NOT rs.eof Then
				pcv_intOptionGroupCount = 0 '// keeps count of the number of options
				xOptionsCnt = 0 '// keeps count of the number of required options
				oCnt = 0
				yCnt = 0
				
				Do until rs.eof				
				yCnt = yCnt + 1	
					
					'// Get the Group Name
					pcv_strOptionGroupDesc=rs("OptionGroupDesc")
					'// Get the Group ID
					pcv_strOptionGroupID=rs("idOptionGroup")
					'// Is it required
					pcv_strOptionRequired=rs("pcProdOpt_Required")			
					'// Primary Key
					pcv_strProdOpt_ID=rs("pcProdOpt_ID")
					'// Sort Order
					strCatSort=rs("pcProdOpt_Order")
					'// Start: Do Option Count
					pcv_intOptionGroupCount = pcv_intOptionGroupCount + 1 
					'// End: Do Option Count
					
					'// Get the number of the Option Group
					pcv_strOptionGroupCount = pcv_intOptionGroupCount
					
					'// Start: Do Required Option Count
					if IsNull(pcv_strOptionRequired) OR pcv_strOptionRequired="" then
							pcv_strOptionRequired=0 '// not required // else it is "1"
					end if			
					if pcv_strOptionRequired=1 then							
						' Keep Tally
						xOptionsCnt = xOptionsCnt + 1
					end if
					'// End: Do Required Option Count
				
					'// Add Table Here
					%>
					<tr bgcolor="#e5e5e5"> 
						<td colspan="4"> 
							<span style="font-size:14px; font-weight: bold;"><%=pcv_strOptionGroupDesc%></span> 
						</td>
						<td colspan="2" align="right">
							Order: <input type="text" name="catSort<%=yCnt%>" size="1" maxlength="3" value="<%=strCatSort%>" style="text-align: right; font-size: 8pt; font-weight: bold; color: #000000; background-color: #99CCFF">
						</td>
					</tr>	
					<tr>
						<th nowrap>&nbsp;</th>	
						<th nowrap>Option Group - Option Attribute</th>							
						<th nowrap><input type="checkbox" name="A<%=yCnt%>" value="1" onclick="javascript:RunCheck<%=yCnt%>(this.checked);" class="clearBorder">&nbsp;Active</th>
						<th nowrap>Price</th>
						<th nowrap>Wholesale Price</th>
						<th nowrap>Order</th>
					</tr>	
                    <tr>
                        <td colspan="6" class="pcCPspacer"></td>
                    </tr>
					<%
					' SELECT DATA SET
					' TABLES: options_optionsGroups, options
					query = 		"SELECT options_optionsGroups.InActive, options_optionsGroups.price, options_optionsGroups.Wprice, "
					query = query & "options_optionsGroups.idoptoptgrp, options_optionsGroups.sortOrder, options.idoption, options.optiondescrip "
					query = query & "FROM options_optionsGroups "
					query = query & "INNER JOIN options "
					query = query & "ON options_optionsGroups.idOption = options.idOption "
					query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcv_strOptionGroupID &" "
					query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
					query = query & "ORDER BY options_optionsGroups.sortOrder, options.optiondescrip;"	
					set rs2=server.createobject("adodb.recordset")
					set rs2=conntemp.execute(query)	
					if err.number<>0 then
						'//Logs error to the database
						'call LogErrorToDatabase()
						'//clear any objects
						'set rs2=nothing
						'//close any connections
						'call closedb()
						'//redirect to error page
						'response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
				
					' If we have data
					if NOT rs2.eof then

						'// clean up the option group description
						if pcv_strOptionGroupDesc<>"" then
							pcv_strOptionGroupDesc=replace(pcv_strOptionGroupDesc,"""","&quot;")
						end if 							
											
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						' Start Loop
						'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
						tmp_start=0
						tmp_end=0
						do until rs2.eof			
						oCnt = oCnt + 1
						if tmp_start=0 then
							tmp_start=oCnt
						end if
						OptInActive=rs2("InActive") ' Is it active?
						if IsNull(OptInActive) OR OptInActive="" then
							OptInActive="0"
						end if
						
						dblOptPrice=rs2("price") '// Price
						dblOptWPrice=rs2("Wprice") '// WPrice
						intIdOptOptGrp=rs2("idoptoptgrp") '// The Id of the Option Group
						intIdOption=rs2("idoption") '// The Id of the Option
						strOptionDescrip=rs2("optiondescrip") '// A description of the Option
						pcv_strSortOrder=rs2("sortorder")
				
						'**************************************************************************************************
						' START: Dispay the Options
						'**************************************************************************************************
						%>
						<tr>                               
							<td width="6%">  
							<a href="modPrdOpta.asp?mode=DEL&id=<%=intIdOptOptGrp%>&idproduct=<%=pIdProduct%>">
								<img src="images/delete2.gif" width="23" height="18" border="0" alt="Remove">
							</a> 
							</td>
							<td width="60%"> 												
								<%=pcv_strOptionGroupDesc%> -  <b><%=strOptionDescrip%></b>
							</td>
							<td nowrap>
								<input name="OptActive<%=oCnt%>" type="checkbox" value="1" <%if (OptInActive<>"") and (OptInActive="1") then%><%else%>checked<%end if%> class="clearBorder">
							</td>
							<td nowrap>
								<%=scCurSign%>
								<input type="text" name="price<%=oCnt%>" value="<%=money(dblOptPrice)%>" size="6" maxlength="10">
							</td>
							<td nowrap>
								<%=scCurSign%> 
								<input type="text" name="Wprice<%=oCnt%>" value="<%=money(dblOptWPrice)%>" size="6" maxlength="10">
								<input type="hidden" name="id<%=oCnt%>" value="<%=intIdOptOptGrp%>">
							</td>
							<td nowrap>          
								<input name="sortOrder<%=oCnt%>" type="text" size="2" value="<%=pcv_strSortOrder%>">
							</td>
						</tr>										
						<% 
						'**************************************************************************************************
						' END: Dispay the Options
						'**************************************************************************************************
					rs2.movenext 
					loop
					if tmp_start>0 then
						tmp_end=oCnt
					end if
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					' END Loop
					'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
					set rs2=nothing	
				end if
				%>				
						
				<tr>                               
					<td colspan="6" nowrap style="border-top: 1px solid #CCC;">
					<% if pcv_strOptionRequired=1 then %>
						<input type="checkbox" name="Required<%=yCnt%>" value="1" checked class="clearBorder">
					<% else %>
						<input type="checkbox" name="Required<%=yCnt%>" value="1" class="clearBorder">
					<% end If %> <b>Required Option
					&nbsp;<%=chr(124)%>&nbsp;						
					<a href="modPrdOpta3.asp?idproduct=<%=pidProduct%>&IdOptionGroup=<%=pcv_strOptionGroupID%>">Add More Attributes</a></b>
					<input type="hidden" name="OptionGroupID<%=yCnt%>" value="<%=pcv_strProdOpt_ID%>">
					</td>
				</tr>                            
				<tr> 								  
					<td colspan="6">
						<script>
							function RunCheck<%=yCnt%>(tstatus)
							{
								CheckUncheckBoxes(<%=tmp_start%>,<%=tmp_end%>,tstatus);
							}
						</script>
					&nbsp;</td>
				</tr>					
				<%
				rs.movenext
			Loop			
			set rs=nothing
			%>
			<tr> 
				<td colspan="6"><hr></td>
			</tr>
			<tr> 
				<td colspan="6" align="center">
					<input type="submit" name="Submit2" value="Update" class="submit2">
					&nbsp;
					<input type="button" name="Clone" value="Copy to other products" onClick="location.href='ApplyOptionsMulti2.asp?action=add&prdlist=<%=pIdProduct%>'">
					&nbsp;
					<input type="button" name="Button" value="Manage Options" onClick="location.href='manageOptions.asp'">
					&nbsp;
					<input type="button" name="Button" value="Locate Another Product" onClick="location.href='LocateProducts.asp?cptype=0'">
					<script>
						function CheckUncheckBoxes(tmp_start,tmp_end,tstatus)
						{
							for (var i=tmp_start;i<=tmp_end;i++)
							eval("document.modifyProduct.OptActive" + i).checked=tstatus;
						}
					</script>
				</td>
			</tr>	
				
			<%	
			Else
			%>	
														
				<tr> 								  
					<td colspan="6">No option group has been added to this product.</td>
				</tr>
			<tr> 
				<td colspan="6"><hr></td>
			</tr>
				<tr> 
					<td colspan="6" align="center">
						<input type="button" name="Button" value="Add New Option Group" onClick="location.href='modPrdOpta1.asp?idproduct=<%=pIdProduct%>'" class="submit2">&nbsp;
						<input type="button" name="Button" value="Manage Options" onClick="location.href='manageOptions.asp'">&nbsp;
						<input type="button" name="Button" value="Locate Another Product" onClick="location.href='LocateProducts.asp?cptype=0'">
					</td>
				</tr>	
				
			<%	
			End If
			%>												
		</table>
	<input type="hidden" name="oCnt" value="<%=oCnt%>">
	<input type="hidden" name="yCnt" value="<%=yCnt%>">
</form>
<!--#include file="AdminFooter.asp"-->