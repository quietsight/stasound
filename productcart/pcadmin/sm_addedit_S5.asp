<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials" 
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% 
Dim query, conntemp, rs

call openDB()

UP=session("sm_UP1")

if (UP="") OR (UP="0") OR (request("a")="new") then
	pcSaleID=request("id")

	if pcSaleID<>"" then
		query="SELECT pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round,pcSales_Param1,pcSales_Param2,pcSales_Tech FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			if request("a")="new" then
				session("sm_pcSaleID")=""
				session("sm_Param1")=""
				session("sm_Param2")=""
				session("sm_FilterTxt")=""
				session("sm_query")=""
				session("sm_UP1")=""
				session("sm_UP2")=""
				session("sm_UP3")=""
				session("sm_UP4")=""
				session("sm_UP5")=""
				session("sm_ChangeTxt")=""
				session("sm_ChangeName")=""
				session("sm_TechDetails")=""
				session("sm_TechDetails1")=""
				session("sm_SaleName")=""
				session("sm_SaleDesc")=""
				session("sm_SaleIcon")=""
				session("sm_loadedSale")=""
				session("sm_PrdCount")=0
			end if
			tmpSalesType=rs("pcSales_Type")
			if tmpSalesType="0" OR tmpSalesType="1" then
				session("sm_UP1")="1"
			else
				session("sm_UP1")="2"
			end if
			session("sm_UP2")=rs("pcSales_TargetPrice")
			session("sm_UP3")=rs("pcSales_Amount")
			if session("sm_UP1")="1" then
				session("sm_UP4")=rs("pcSales_Type")
			else
				session("sm_UP4")=rs("pcSales_Relative")
			end if
			session("sm_UP5")=rs("pcSales_Round")
			session("sm_Param1")=rs("pcSales_Param1")
			session("sm_Param2")=rs("pcSales_Param2")
			session("sm_TechDetails")=rs("pcSales_Tech")
			
			session("sm_query")="SELECT DISTINCT products.idProduct,products.sku,products.description FROM " & session("sm_Param1") & " WHERE " & session("sm_Param2")
			
			Set cn = Server.CreateObject("ADODB.Connection")
			Set cmd = Server.CreateObject("ADODB.Command")
			cn.Open scDSN
			Set cmd.ActiveConnection = cn
			cmd.CommandText = "uspGetPrdCount"
			cmd.CommandType = adCmdStoredProc

			cmd.Parameters.Refresh
			cmd.Parameters("@Param1") = session("sm_Param1")
			cmd.Parameters("@Param2") = session("sm_Param2")
			cmd.Execute
	
			tmpPrdCount=cmd.Parameters("@SMCOUNT")
			session("sm_PrdCount")=tmpPrdCount
			Set cmd=nothing
			set cn=nothing
			
			UP=session("sm_UP1")
			TempStr2=""
			tmpChangeName=""
	
			if UP="1" then
				priceSelect=session("sm_UP2")
				Select Case priceSelect
				Case "0": TempStr2="The Online Price"
				tmpChangeName="Online Price"
				Case "-1": TempStr2="The Wholesale Price"
				tmpChangeName="Wholesale Price"
				Case Else:
					tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect
					set rstemp4=connTemp.execute(tmpquery)
					TempStr2="The Price in Pricing Category: '" & rstemp4("pcCC_Name") & "'"
					tmpChangeName=rstemp4("pcCC_Name")
					set rstemp4=nothing
				End Select
			
				Select Case session("sm_UP4")
					Case "0": TempStr2=TempStr2 & " will be reduced by: " & session("sm_UP3") & "%"
					Case "1": TempStr2=TempStr2 & " will be reduced by: " & scCurSign & session("sm_UP3")
				End Select
				if session("sm_UP5")="1" then
					TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
				else
					if session("sm_UP5")="0" then
						TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
					end if
				end if
			end if
	
			if UP="2" then
				priceSelect1=session("sm_UP2")
				priceSelect2=session("sm_UP4")
				Select Case priceSelect1
				Case "0": TempStr2="Make the online price " & session("sm_UP3") & "% "
				tmpChangeName="Online Price"
				Case "-1": TempStr2="Make the wholesale price " & session("sm_UP3") & "% "
				tmpChangeName="Wholesale Price"
				Case Else:
					tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect1
					set rstemp4=connTemp.execute(tmpquery)
					TempStr2="Change Price in Pricing Category: '" & rstemp4("pcCC_Name") & "' " & session("sm_UP3") & "% "
					tmpChangeName=rstemp4("pcCC_Name")
					set rstemp4=nothing
				End Select
				Select Case priceSelect2
				Case "0": TempStr2=TempStr2 & "of the online price."
				Case "-2": TempStr2=TempStr2 & "of the list price."
				Case "-1": TempStr2=TempStr2 & "of the wholesale price."
				'Start SDBA
				Case "-3": TempStr2=TempStr2 & "of the product cost."
				'End SDBA
				Case Else:
					tmpquery="Select pcCC_Name FROM pcCustomerCategories WHERE idcustomerCategory=" & priceSelect2
					set rstemp4=connTemp.execute(tmpquery)
					TempStr2=TempStr2 & "of the price in pricing category: '" & rstemp4("pcCC_Name") & "'"
					set rstemp4=nothing
				End Select
				if session("sm_UP5")="1" then
					TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest integer."
				else
					if session("sm_UP5")="0" then
						TempStr2=TempStr2 & "<br>The updated price will be rounded to the nearest hundredth."
					end if
				end if
			end if

			session("sm_ChangeTxt")=TempStr2
			session("sm_ChangeName")=tmpChangeName
			
			session("sm_TechDetails1")="Products:<div style=""padding-top: 5px; font-weight:bold;"">The sale will affect " & session("sm_PrdCount") & " product(s) in your store.</div><div style=""padding-top: 5px;"" class=""pcSmallText"">If you just updated the product included in the sale, this number may be incorrect. it will be updated when you save the sale.</div>"
			session("sm_pcSaleID")=pcSaleID
		else
			pcSaleID=""
			session("sm_pcSaleID")=pcSaleID
		end if
		set rs=nothing
	else
		call closedb()
		response.redirect "sm_addedit_S1.asp"
		response.end
	end if
end if

if (request("Go")<>"") OR (request("Back")<>"") then
	session("sm_SaleName")=request("salename")
	session("sm_SaleDesc")=request("saledesc")
	session("sm_SaleIcon")=request("saleicon")
end if

if request("Back")<>"" then
	call closedb()
	response.redirect "sm_addedit_S4.asp"
	response.end
end if

if request("Go")<>"" then
	
	UpdatedSale=0
	
	tmpTargetPrice=session("sm_UP2")
	if session("sm_UP1")="1" then
		tmpSalesType=session("sm_UP4")
		tmpRelative=0
	else
		tmpSalesType=2
		tmpRelative=session("sm_UP4")
	end if
	tmpSalesAmount=session("sm_UP3")
	tmpRound=session("sm_UP5")
	
	dim dtTodaysDate
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	end if
	
	pcSaleID=session("sm_pcSaleID")
	
	if pcSaleID<>"" then
		query="SELECT pcSales_ID FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			set rs=nothing
			query="UPDATE pcSales SET pcSales_TargetPrice=" & tmpTargetPrice & ", pcSales_Type=" & tmpSalesType & ", pcSales_Relative=" & tmpRelative & ", pcSales_Amount=" & tmpSalesAmount & ", pcSales_Round=" & tmpRound & ", pcSales_Name='" & pcf_ReplaceCharacters(replace(session("sm_SaleName"),"'","''")) & "', pcSales_Desc='" & pcf_ReplaceCharacters(replace(session("sm_SaleDesc"),"'","''")) & "', pcSales_ImgURL='" & session("sm_SaleIcon") & "',pcSales_EditedDate='" & dtTodaysDate & "', pcSales_Param1='" & replace(session("sm_Param1"),"'","''") & "', pcSales_Param2='" & replace(session("sm_Param2"),"'","''") & "', pcSales_Tech='" & replace(session("sm_TechDetails"),"'","''") & "' WHERE pcSales_ID=" & pcSaleID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
			UpdatedSale=1
		end if
		set rs=nothing
	end if
	if UpdatedSale=0 then
		query="INSERT INTO pcSales (pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round,pcSales_Name,pcSales_Desc,pcSales_ImgURL,pcSales_CreatedDate,pcSales_Param1,pcSales_Param2,pcSales_Tech) VALUES (" & tmpTargetPrice & "," & tmpSalesType & "," & tmpRelative & "," & tmpSalesAmount & "," & tmpRound & ",'" & pcf_ReplaceCharacters(replace(session("sm_SaleName"),"'","''")) & "','" & pcf_ReplaceCharacters(replace(session("sm_SaleDesc"),"'","''")) & "','" & session("sm_SaleIcon") & "','" & dtTodaysDate & "','" & replace(session("sm_Param1"),"'","''") & "','" & replace(session("sm_Param2"),"'","''") & "','" & replace(session("sm_TechDetails"),"'","''") & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="SELECT TOP 1 pcSales_ID FROM pcSales ORDER BY pcSales_ID DESC;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			pcSaleID=rs("pcSales_ID")
			if IsNull(pcSaleID) then
				pcSaleID=0
			end if
		end if
		set rs=nothing
		query="UPDATE pcSales_Pending SET pcSales_ID=" & pcSaleID & " WHERE pcSales_ID=0;"
		set rs=connTemp.execute(query)
		set rs=nothing
		query="UPDATE pcSales SET pcSales_Param2='pcSales_Pending.pcSales_ID=" & pcSaleID & "' WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
	end if
	
	session("sm_pcSaleID")=""
	session("sm_Param1")=""
	session("sm_Param2")=""
	session("sm_UP1")=""
	session("sm_UP2")=""
	session("sm_UP3")=""
	session("sm_UP4")=""
	session("sm_UP5")=""
	session("sm_ChangeTxt")=""
	session("sm_ChangeName")=""
	session("sm_TechDetails")=""
	session("sm_TechDetails1")=""
	session("sm_SaleName")=""
	session("sm_SaleDesc")=""
	session("sm_SaleIcon")=""
	session("sm_loadedSale")=""
	session("sm_PrdCount")=0
	
	call closedb()
	
	if UpdatedSale=0 then
		response.redirect "sm_manage.asp?m=1&id=" & pcSaleID
	else
		response.redirect "sm_manage.asp?m=2&id=" & pcSaleID
	end if	

else

	pcSaleID=session("sm_pcSaleID")

	if (pcSaleID<>"") AND (session("sm_loadedSale")<>"1") then
		query="SELECT pcSales_Name,pcSales_Desc,pcSales_ImgURL FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			session("sm_SaleName")=rs("pcSales_Name")
			session("sm_SaleDesc")=rs("pcSales_Desc")
			session("sm_SaleIcon")=rs("pcSales_ImgURL")
			session("sm_loadedSale")="1"
		else
			if session("sm_loadedSale")="" then
				pcSaleID=""
				session("sm_pcSaleID")=pcSaleID
			end if
		end if
		set rs=nothing
	end if

end if

pageIcon="pcv4_icon_salesManager.png"
if pcSaleID="" then
pageTitle="Sales Manager - Create New Sale - Step 4: Save Sale"
else
pageTitle="Sales Manager - Edit Sale - Step 4: Save Sale"
end if


%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<script language="JavaScript">
<!--
	
function Form1_Validator(theForm)
{
	if (theForm.what.value=="0")
	{

	if (theForm.salename.value == "")
 	{
		    alert("Please enter 'Sale Name'");
			theForm.salename.focus();
		    return (false);
	}
	
 	if (theForm.saledesc.value == "")
 	{
		    alert("Please enter 'Sale Description'");
			theForm.saledesc.focus();
		    return (false);
	}
	if (theForm.saleicon.value == "")
 	{
		    alert("Please enter 'Sale Icon'");
			theForm.saleicon.focus();
		    return (false);
	}
	
	}
	
return (true);
}

function newWindow(file,window)
{
	msgWindow=open(file,window,'resizable=no,width=400,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
		
function chgWin1(file,window)
{
	msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}

//-->
</script>      

<form name="UpdateForm" id="UpdateForm" action="sm_addedit_S5.asp?a=post" method="post" onsubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td width="15%" nowrap align="right">                          
		Sale Name:<img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5">
	</td>
	<td width="85%"><input type="text" name="salename" id="salename" value="<%=session("sm_SaleName")%>" size="50"></td>
</tr>
<tr>
	<td valign="top" nowrap align="right">                          
		Sale Description:<img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5">
        <div style="padding-top:6px;"><input type="button" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_401")%>" onclick="newWindow('pop_HtmlEditor.asp?iform=UpdateForm&fi=saledesc','window2')"></div>
	</td>
	<td>
		<textarea name="saledesc" id="saledesc" cols="50" rows="10"><%=session("sm_SaleDesc")%></textarea>
	</td>
</tr>
<tr>
	<td width="15%" align="right">                          
		Sale Icon:<img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5">
	</td>
	<td>
		<input type="text" name="saleicon" id="saleicon" value="<%=session("sm_SaleIcon")%>" size="50"> <a href="javascript:;" onClick="chgWin1('../pc/imageDir.asp?ffid=saleicon&fid=UpdateForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;
		<a href="javascript:;" onClick="window.open('imageuploada_popup.asp?fi=UpdateForm.saleicon','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=460')"><img src="images/sortasc_blue.gif" alt="Upload your images"></a>&nbsp;
		<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%><a href="javascript:;" onClick="window.open('imageuploada_popup.asp?fi=UpdateForm.saleicon','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=460')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_303")%></a>.
	</td>
</tr>
<% if trim(session("sm_SaleIcon"))<>"" then %>
<tr>
	<td>&nbsp;</td>
	<td>Currently using this icon: <img src="../pc/catalog/<%=session("sm_SaleIcon")%>"></td>
</tr>
<% end if %>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td valign="top" nowrap>Products Manager:                         
	</td>
	<td valign="top"><a href="sm_addedit_S1.asp?a=rev&b=2">Click here</a> to view/edit list products included in this sale</td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr>
	<td valign="top" nowrap>                          
		Sales Recap:
	</td>
	<td valign="top"><%=session("sm_TechDetails")%><%=session("sm_TechDetails1")%></td>
</tr>
<tr>
	<td colspan="2"><hr></td>
</tr>
<tr align="center">
	<td colspan="2">
		<%if Clng(session("sm_PrdCount"))>0 then%>
		<input type="submit" name="Go" onclick="document.UpdateForm.what.value='0';" value="Save Sale" class="submit2">&nbsp;
		<%end if%>
		<input type="submit" name="Back" onclick="document.UpdateForm.what.value='1';" value="Edit Sale" class="ibtnGrey">
		<input type="hidden" name="what" value="0">
	</td>
</tr>
</table>
</form>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->