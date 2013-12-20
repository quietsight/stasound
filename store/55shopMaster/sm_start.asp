<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
IF request("a")="start" THEN
	pageTitle="Sales Manager - Start a Sale - Preview"
ELSE
	pageTitle="Sales Manager - Start a Sale"
END IF
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="sm_check.asp"-->
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

call opendb()

IF (request("a")="start") OR (request("a")="run") THEN
	pcSaleID=request("id")
	if pcSaleID="" then
		call closedb()
		response.redirect "sm_manage.asp"
	else
		if Not (IsNumeric(pcSaleID)) then
			call closedb()
			response.redirect "sm_manage.asp"
		end if
	end if
END IF

query="SELECT pcSales_ID,pcSales_Name FROM pcSales WHERE pcSales_Removed=0 AND (pcSales_ID NOT IN (SELECT pcSales_ID FROM pcSales_Completed)) ORDER BY pcSales_Name ASC;"
set rs=connTemp.execute(query)
if rs.eof then
	set rs=nothing
	call closedb()
	response.redirect "sm_addedit_S1.asp?a=new"
else
	pcArr=rs.getRows()
	intCount=ubound(pcArr,2)
	set rs=nothing
end if

%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%IF request("a")="start" THEN
	query="SELECT pcSales_Param1,pcSales_Param2,pcSales_Tech,pcSales_Name FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "sm_manage.asp"
	else
		tmpParam1=rs("pcSales_Param1")
		tmpParam2=" ((products.pcSC_ID IS NULL) OR (products.pcSC_ID=0)) AND " & rs("pcSales_Param2")
		tmpParam2A=rs("pcSales_Param2")
		tmpSaleDetails=rs("pcSales_Tech")
		tmpSaleName=rs("pcSales_Name")
		set rs=nothing
		
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspGetPrdCount"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = tmpParam1
		cmd.Parameters("@Param2") = tmpParam2
		cmd.Execute
	
		tmpPrdCount=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
		
		tmpParam2=tmpParam2A
		
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspGetPrdCount"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = tmpParam1
		cmd.Parameters("@Param2") = tmpParam2
		cmd.Execute
	
		tmpPrdCountORG=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
	%>
	<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
	<td>
		Sale name: <b><%=tmpSaleName%></b><br><br>
		<%if tmpPrdCount>0 then%>
			The sale will affect <b><%=tmpPrdCount%></b> product(s) in your store.<br><br>
			<a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcSaleID%>&b=4">Click here</a> to view/edit list products included in this sale
			<br><br>
		<%end if%>
		<%=tmpSaleDetails%>
	</td>
	</tr>
	<%if (clng(tmpPrdCountORG)>clng(tmpPrdCount)) AND (tmpPrdCount>0) then%>
	<tr>
		<td>
			<div class="pcCPmessageInfo">
				Not all selected products are eligible to be placed on sale as they are already part of a different, live sale.<br>
				Of the <strong><%=tmpPrdCountORG%></strong> product(s) that you had selected, only <strong><%=tmpPrdCount%></strong> can be placed on sale at this time.
			</div>
		</td>
	</tr>
	<%else%>
	<%if tmpPrdCount=0 then%>
	<tr>
		<td>
			<div class="pcCPmessage">
				There's a problem. The sale - &quot;as is&quot; - does not affect any product(s) in your store. They might already be 'on sale', or may have been removed from your store. You need to edit the sale and change the product selection.
			</div>
		</td>
	</tr>
	<%end if%>
	<%end if%>
	<tr>
		<td class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td>
			<%if tmpPrdCount>0 then%>
				<input type="button" name="Run" value="Start Sale Now" onclick="javascript:if (confirm('You are about to start this sale. The price changes you have configured will now be applied to the selected products, and will immediately be live in your storefront. Do you want to continue?')) location='sm_start.asp?a=run&id=<%=pcSaleID%>';" class="submit2">
			<%else%>
				<input type="button" name="Back" value="View/Edit Pending Sales" onclick="location='sm_manage.asp';" class="submit2">
			<%end if%>
		</td>
	</tr>
	</table>	
	<%end if
ELSE
IF request("a")="run" THEN

	query="SELECT pcSales_TargetPrice,pcSales_Type,pcSales_Relative,pcSales_Amount,pcSales_Round,pcSales_Name,pcSales_Desc,pcSales_ImgURL,pcSales_Param1,pcSales_Param2,pcSales_Tech FROM pcSales WHERE pcSales_ID=" & pcSaleID & ";"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmpTargetPrice=rs("pcSales_TargetPrice")
		tmpSaleType=rs("pcSales_Type")
		tmpRelative=rs("pcSales_Relative")
		tmpSaleAmount=rs("pcSales_Amount")
		tmpRound=rs("pcSales_Round")
		tmpSaleName=rs("pcSales_Name")
		tmpSaleDesc=rs("pcSales_Desc")
		tmpImgURL=rs("pcSales_ImgURL")
		tmpParam1=rs("pcSales_Param1")
		tmpParam2=" ((products.pcSC_ID IS NULL) OR (products.pcSC_ID=0)) AND " & rs("pcSales_Param2")
		tmpSaleDetails=rs("pcSales_Tech")
	end if
	set rs=nothing
	
	'// Save all prices to pcCC_Pricing before starting the backup 
	if CLng(tmpTargetPrice)>0 then
		query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & tmpTargetPrice
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		if NOT rs.eof then 
			intIdcustomerCategory=rs("idcustomerCategory")
			strpcCC_Name=rs("pcCC_Name")
			strpcCC_CategoryType=rs("pcCC_CategoryType")
			intpcCC_ATBPercentage=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			set rs=nothing
			if intpcCC_ATB_Off="Retail" then
				intpcCC_ATBPercentOff=0
			else
				intpcCC_ATBPercentOff=1
			end if
						
			if strpcCC_CategoryType="ATB" then
				Set cn = Server.CreateObject("ADODB.Connection")
				Set cmd = Server.CreateObject("ADODB.Command")
				cn.Open scDSN
				Set cmd.ActiveConnection = cn
				cmd.CommandText = "uspAddCatPrices"
				cmd.CommandType = adCmdStoredProc

				cmd.Parameters.Refresh
				cmd.Parameters("@Param1") = tmpParam1
				cmd.Parameters("@Param2") = tmpParam2
				cmd.Parameters("@IDCat") = intIdcustomerCategory
				cmd.Parameters("@CAmount") = Round(100-cdbl(intpcCC_ATBPercentage),2)*0.01
				cmd.Parameters("@CType") = intpcCC_ATBPercentOff
				cmd.Execute
	
				tmpPrdCount=cmd.Parameters("@SMCOUNT")

				Set cmd=nothing
				set cn=nothing
				
			end if
		end if
	end if
	
	'// Save all prices to pcCC_Pricing before starting the backup
	if CLng(tmpRelative)>0 then
		query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType, pcCC_ATB_Percentage, pcCC_ATB_Off FROM pcCustomerCategories WHERE idcustomerCategory=" & tmpRelative
		SET rs=Server.CreateObject("ADODB.RecordSet")
		SET rs=conntemp.execute(query)
		if NOT rs.eof then 
			intIdcustomerCategory=rs("idcustomerCategory")
			strpcCC_Name=rs("pcCC_Name")
			strpcCC_CategoryType=rs("pcCC_CategoryType")
			intpcCC_ATBPercentage=rs("pcCC_ATB_Percentage")
			intpcCC_ATB_Off=rs("pcCC_ATB_Off")
			set rs=nothing
			if intpcCC_ATB_Off="Retail" then
				intpcCC_ATBPercentOff=0
			else
				intpcCC_ATBPercentOff=1
			end if
						
			if strpcCC_CategoryType="ATB" then
				Set cn = Server.CreateObject("ADODB.Connection")
				Set cmd = Server.CreateObject("ADODB.Command")
				cn.Open scDSN
				Set cmd.ActiveConnection = cn
				cmd.CommandText = "uspAddCatPrices"
				cmd.CommandType = adCmdStoredProc

				cmd.Parameters.Refresh
				cmd.Parameters("@Param1") = tmpParam1
				cmd.Parameters("@Param2") = tmpParam2
				cmd.Parameters("@IDCat") = intIdcustomerCategory
				cmd.Parameters("@CAmount") = Round(100-cdbl(intpcCC_ATBPercentage),2)*0.01
				cmd.Parameters("@CType") = intpcCC_ATBPercentOff
				cmd.Execute
	
				tmpPrdCount=cmd.Parameters("@SMCOUNT")

				Set cmd=nothing
				set cn=nothing
				
			end if
		end if
	
	end if
	
	'//Create New Sale Record
	
	dim dtTodaysDate
	dtTodaysDate=Date()
	if SQL_Format="1" then
		dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	else
		dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
	end if
	
	query="INSERT INTO pcSales_Completed (pcSales_ID,pcSC_Status,pcSC_StartedDate,pcSC_SaveName,pcSC_SaveDesc,pcSC_SaveTech,pcSC_SaveIcon) VALUES (" & pcSaleID & ",1,'" & dtTodaysDate & "','" & replace(tmpSaleName,"'","''") & "','" & replace(tmpSaleDesc,"'","''") & "','" & replace(tmpSaleDetails,"'","''") & "','" & tmpImgURL & "');"
	set rs=connTemp.execute(query)
	set rs=nothing
	query="SELECT TOP 1 pcSC_ID FROM pcSales_Completed WHERE pcSales_ID=" & pcSaleID & " ORDER BY pcSC_ID DESC;"
	set rs=connTemp.execute(query)
	pcSCID=0
	if not rs.eof then
		pcSCID=rs("pcSC_ID")
	end if
	set rs=nothing%>
	
	<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Processing Status</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
	<td>
	<%
	
		'Count affected products
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspGetPrdCount"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = tmpParam1
		cmd.Parameters("@Param2") = tmpParam2
		cmd.Execute
	
		tmpPrdCount=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
		
	%>
	<img src="images/pc_checkmark_sm.gif" width="16" height="16" alt="Task performed"> Preparing to update <b><%=tmpPrdCount%></b> product(s)...<br><br>
	<%
	
		'Back Up products
		dtTodaysDate=Date()
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		end if
		
		query="UPDATE pcSales_Completed SET pcSC_BUStartedDate='" & dtTodaysDate & "' WHERE pcSC_ID=" & pcSCID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspBackUpPrices"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = tmpParam1
		cmd.Parameters("@Param2") = tmpParam2
		cmd.Parameters("@SCID") = pcSCID
		cmd.Parameters("@SalesID") = pcSaleID
		cmd.Parameters("@TPrice") = tmpTargetPrice
		cmd.Execute
	
		tmpBUPrdCount=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
		
		tmpSCStatus=1
		if Clng(tmpBUPrdCount)<>Clng(tmpPrdCount) then
			tmpSCStatus=3
		end if
		
		dtTodaysDate=Date()
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
		end if
		
		query="UPDATE pcSales_Completed SET pcSC_BUComDate='" & dtTodaysDate & "',pcSC_BUTotal=" & tmpBUPrdCount & " WHERE pcSC_ID=" & pcSCID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		if tmpSCStatus=3 then
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate)) & " " & time()
			end if
			query="UPDATE pcSales_Completed SET pcSC_StoppedDate='" & dtTodaysDate & "',pcSC_Status=3 WHERE pcSC_ID=" & pcSCID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	
	%>
	<img src="images/pc_checkmark_sm.gif" width="16" height="16" alt="Task performed"> <b><%=tmpBUPrdCount%></b> product(s) have been backed up.<br><br>
	<%
	if Clng(tmpBUPrdCount)<>Clng(tmpPrdCount) then%>
	<div class="pcCPmessage">
	The Sale has been stopped because the number of backed-up product(s) is different than the number of affected product(s)
	</div>
	<br><br>
	<%end if
	'Change Prices
	IF tmpSCStatus=1 THEN
		Set cn = Server.CreateObject("ADODB.Connection")
		Set cmd = Server.CreateObject("ADODB.Command")
		cn.Open scDSN
		Set cmd.ActiveConnection = cn
		cmd.CommandText = "uspChangePrices"
		cmd.CommandType = adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@Param1") = tmpParam1
		cmd.Parameters("@Param2") = tmpParam2
		cmd.Parameters("@TPrice") = tmpTargetPrice
		cmd.Parameters("@CType") = tmpSaleType
		cmd.Parameters("@Relative") = tmpRelative
		if tmpSaleType="2" then
			cmd.Parameters("@Amount") = Round(tmpSaleAmount,2)*0.01
		else
			if tmpSaleType="1" then
				cmd.Parameters("@Amount") = Round(tmpSaleAmount,2)
			else
				cmd.Parameters("@Amount") = Round(100-cdbl(tmpSaleAmount),2)*0.01
			end if
		end if
		
		cmd.Parameters("@CRound") = tmpRound
		cmd.Parameters("@SCID") = pcSCID
		cmd.Execute
	
		tmpChangedPrdCount=cmd.Parameters("@SMCOUNT")

		Set cmd=nothing
		set cn=nothing
		
		%>
		<img src="images/pc_checkmark_sm.gif" width="16" height="16" alt="Task performed"> <b><%=tmpChangedPrdCount%></b> product(s) have been placed 'On Sale'.<br><br>
		<%
		
		if Clng(tmpChangedPrdCount)=Clng(tmpPrdCount) then
		
			query="UPDATE pcSales_Completed SET pcSC_STATUS=2 WHERE pcSC_ID=" & pcSCID & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		
			%>
			<div class="pcCPmessageSuccess">
		    The sale has been successfully started and is now active in your storefront. <br>
            <a href="../pc/showspecials.asp" target="_blank">View running sales</a> in the storefront. View &amp; edit <a href="sm_manage.asp">pending sales</a>. View <a href="sm_sales.asp">current &amp; completed sales</a>.</div>
			<br><br>
		<%else%>
			<div class="pcCPmessage">
				The Sale has been stopped because the number of backed-up product(s) is different than the number of affected product(s)
			</div>
			<br><br>
		<%end if
		
	END IF%>
	</td>
	</tr>
	<tr>
	<td class="pcCPspacer"></td>
	</tr>
	</table>
<%ELSE%>
<form name="Form1" action="sm_start.asp?a=start" method="post" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td>Please select the Sale that you would like to start:</td>
</tr>
<tr>
	<td class="pcCPspacer"></td>
</tr>
<tr>
	<td>
		<select name="id" id="id">
			<%For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i)%></option>
			<%Next%>
		</select>
	</td>
</tr>
<tr>
	<td class="pcCPspacer"><hr></td>
</tr>
<tr>
	<td>
		<input type="submit" name="Preview" value=" Continue " class="submit2">
	</td>
</tr>
</table>
</form>
<%END IF
END IF%>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->