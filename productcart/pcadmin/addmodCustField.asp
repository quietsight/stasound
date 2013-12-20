<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
if request("idcustfield")<>"" then
	pageTitle="View/Modify Special Customer Fields"
else
	pageTitle="Add New Special Customer Fields"
end if %>
<% Section="layout" %>
<%PmAdmin=7%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<% 
Dim connTemp,rs,query
call opendb()

pcv_ID=""
pcv_ID=request("idcustfield")

if request("action")="run" then
	pcv_FName=replace(request("pcv_FName"),"'","''")
	pcv_FName=replace(pcv_FName,"&quot;","""")
	pcv_FDesc=replace(request("pcv_FDesc"),"'","''")
	pcv_FDesc=replace(pcv_FDesc,"&quot;","""")
	pcv_FType=request("pcv_FType")
	if pcv_FType="" then
		pcv_FType="0"
	end if
	pcv_FLen=request("pcv_FLen")
	if (pcv_FLen="") or (pcv_FLen="0") then
		pcv_FLen="20"
	end if
	pcv_FMax=request("pcv_FMax")
	if pcv_FMax="" then
		pcv_FMax=0
	end if
	pcv_FReq=request("pcv_FReq")
	if pcv_FReq="" then
		pcv_FReq=0
	end if
	pcv_ShowReg=request("pcv_ShowReg")
	if pcv_ShowReg="" then
		pcv_ShowReg=0
	end if
	pcv_ShowCheckout=request("pcv_ShowCheckout")
	if pcv_ShowCheckout="" then
		pcv_ShowCheckout=0
	end if
	pcv_FValue=replace(request("pcv_FValue"),"'","''")
	pcv_Categories=request("pcv_Categories")
	if trim(pcv_Categories)="" then
		pcv_PricingCats=0
	else
		pcv_PricingCats=1
	end if
	
	if (pcv_ID<>"") and (pcv_ID<>"0") and IsNumeric(pcv_ID) then
		query="UPDATE pcCustomerFields SET pcCField_Name='" & pcv_FName & "',pcCField_Description='" & pcv_FDesc & "',pcCField_FieldType=" & pcv_FType & ",pcCField_Length=" & pcv_FLen & ",pcCField_Maximum=" & pcv_FMax & ",pcCField_Required=" & pcv_FReq & ",pcCField_ShowOnReg=" & pcv_ShowReg & ",pcCField_ShowOnCheckout=" & pcv_ShowCheckout & ",pcCField_Value='" & pcv_FValue & "',pcCField_PricingCategories=" & pcv_PricingCats & " WHERE pcCField_ID=" & pcv_ID & ";"
	else
		query="INSERT INTO pcCustomerFields (pcCField_Name,pcCField_Description,pcCField_FieldType,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_ShowOnReg,pcCField_ShowOnCheckout,pcCField_Value,pcCField_PricingCategories) VALUES ('" & pcv_FName & "','" & pcv_FDesc & "'," & pcv_FType & "," & pcv_FLen & "," & pcv_FMax & "," & pcv_FReq & "," & pcv_ShowReg & "," & pcv_ShowCheckout & ",'" & pcv_FValue & "'," & pcv_PricingCats & ");"
	end if
	
	set rs=connTemp.execute(query)
	set rs=nothing
	
	if (pcv_ID<>"") and (pcv_ID<>"0") and IsNumeric(pcv_ID) then
		pcMessage=4
	else
		query="SELECT pcCField_ID FROM pcCustomerFields Order by pcCField_ID desc;"
		set rs=connTemp.execute(query)
		pcv_ID=rs("pcCField_ID")
		set rs=nothing
		pcMessage=3
	end if
	
	query="DELETE FROM pcCustFieldsPricingCats WHERE pcCField_ID=" & pcv_ID & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	
	if trim(pcv_Categories)<>"" then
		pcA=split(pcv_Categories,",")
		For i=lbound(pcA) to ubound(pcA)
			if trim(pcA(i))<>"" then
				query="INSERT INTO pcCustFieldsPricingCats (pcCField_ID,idCustomerCategory) VALUES (" & pcv_ID & "," & pcA(i) & ");"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		Next
	end if
	
	call closeDb()
	response.redirect "manageCustFields.asp?message=" & pcMessage

end if

pcv_FName=""
pcv_FDesc=""
pcv_FType=0
pcv_FLen=20
pcv_FMax=0
pcv_FReq=""
pcv_ShowReg=""
pcv_ShowCheckout=""
pcv_FValue=""
pcv_PricingCats=""

'Get data from database if view/modify existing field
if (pcv_ID<>"") and (pcv_ID<>"0") and IsNumeric(pcv_ID) then

	query="SELECT pcCField_ID,pcCField_Name,pcCField_Description,pcCField_FieldType,pcCField_Length,pcCField_Maximum,pcCField_Required,pcCField_ShowOnReg,pcCField_ShowOnCheckout,pcCField_Value,pcCField_PricingCategories FROM pcCustomerFields WHERE pcCField_ID=" & pcv_ID & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		pcv_ID=rs("pcCField_ID")
		pcv_FName=rs("pcCField_Name")
		pcv_FDesc=rs("pcCField_Description")
		pcv_FType=rs("pcCField_FieldType")
		pcv_FLen=rs("pcCField_Length")
		pcv_FMax=rs("pcCField_Maximum")
		pcv_FReq=rs("pcCField_Required")
		pcv_ShowReg=rs("pcCField_ShowOnReg")
		pcv_ShowCheckout=rs("pcCField_ShowOnCheckout")
		pcv_FValue=rs("pcCField_Value")
		pcv_PricingCats=rs("pcCField_PricingCategories")
		if pcv_FName<>"" AND isNULL(pcv_FName)=False then
			pcv_FName=replace(pcv_FName,"""","&quot;")
		end if
		if pcv_FDesc<>"" AND isNULL(pcv_FDesc)=False then
			pcv_FDesc=replace(pcv_FDesc,"""","&quot;")
		end if
	end if

	set rs=nothing

end if

%>
<!--#include file="AdminHeader.asp"-->

<script language="JavaScript">
<!--

function isDigit(s)
{
	var test=""+s;
	if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
	{
		return(true) ;
	}
	return(false);
}
	
function allDigit(s)
{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (isDigit(c)==false)
		{
			return (false);
		}
	}
	return (true);
}

function Form1_Validator(theForm)
{
	if (theForm.pcv_FName.value=="")
	{
		alert("Please enter value for this field.")
		theForm.pcv_FName.focus();
		return (false);
	}
	
	if (theForm.pcv_FDesc.value=="")
	{
		alert("Please enter value for this field.")
		theForm.pcv_FDesc.focus();
		return (false);
	}
	
	if ((theForm.pcv_ShowReg.checked==false) && (theForm.pcv_ShowCheckout.checked==false))
	{
		if (confirm('You selected not to show this field on both the checkout and registration page. This means that the field will not be shown anywhere in the storefront. Are you sure?')==false)
		{
			return (false);
		}
	}
	
	if (theForm.pcv_FLen.value!="")
	{
		if (allDigit(theForm.pcv_FLen.value) == false)
		{
			alert("Please enter a integer value for this field.");
			theForm.pcv_FLen.focus();
			return (false);
		}
	}
	
	if (theForm.pcv_FMax.value!="")
	{
		if (allDigit(theForm.pcv_FMax.value) == false)
		{
			alert("Please enter a integer value for this field.");
			theForm.pcv_FMax.focus();
			return (false);
		}
	}

	return (true);
}

function SetCheckBox(tvalue)
{
	if (tvalue!="")
	{
		document.getElementById("pcv_ShowReg1").disabled=true;
		document.getElementById("pcv_ShowReg1").checked=false;
	}
	else
	{
		document.getElementById("pcv_ShowReg1").disabled=false;
	}
}

// -->	
</script>
	 
<form name=form1 action="addmodCustField.asp?action=run" method="post" class="pcForms" onsubmit="return Form1_Validator(this)">
	<table class="pcCPcontent">
		<tr>
			<td width="20%">Field Label:</td>
			<td width="80%"><input type="text" size="40" name="pcv_FName" value="<%=pcv_FName%>"></td>
		</tr>
		<tr>
			<td valign="top">Field Description:</td>
			<td><textarea name="pcv_FDesc" rows="10" cols="60"><%=pcv_FDesc%></textarea></td>
		</tr>
		<tr>
			<td valign="top">Field Type:</td>
			<td>
				<input type="radio" name="pcv_FType" value="0" <%if pcv_FType="0" then%>checked<%end if%> class="clearBorder"> Input Field<br>
				<input type="radio" name="pcv_FType" value="1" <%if pcv_FType="1" then%>checked<%end if%> class="clearBorder"> Checkbox
			</td>
		</tr>
		<tr>
			<td>Field Value:</td>
			<td><input type="text" size="40" name="pcv_FValue" value="<%=pcv_FValue%>">&nbsp;<em>E.g. &quot;<strong>Yes</strong>&quot;</em>&nbsp;&nbsp;<span class="pcCPnotes">NOTE: This is needed if the Field Type is &quot;Checkbox&quot;</span>
			</td>
		</tr>
		<tr>
			<td>Field Length:</td>
			<td><input type="text" size="40" name="pcv_FLen" value="<%=pcv_FLen%>"></td>
		</tr>
		<tr>
			<td>Field Maximum:</td>
			<td><input type="text" size="40" name="pcv_FMax" value="<%=pcv_FMax%>"> <i>(0 = unlimited)</i></td>
		</tr>
		<tr>
			<td>Field Required:</td>
			<td><input type="checkbox" value="1" name="pcv_FReq" <%if pcv_FReq="1" then%>checked<%end if%> class="clearBorder"></td>
		</tr>
		<tr>
			<td align="right"><input id="pcv_ShowReg1" type="checkbox" name="pcv_ShowReg" value="1" <%if pcv_ShowReg="1" then%>checked<%end if%> class="clearBorder"></td>
			<td>Show on registration page</td>
		</tr>
		<tr>
			<td align="right"><input type="checkbox" name="pcv_ShowCheckout" value="1" <%if pcv_ShowCheckout="1" then%>checked<%end if%> class="clearBorder"></td>
			<td>Show on checkout page</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<%query="SELECT idCustomerCategory,pcCC_Name FROM pcCustomerCategories;"
		set rs=connTemp.execute(query)
		
		if not rs.eof then
		pcArr=rs.GetRows()
		intCount=ubound(pcArr,2)
		set rs=nothing%>
		<tr>
			<td colspan="2">
				Show only to customers in the following Pricing Categories:
            </td>
        </tr>
        <tr>
        	<td valign="top">
				<div class="pcSmallText">To select more than one category, keep the CTRL button down)</div>
            </td>
			<td>
				<select id="pcv_Categories" name="pcv_Categories" multiple size="10" onchange="javascript:SetCheckBox(this.value);">
				<%
				pcv_HaveSelectCat=0
				For i=0 to intCount
					pcv_SelectCat=0
					if (pcv_ID<>"") and (pcv_ID<>"0") and IsNumeric(pcv_ID) then
						query="SELECT idCustomerCategory FROM pcCustFieldsPricingCats WHERE pcCField_ID=" & pcv_ID & " and  idCustomerCategory=" & pcArr(0,i) & ";"
						set rs=connTemp.execute(query)
						if not rs.eof then
							pcv_SelectCat=1
							pcv_HaveSelectCat=1
						end if
						set rs=nothing
					end if%>
					<option value="<%=pcArr(0,i)%>" <%if pcv_SelectCat=1 then%>selected<%end if%>><%=pcArr(1,i)%></option>
				<%Next%>
				</select>
				<%if pcv_HaveSelectCat=1 then%>
				<script>
					SetCheckBox("ei");
				</script>
				<%end if%>
			</td>
		</tr>
		<%
		end if
		set rs=nothing
		call closeDb()
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center">
			<input type="submit" name="submit" value="<%if (pcv_ID<>"") and (pcv_ID<>"0") and IsNumeric(pcv_ID) then%>Update Special Customer Field<%else%>Add New Special Customer Field<%end if%>" class="submit2">
			&nbsp;
			<input type="button" name="back" value="Back" onClick="javascript:location='manageCustFields.asp';">
			<input type="hidden" name="idcustfield" value="<%=pcv_ID%>">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->