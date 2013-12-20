<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
validfields=request.form("validfields")

R1=0
R2=0
R3=0
R4=0
mfilter=""
herror=false
mtemp=""
D1=0
D2=0
S1=0
S2=0

For i=1 to validfields

	Select Case ucase(request("T" & i))
		Case "SKU": R1=1
		Case "NAME": R2=1
		Case "DESCRIPTION": R3=1
		Case "ONLINE PRICE": R4=1
		Case "PRODUCT TYPE": D1=1
		Case "DOWNLOADABLE FILE LOCATION","MAKE DOWNLOAD URL EXPIRE","URL EXPIRATION IN DAYS","USE LICENSE GENERATOR","LOCAL GENERATOR","REMOTE GENERATOR": D2=1
        Case "LICENSE FIELD LABEL (1)","LICENSE FIELD LABEL (2)","LICENSE FIELD LABEL (3)","LICENSE FIELD LABEL (4)","LICENSE FIELD LABEL (5)","ADDITIONAL COPY": D2=1
        Case "DROP-SHIPPER ID": S1=1
        Case "DROP-SHIPPER IS ALSO A SUPPLIER": S2=1
	End Select
	if trim(ucase(request("T" & i)))<>"" then
	if instr(mfilter,"*" & ucase(request("T" & i)) & "*")>0 then
		herror=true
	else
		mfilter=mfilter & "*" & ucase(request("T" & i)) & "*"
	end if
	end if
	mtemp= mtemp & "&" & "T" & i & "=" & request("T" & i)
Next

if (session("append")="1") then
	if (R1=0) then
		msg="Please make sure that the SKU field is mapped."
		response.redirect "step2.asp?msg=" & msg & mtemp
	end if
else
	if r1+r2+r3+r4<>4 then
		msg="Please make sure that the following fields are mapped:<br>1. SKU<br>2. Name<br>3. Description<br>4. Online Price"
		response.redirect "step2.asp?msg=" & msg & mtemp
	end if
end if

if herror=true then
	msg="Some of the mapping instructions are overlapping. Please make sure that the database fields are mapped uniquely."
	response.redirect "step2.asp?msg=" & msg & mtemp
end if
if (D2=1) and (D1<>1) then
	msg="Please map the 'Product Type' field before import/append any Downloadable Product Fields."
	response.redirect "step2.asp?msg=" & msg & mtemp
end if
if (S2=1) and (S1<>1) then
	msg="Please map the 'Drop-Shipper ID' field before map the 'Drop-Shipper is also a Supplier' field."
	response.redirect "step2.asp?msg=" & msg & mtemp
end if

if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " PRODUCT DATA WIZARD - Confirm Mappings"
%>
<% section = "products" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form method="post" action="step4.asp" class="pcForms">
<table class="pcCPcontent">
    <tr>
	<td valign="top" colspan="2">
        <table class="pcCPcontent">
        <tr>
            <td colspan="2"><h2>Steps:</h2></td>
        </tr>
        <tr>
            <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
            <td width="95%"><font color="#A8A8A8">Select product data file</font></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step2.gif"></td>
            <td><font color="#A8A8A8">Map fields</font></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step3a.gif"></td>
            <td><strong>Confirm mapping</strong></td>
        </tr>
        <tr>
            <td align="right"><img border="0" src="images/step4.gif"></td>
            <td><font color="#A8A8A8"><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</font></td>
        </tr>
        </table>
	<br>
	Please make sure that the database fields are mapped correctly. If not, click &quot;Back to Step 2&quot; button to try again.</font><br>
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th width="40%">From:</th>
	<th width="60%">To:</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<% validfields=0
For i=1 to request("validfields")
	if trim(request("T" & i))<>"" then
	validfields=validfields+1%>
<tr>
	<td width="40%">
		<%=request("F" & i)%>
		<input type=hidden name="P<%=validfields%>" value="<%=request("P" & i)%>">
		<input type=hidden name="F<%=validfields%>" value="<%=request("F" & i)%>">
	</td>
	<td width="60%">
		<%=request("T" & i)%>
		<input type=hidden name="T<%=validfields%>" value="<%=request("T" & i)%>">
	</td>
</tr>
	<%end if
Next%>  
<tr>
	<td colspan="2" class="pcCPspacer"><hr></td>
</tr>                 
<tr>
	<td colspan="2">
		<input type="hidden" name="validfields" value="<%=validfields%>">         
		<input type="button" name="backstep" value="<< Back to Step 2" onClick="location='step2.asp?a=1<%=mtemp%>';">&nbsp; 
        <input type="submit" name="submit" value="Go to Step 4 >>" onClick="pcf_Open_Import();" class="submit2">
 		<%
        '// Loading Window
        '	>> Call Method with OpenHS();
        response.Write(pcf_ModalWindow("Updating store database... This could take several minutes. Do not close this page.", "Import", 300))
        %>
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->