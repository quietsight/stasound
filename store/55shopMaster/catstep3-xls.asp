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

For i=1 to validfields
	Select Case request("T" & i)
		Case "Category Name": R1=1
		Case "Category ID": R2=1
		Case "Featured Sub-Category Name":
			if (session("append")<>"1") then
				R3=1
			end if
		Case"Featured Sub-Category ID":
			if (session("append")<>"1") then
				R4=1
			end if
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
	if (R1=0) and (R2=0) then
		msg="Please make sure that the Category Name field or the Category ID field is mapped."
		response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
	end if
else
	if (R1=0) then
		msg="Please make sure that the Category Name field is mapped."
		response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
	end if
	if (R2=1) then
		msg="Can not map the Category ID field when import new category."
		response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
	end if
	if (R3=1) then
		msg="Can not map the Featured Sub-Category Name field when import new category."
		response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
	end if
	if (R4=1) then
		msg="Can not map the Featured Sub-Category ID field when import new category."
		response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
	end if
end if

if herror=true then
	msg="Some of the mapping instructions are overlapping. Please make sure that the database fields are mapped uniquely."
	response.redirect "catstep2-xls.asp?msg=" & msg & mtemp
end if

if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " CATEGORY DATA WIZARD - Confirm Mappings"
%>
<% section = "products" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<form method="post" action="catstep4-xls.asp" class="pcForms">
<table class="pcCPcontent">
    <tr>
	<td valign="top" colspan="2">
    <table class="pcCPcontent">
    <tr>
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Select category data file</font></td>
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
	<td>
		<%=request("F" & i)%>
		<input type=hidden name="P<%=validfields%>" value="<%=request("P" & i)%>">
		<input type=hidden name="F<%=validfields%>" value="<%=request("F" & i)%>">
	</td>
	<td>
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
		<input type="button" name="backstep" value="<< Back to Step 2" onClick="location='catstep2-xls.asp?a=1<%=mtemp%>';">&nbsp; 
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