<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews: View/Edit Values for Drop-down List"
pageIcon="pcv4_icon_reviews.png"
section="reviews"
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% 

Dim rs, connTemp, query
Dim pcv_ID
pcv_ID=request("IDField")

if not validNum(pcv_ID) then
	response.redirect "prv_FieldManager.asp"
end if
	
IF request("action")="add" then
	
	call openDb()

	query="DELETE FROM pcRevLists WHERE pcRL_IDField=" & pcv_ID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	Dim pcv_Values
	pcv_Values=request("pcv_values")
	set rs=nothing
	
	pcArray=split(pcv_Values,vbcrlf)
	For k=lbound(pcArray) to ubound(pcArray)
		if trim(pcArray(k))<>"" then
				pcv_OName=trim(pcArray(k))
				if (pcv_OName<>"") then
				pcv_OName=replace(pcv_OName,"'","''")
				pcv_OValue=pcv_OName
				query="INSERT INTO pcRevLists (pcRL_IDField,pcRL_Name,pcRL_Value) VALUES (" & pcv_ID & ",'" & pcv_OName & "','" & pcv_OValue & "')"  
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				set rs=nothing
				end if
		end if
	Next
	
	call closedb()
	response.redirect "prv_FieldManager.asp"

END IF

call opendb()
query="SELECT pcRL_Name,pcRL_Value FROM pcRevLists WHERE pcRL_IDField=" & pcv_ID
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_strValues=""
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
	for k=0 to intCount
		pcv_strValues=pcv_strValues & pcArray(1,k) & vbcrlf
	next
end if
set rs=nothing

query="SELECT pcRF_Name FROM pcRevFields WHERE pcRF_IDField=" & pcv_ID
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_FieldName=rs("pcRF_Name")
set rs=nothing
call closedb()

%>
<script>
function Form1_Validator(theForm)
{
  if (theForm.pcv_values.value == "")
  {
  alert("Enter a value for this field.");
  theForm.pcv_values.focus();
  return(false);
  }
  return(true);
}
</script>

<form method="POST" action="prv_EditList.asp?action=add" name="checkboxform" onSubmit="return Form1_Validator(this)" class="pcForms">
	<input type="hidden" name="IDField" value="<%=request("IDField")%>">
	<table class="pcCPcontent">
		<tr> 
			<td nowrap width="20%" align="right">Drop-down List:</td>
			<td align="left" width="80%"><b><%=pcv_FieldName%></b></td>
		</tr>
		<tr> 
			<td nowrap valign="top" align="right">Values:
			<div class="pcCPnotes" style="margin-top: 10px;">One value per line. E.g.:<br>
			<em>Yes<br>
			No<br>
			Maybe</em>
            </div>
			</td>
			<td>
				<textarea rows="10" cols="55" name="pcv_values"><%=pcv_strValues%></textarea>
			</td>
		</tr>
		<tr>
			<td align="center" colspan="2">
				<input type="submit" value=" Update Values " class="submit2">&nbsp;
				<input type="button" value="Back" onClick="javascript:history.back()">
			 </td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->