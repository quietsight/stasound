<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.Buffer = true %>
<% pageTitle="Edit State-specific Tax Exemption" %>
<%PmAdmin=6%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/pcCategoriesList.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim rstemp, connTemp, strSQL, pid
Dim LogicIssues,LogicErrStr

pcv_TaxRateID=request("rid")
if trim(pcv_TaxRateID)="" then
	response.redirect "manageTaxEptZone.asp"
end if

Dim Arr1 (100,4)

Dim pcArr1
TotalRules=0

RulesCount=0

LogicIssues=0
HaveIssues=0
LogicErrStr=""

ErrorStr=""
ItemName=""
MsgBack=""
ErrorArr=""

call opendb()

if (request("action")="upd") and (request("submit2")<>"") then
	UpdSucess=0
	Count=request.form("Count")
	if Count>"0" then
	For i=1 to Count
	IF request("C" & i)="1" THEN
		tmp1=request.form("pcv_ID1_" & i)
		tmp2=request.form("pcv_ID2_" & i)
		tmp3=request.form("pcv_ID3_" & i)
		query="DELETE FROM pcTaxEpt WHERE (pcTEpt_StateCode LIKE 'NONE') AND (pcTEpt_ProductList LIKE '" & tmp2 & "') AND (pcTEpt_CategoryList LIKE '" & tmp3 & "') AND (pcTaxZoneRate_ID LIKE '" & tmp1 & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
	END IF
	Next
	msg="Tax Rules were deleted successfully!"
	response.redirect "manageTaxEptZone.asp?s=1&msg=" & Server.UrlEncode(msg)
	end if
end if


if (request("action")="upd") and (request("submit1")<>"") then
	UpdSucess=0
	Count=request.form("Count")
	if Count>"0" then

	For i=1 to Count
	IF request("C" & i)="1" THEN
		tmp1=request.form("pcv_ID1_" & i)
		tmp4=request.form("pcv_ID2_" & i)
		tmp5=request.form("pcv_ID3_" & i)
		tmp2=request.form("pcv_Rule2_" & i)
		if tmp2<>"" then
			tmp2=replace(tmp2," ","")
			if right(tmp2,1)<>"," then
				tmp2=tmp2 & ","
			end if
		end if
		tmp3=request.form("pcv_Rule3_" & i)
		if tmp3<>"" then
			tmp3=replace(tmp3," ","")
			if right(tmp3,1)<>"," then
				tmp3=tmp3 & ","
			end if
		end if
		tmp6=0
		if tmp2 & tmp3 <>"" then
			query="UPDATE pcTaxEpt Set pcTEpt_ProductList='" & tmp2 & "',pcTEpt_CategoryList='" & tmp3 & "', pcTEpt_EptAll=" & tmp6 & " WHERE (pcTEpt_StateCode LIKE 'NONE') AND (pcTEpt_ProductList LIKE '" & tmp4 & "') AND (pcTEpt_CategoryList LIKE '" & tmp5 & "') AND (pcTaxZoneRate_ID LIKE '" & tmp1 & "');"
			set rs=connTemp.execute(query)
			set rs=nothing
			UpdSucess=UpdSucess+1
		else
			ErrorStr="Cannot update some tax rules because they are not logical."
		end if
	END IF
	Next
	if UpdSucess>0 then
	if ErrorStr="" then
		msg="Tax Rules were updated successfully!"
	else
		msg="All others rules were updated successfully!"
	end if
	end if
	end if
	call closedb()
	response.redirect "editTaxEptZone.asp?rid=" & pcv_TaxRateID & "&s=1&msg=" & Server.UrlEncode(msg) & "&ErrorStr=" & Server.UrlEncode(ErrorStr)
end if

Function getStateList()
	Dim tmp1,rs1
	tmp1=""
	query="SELECT stateCode,stateName FROM states WHERE NOT (stateName like '%Canada%') order by stateCode;"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArray1=rs1.getRows()
		intCount1=ubound(pcArray1,2)
		set rs1=nothing
		For i=0 to intCount1
			tmp1=tmp1 & "<option value=""" & pcArray1(0,i) & """>" & pcArray1(1,i) & "</option>" & vbcrlf
		Next
	end if
	set rs1=nothing
	getStateList=tmp1
End Function

Function getItemList()
	Dim tmp1,rs1
	tmp1=""
	query="SELECT idproduct,description FROM products WHERE active=-1 AND removed=0 AND configOnly=0 order by description;"
	set rs1=connTemp.execute(query)
	
	if not rs1.eof then
		pcArray1=rs1.getRows()
		intCount1=ubound(pcArray1,2)
		set rs1=nothing
		For i=0 to intCount1
			tmp1=tmp1 & "<option value=""" & pcArray1(0,i) & """>" & pcArray1(1,i) & "</option>" & vbcrlf
		Next
	end if
	set rs1=nothing
	getItemList=tmp1
End Function
	
%>
<% 
if request.querystring("ErrorStr")<>"" then %>
	<div class="pcCPmessage"><%=request.querystring("ErrorStr")%></div>
<% end if %>
<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<script>
	function UpDown(tabid)
	{
	var etab=eval(tabid);
		if (etab.style.display=='')
		{ etab.style.display='none';
		}
		else
		{ etab.style.display='';
		}
	}
</script>
<form method="POST" name=Form1 action="editTaxEptZone.asp?action=upd" class="pcForms">
<input type=hidden name="rid" value="<%=pcv_TaxRateID%>">
	<table class="pcCPcontent">
	<tr>
	<td colspan="2">To select more than one item, keep the CTRL key pressed down.</td>
	</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
<% query="SELECT pcTaxEpt.pcTaxZoneRate_ID, pcTaxEpt.pcTEpt_ProductList, pcTaxEpt.pcTEpt_CategoryList, pcTaxEpt.pcTEpt_EptAll, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneDescriptions.pcTaxZoneDesc FROM pcTaxZoneDescriptions INNER JOIN ((pcTaxEpt INNER JOIN pcTaxZoneRates ON pcTaxEpt.pcTaxZoneRate_ID = pcTaxZoneRates.pcTaxZoneRate_ID) INNER JOIN pcTaxZonesGroups ON pcTaxZoneRates.pcTaxZoneRate_ID = pcTaxZonesGroups.pcTaxZoneRate_ID) ON pcTaxZoneDescriptions.pcTaxZoneDesc_ID = pcTaxZonesGroups.pcTaxZoneDesc_ID WHERE (((pcTaxEpt.pcTaxZoneRate_ID)="&pcv_TaxRateID&"));"
set rs=connTemp.execute(query)
Count=0
if rs.eof then
set rs=nothing%>
<tr>
<td colspan="3">
<p><strong>This Store does not have Zone Tax Exemptions</strong></p></td></tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
<%else
	pcArray=rs.GetRows()
	intCount=ubound(pcArray,2)
	pcArr1=pcArray
	TotalRules=intCount
	set rs=nothing

tmpStr2=""
tmpStr2=getItemList()
%>
<tr> 
<th align="left" nowrap> 
<p>Rule : Zone</p></th>
<th align="left" nowrap> 
<p>Exemptions by Products &amp; Categories</p></th>
</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
<%For i=0 to intCount
	Count=Count+1
	tmpStr1=""
	tmpStr4=""
	tmpStr5=""
	
	tmpStr4=tmpStr2
	tmpStr5=tmpStr3
	
	tmpA=split(pcArray(1,i),",")
	For j=lbound(tmpA) to ubound(tmpA)
		if trim(tmpA(j))<>"" then
			tmpStr4=replace(tmpStr4,"value=""" & tmpA(j) & """>","value=""" & tmpA(j) & """ selected>")
		end if
	Next
	
	tmpB=pcArray(2,i)

%>
<tr>
<td valign="top" nowrap><p>
<input type=hidden name="C<%=Count%>" value="1">
<%=pcArray(5,i)&" : "&pcArray(4,i)%>
<input type=hidden name="pcv_ID1_<%=Count%>" value="<%=pcArray(0,i)%>">
<input type=hidden name="pcv_ID2_<%=Count%>" value="<%=pcArray(1,i)%>">
<input type=hidden name="pcv_ID3_<%=Count%>" value="<%=pcArray(2,i)%>"></p></td>
<td valign="top"><p>
Select tax-exempt products:<br>
<select name="pcv_Rule2_<%=Count%>" multiple size="5">
<%=tmpStr4%>
</select>
<br><br>
Select tax-exempt categories:<br>
	<%pcv_tmp1=tmpB
	if pcv_tmp1<>"" then
		pcv_tmp1=replace(pcv_tmp1," ","")
	end if
	cat_DropDownName="pcv_Rule3_" & Count
	cat_Type="1"
	cat_DropDownSize="5"
	cat_MultiSelect="1"
	cat_ExcBTOHide="1"
	cat_StoreFront="0"
	cat_ShowParent="1"
	cat_DefaultItem=""
	cat_SelectedItems=pcv_tmp1
	cat_ExcItems=""
	cat_ExcSubs="0"
	cat_EventAction=""
	%>
	<%call pcs_CatList()%></p></td>
</tr>
	<tr>
		<td colspan="8" class="pcCPspacer"></td>
	</tr>
<%Next%>
<%end if%>
<tr>
	<td colspan="8"><hr></td>
</tr>
<tr> 
<td></td>
<td><%if Count<>0 then%>
  <input type=submit name=submit1 value=" Update " class="submit2">
  &nbsp;
  <input type=submit name=submit2 value=" Delete " onclick="javascript:if (confirm('You are about to remove selected tax rules from your database. Are you sure you want to complete this action?')) {return(true)} else {return(false)}">&nbsp;
  <%end if%>
  <input type=button name=addnew value=" Add New Exemption Rules " onclick="location='AddTaxEptZone.asp';">
  &nbsp;
  <input type=button name=goback value=" Back " onclick="location='manageTaxEptZone.asp';">
  <input type=hidden name=Count value="<%=Count%>">
  </td>
</tr>                   
</table>
</form>
<%call closedb()%><!--#include file="AdminFooter.asp"-->