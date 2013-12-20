<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.Buffer = true %>
<% pageTitle="Manage State-specific Tax Exemption" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% 
Dim rstemp, connTemp, strSQL, pid
Dim LogicIssues,LogicErrStr

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
	set rs = nothing
	call closedb()
	msg="Tax Rules were deleted successfully!"
	response.redirect "manageTaxEptZone.asp?s=1&msg=" & Server.UrlEncode(msg)
	end if
end if

Dim pcArray2,intCount2,tmpParent
tmpParent=""

Function getParentList()
Dim rs1
query="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory FROM categories WHERE ((categories.idCategory) Not in (SELECT categories_products.idCategory FROM categories_products)) AND idCategory<>1 ORDER BY categories.idcategory asc;"
set rs1=conntemp.execute(query)

if not rs1.eof then
	pcArray2=rs1.getRows()
	intCount2=ubound(pcArray2,2)
	set rs1=nothing
end if
set rs1=nothing
End Function

Function FindParent(idCat)
	Dim k
	if clng(idCat)<>1 then
	For k=0 to intCount2
		if (clng(pcArray2(0,k))=clng(idCat)) and (clng(pcArray2(0,k))<>1)	then
			if tmpParent<>"" then
			tmpParent="/" & tmpParent
			end if
			tmpParent=pcArray2(1,k) & tmpParent
			FindParent(pcArray2(2,k))
			exit for
		end if
	Next
	end if
End function

if request.querystring("ErrorStr")<>"" then %>
	<div class="pcCPmessage">
		<b>Error(s):</b>
		<%=request.querystring("ErrorStr")%>
	</div>
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
<form method="POST" name=Form1 action="manageTaxEptZone.asp?action=upd" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="3"><p>To select more than one item, keep the CTRL key pressed down.</p></td>
    </tr>
    <tr>
        <td colspan="8" class="pcCPspacer"></td>
    </tr>
	<%
    query="SELECT pcTaxEpt.pcTaxZoneRate_ID, pcTaxEpt.pcTEpt_ProductList, pcTaxEpt.pcTEpt_CategoryList, pcTaxZoneRates.pcTaxZoneRate_Name, pcTaxZoneDescriptions.pcTaxZoneDesc FROM pcTaxZoneDescriptions INNER JOIN ((pcTaxEpt INNER JOIN pcTaxZoneRates ON pcTaxEpt.pcTaxZoneRate_ID = pcTaxZoneRates.pcTaxZoneRate_ID) INNER JOIN pcTaxZonesGroups ON pcTaxZoneRates.pcTaxZoneRate_ID = pcTaxZonesGroups.pcTaxZoneRate_ID) ON pcTaxZoneDescriptions.pcTaxZoneDesc_ID = pcTaxZonesGroups.pcTaxZoneDesc_ID;"' WHERE (((pcTaxEpt.pcTaxZoneRate_ID)=38));"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=connTemp.execute(query)
    Count=0
    if rs.eof then
    set rs=nothing%>
		<tr><td colspan="3"><div class="pcCPmessage">No product/category tax exemption has been setup for this tax rule.</div></td></tr>
<%else
	pcArray=rs.GetRows()
	intCount=ubound(pcArray,2)
	pcArr1=pcArray
	TotalRules=intCount
	set rs=nothing

	getParentList()
	%>
    <tr> 
    <th nowrap>Rule : Zone</th>
    <th nowrap>Exemptions by Products &amp; Categories</th>
    <th></th>
    </tr>
    <tr>
        <td colspan="8" class="pcCPspacer"></td>
    </tr>
<%For i=0 to intCount
	Count=Count+1
	tmpStr4=""
	tmpStr5=""
	
	tmpA=split(pcArray(1,i),",")
	query=""
	For j=lbound(tmpA) to ubound(tmpA)
		if trim(tmpA(j))<>"" then
			if query="" then
				query="SELECT description FROM Products WHERE idproduct=" & tmpA(j)
			else
				query=query & " OR idproduct=" & tmpA(j)
			end if
		end if
	Next
	if query<>"" then
		query=query & " ORDER by description ASC;"
		set rsP=connTemp.execute(query)
		if not rsP.eof then
			pcA=rsP.getRows()
			intA=ubound(pcA,2)
			set rsP=nothing
			For j=0 to intA
				tmpStr4=tmpStr4 & "<li>" & pcA(0,j) & "</li>" & vbcrlf
			Next
		end if
		set rsP=nothing
	end if
	
	if tmpStr4<>"" then
	tmpStr4="<ul>" & tmpStr4 & "</ul>"
	end if
	
	tmpB=split(pcArray(2,i),",")
	query=""
	For j=lbound(tmpB) to ubound(tmpB)
	if trim(tmpB(j))<>"" then
		if query="" then
			query="SELECT idCategory,idParentCategory,categoryDesc FROM Categories WHERE idCategory=" & tmpB(j)
		else
			query=query & " OR idCategory=" & tmpB(j)
		end if
	end if
	Next
	
	if query<>"" then
		query=query & " ORDER by categoryDesc ASC;"
		set rsP=connTemp.execute(query)
		if not rsP.eof then
			pcA=rsP.getRows()
			intA=ubound(pcA,2)
			set rsP=nothing
			For j=0 to intA
				tmpParent=""
				FindParent(pcA(1,j))
				if tmpParent<>"" then
				tmpParent=" [" & tmpParent & "]"
				end if
				tmpStr5=tmpStr5 & "<li>" & pcA(2,j) & tmpParent & "</li>" & vbcrlf
			Next
		end if
		set rsP=nothing
	end if
	if tmpStr5<>"" then
	tmpStr5="<ul>" & tmpStr5 & "</ul>"
	end if

%>
        <tr>
        <td valign="top" nowrap>
        <input type=checkbox name="C<%=Count%>" value="1" class="clearBorder">
        <%=pcArray(4,i)&" : "&pcArray(3,i)%>
        <input type=hidden name="pcv_ID1_<%=Count%>" value="<%=pcArray(0,i)%>">
        <input type=hidden name="pcv_ID2_<%=Count%>" value="<%=pcArray(1,i)%>">
        <input type=hidden name="pcv_ID3_<%=Count%>" value="<%=pcArray(2,i)%>">
        </td>
        <td valign="top">
        <%if tmpStr4<>"" then%>
        Tax-exempt products:<br>
        <%=tmpStr4%>
        <%end if%>
        <%if tmpStr5<>"" then%>
        Tax-exempt categories:<br>
        <%=tmpStr5%>
        <%end if%>
        </td>
        <td valign="top"><a href="editTaxEptZone.asp?rid=<%=pcArray(0,i)%>">Edit</a></td>
        </tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
<%Next%>
<%end if%>
<%if Count<>0 then%>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
<tr>
<td colspan="3" class="cpLinksList">
<a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a>
</td>
</tr>
<%end if%>
<tr>
    <td colspan="8" class="pcCPspacer"><hr></td>
</tr>
<tr> 
<td colspan="3">
<input type="button" name="addnew" value="Add New Zone-based Tax Exemption" onclick="location='AddTaxEptZone.asp';" class="submit2">&nbsp;
<%if Count<>0 then%>
<input type=submit name=submit2 value=" Delete Selected " onclick="javascript:if (confirm('You are about to remove selected tax rules from your database. Are you sure you want to complete this action?')) {return(true)} else {return(false)}">&nbsp;
<%end if%>
<input type="button" name="goback" value="Back" onclick="location='viewTax.asp'">
<input type="hidden" name="Count" value="<%=Count%>">
</td>
</tr>                   
</table>
<%if Count>0 then%>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.Form1.C" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.Form1.C" + j); 
if (box.checked == true) box.checked = false;
   }
}

//-->
</script>
<%end if%>
</form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->