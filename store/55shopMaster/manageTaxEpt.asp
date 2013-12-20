<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.Buffer = true %>
<% pageTitle="Tax by Location - Manage Product and Category Tax Exemptions" %>
<%PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/taxsettings.asp" -->
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
		query="DELETE FROM pcTaxEpt WHERE (pcTEpt_StateCode LIKE '" & tmp1 & "') AND (pcTEpt_ProductList LIKE '" & tmp2 & "') AND (pcTEpt_CategoryList LIKE '" & tmp3 & "');"
		set rs=connTemp.execute(query)
		set rs=nothing
	END IF
	Next
	msg="Tax Rules were deleted successfully!"
	call closeDb()
	response.redirect "manageTaxEpt.asp?s=1&msg=" & Server.UrlEncode(msg)
	end if
end if

Dim pcArray2,intCount2,tmpParent
tmpParent=""

Function getParentList()
Dim rs1
query="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory FROM categories WHERE ((categories.idCategory) Not in (SELECT categories_products.idCategory FROM categories_products)) AND idCategory<>1 ORDER BY categories.idcategory asc;"
set rs1=Server.CreateObject("ADODB.Recordset")
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

%>

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
<form method="POST" name="Form1" action="manageTaxEpt.asp?action=upd" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="3">To select more than one item, keep the CTRL key pressed down.</td>
    </tr>
	<%
    query="SELECT pcTaxEpt.pcTEpt_StateCode,states.StateName,pcTaxEpt.pcTEpt_ProductList,pcTaxEpt.pcTEpt_CategoryList FROM pcTaxEpt,States WHERE States.StateCode=pcTaxEpt.pcTEpt_StateCode;"
    set rs=connTemp.execute(query)
    Count=0
    if rs.eof then
    set rs=nothing%>
    <tr>
        <td colspan="3">
        <div class="pcCPmessage">This Store does not have location-specific Tax Exemptions</div>
        </td>
    </tr>
	<%
	else
		pcArray=rs.GetRows()
		intCount=ubound(pcArray,2)
		pcArr1=pcArray
		TotalRules=intCount
		set rs=nothing
		getParentList()
	%>
    <tr> 
	    <td colspan="3" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th align="left" nowrap>State/Province</th>
        <th align="left" nowrap colspan="2">Exemptions by Products &amp; Categories</th>
        <td></td>
    </tr>
    <tr> 
	    <td colspan="3" class="pcCPspacer"></td>
    </tr>
<%
	For i=0 to intCount
	Count=Count+1
	tmpStr4=""
	tmpStr5=""
	
	tmpA=split(pcArray(2,i),",")
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
	
	tmpB=split(pcArray(3,i),",")
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
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
            <td valign="top" nowrap>
            <input type="hidden" name="pcv_ID1_<%=Count%>" value="<%=pcArray(0,i)%>">
            <input type="hidden" name="pcv_ID2_<%=Count%>" value="<%=pcArray(2,i)%>">
            <input type="hidden" name="pcv_ID3_<%=Count%>" value="<%=pcArray(3,i)%>">
            <input type="checkbox" name="C<%=Count%>" value="1" class="clearBorder">
            <a href="editTaxEpt.asp?stateCode=<%=pcArray(0,i)%>"><%=pcArray(1,i)%></a>
            </td>
            <td valign="top">
            <%if tmpStr4<>"" then%>
            <strong>Tax-exempt products</strong>:<br>
            <%=tmpStr4%>
            <%end if%>
            <%if tmpStr5<>"" then%>
            <strong>Tax-exempt categories</strong>:<br>
            <%=tmpStr5%>
            <%end if%>
            </td>
            <td valign="top" align="right"><a href="editTaxEpt.asp?stateCode=<%=pcArray(0,i)%>"><img src="images/pcIconGo.jpg" width="12" height="12" alt="Edit" title="Edit these exemptions"></a></td>
        </tr>
		<%Next%>
        <%end if%>
        <%if Count<>0 then%>
        <tr> 
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="3" class="cpLinksList">
                <a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a>
            </td>
        </tr>
        <%end if%>
        <tr> 
            <td colspan="3" class="pcCPspacer"></td>
        </tr>
        <tr> 
            <td colspan="3">
            <%if Count<>0 then%>
            <input type="submit" name="submit2" value="Delete Selected" onclick="javascript:if (confirm('You are about to remove selected tax rules from your database. Are you sure you want to complete this action?')) {return(true)} else {return(false)}" class="submit2">&nbsp;
            <%end if%>
            <input type="button" name="addnew" value="Add New Tax Exemption Rule" onclick="location='AddTaxEpt.asp';">&nbsp;
            
            <% 
			if ptaxfile=1 then
				pcvBackURL = "AdminTaxsettings.asp?nofile=0"
			else
				pcvBackURL = "viewTax.asp"
			end if
			%>
            <input type="button" name="back" value="Manage Taxes" onClick="document.location.href='<%=pcvBackURL%>';">
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