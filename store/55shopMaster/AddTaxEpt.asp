<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Response.Buffer = true %>
<% pageTitle="Add State-specific Tax Exemption" %>
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

msg=request.querystring("msg")

ErrorStr=""
ItemName=""
MsgBack=""
ErrorArr=""

call opendb()

Function RepStr(scStr,iditem)
Dim tmp1,tmp2,t1,t2
tmp1=scStr
t1=Instr(tmp1,"value=""" & iditem & """>")
if t1>0 then
	t2=Instr(t1,tmp1,"</option>")+9

	tmp2="<option " & mid(tmp1,t1,t2-t1) & vbcrlf
	tmp1=replace(tmp1,tmp2,"")
end if

RepStr=tmp1

End Function

if request("action")="add" then
	UpdSucess=0
	Count=request.form("Count")
	if Count>"0" then

	For i=1 to Count
		ErrorArr=""
		tmp1=request.form("pcv_Rule1_" & i)
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
		if (tmp1<>"") and ((tmp2<>"") or (tmp3<>"")) then
			query="INSERT INTO pcTaxEpt (pcTEpt_StateCode,pcTEpt_ProductList,pcTEpt_CategoryList) values ('" & tmp1 & "','" & tmp2 & "','" & tmp3 & "')"
			set rs=connTemp.execute(query)
			set rs=nothing
			UpdSucess=UpdSucess+1
		else
		if not (tmp1&tmp2&tmp3="") then
			ErrorStr="Cannot create some tax rules because they are not logical."
			MsgBack=MsgBack & "&Rule1_" & i & "=" & tmp1 & "&Rule2_" & i & "=" & Server.URLEncode(tmp2) & "&Rule3_" & i & "=" & Server.URLEncode(tmp3)
		end if
		end if
	Next
	if UpdSucess>0 then
		if ErrorStr="" then
			msg="Tax Rules were added successfully!"
		else
			msg="All others rules were added successfully!"
		end if
	end if
	end if
	
	call closeDb()
	response.redirect "AddTaxEpt.asp?msg=" & Server.UrlEncode(msg) & "&ErrorStr=" & Server.UrlEncode(ErrorStr) & MsgBack
end if
	
'// Show Messages - START
if request.querystring("ErrorStr")<>"" then %>
<div class="pcCPmessage"><%=request.querystring("ErrorStr")%></div>
<% 
end if

if msg<>"" then %>
<div class="pcCPmessageSuccess"><%=msg%></div>
<% 
end if
'// Show Messages - END

Function getStateList()
Dim tmp1,rs1
tmp1=""
query="SELECT stateCode,stateName FROM states WHERE NOT (stateName like '%Canada%') order by stateCode;"
set rs1=Server.CreateObject("ADODB.Recordset")
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
set rs1=Server.CreateObject("ADODB.Recordset")
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
<form method="POST" action="AddTaxEpt.asp?action=add" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="2">To select more than one item, keep the CTRL key pressed down.</td>
    </tr>

<%
Count=1
tmpStr=""
tmpStr1=""
tmpStr=getStateList()
tmpStr1=tmpStr

query="SELECT pcTEpt_StateCode FROM pcTaxEpt group by pcTEpt_StateCode;"
set rs1=Server.CreateObject("ADODB.Recordset")
set rs1=connTemp.execute(query)

if not rs1.eof then
	pcArray3=rs1.GetRows()
	intCount3=ubound(pcArray3,2)
	For m=0 to intCount3
	if trim(pcArray3(0,m))<>"" then
		tmpStr1=RepStr(tmpStr1,pcArray3(0,m))
	end if
	Next
end if
set rs1=nothing

tmpStr2=""
tmpStr2=getItemList()

%>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th align="left" nowrap>State/Province</th>
        <th align="left" nowrap>Exemptions by Products &amp; Categories</th>
    </tr>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
<%
	tmpStr4=tmpStr1
	if request.querystring("Rule1_" & Count)<>"" then
		tmpStr4=replace(tmpStr4,"value=""" & request.querystring("Rule1_" & Count) & """>","value=""" & request.querystring("Rule1_" & Count) & """ selected>")
	end if
	tmpStr5=tmpStr2
	if request.querystring("Rule2_" & Count)<>"" then
		tmpA=split(request.querystring("Rule2_" & Count),",")
		For j=lbound(tmpA) to ubound(tmpA)
		if trim(tmpA(j))<>"" then
			tmpStr5=replace(tmpStr5,"value=""" & tmpA(j) & """>","value=""" & tmpA(j) & """ selected>")
		end if
		Next
	end if
%>
    <tr>
        <td valign="top">
        <select name="pcv_Rule1_<%=Count%>">
        <option value="" selected>-- No Selection --</option>
        <%=tmpStr4%>
        </select>
        </td>
        <td valign="top">
        <strong>Select tax-exempt products</strong>:
        <br />
        <select name="pcv_Rule2_<%=Count%>" multiple size="5">
        <%=tmpStr5%>
        </select>
        <br /><br />
        <strong>Select tax-exempt categories</strong>:
        <br />
            <%pcv_tmp1=request.querystring("Rule3_" & Count)
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
            <!--#include file="../includes/pcCategoriesList.asp"-->
            <%call pcs_CatList()%>
            <br /><br />
        </td>
    </tr>
    <tr>
    	<td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td colspan="2" align="center">
        <input type="submit" name="submit" value="Add Rules" class="submit2">&nbsp;
        <input type="button" name="goback" value="Manage Existing Tax Exemptions" onclick="location='manageTaxEpt.asp';">
        <input type="hidden" name="Count" value="<%=Count%>">
        </td>
    </tr>                   
</table>
</form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->