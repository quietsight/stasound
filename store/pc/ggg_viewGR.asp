<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/openDb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="ggg_inc_chkEPPrices.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->

<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If
%>
<!--#include file="pcStartSession.asp"-->
<%
dim conntemp

call openDb()

grCode=getUserInput(request("grCode"),0)
gOrder=getUserInput(request("gOrder"),0)
gSort=getUserInput(request("gSort"),0)
if gOrder="" then
	gOrder="products.Description"
end if
if gSort="" then
	gSort="ASC"
end if

mOrder=" ORDER by " & gOrder & " " & gSort

if grCode="" then
	response.redirect "msg.asp?message=98"
end if

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_Code='" & grCode & "' and pcEv_Active=1;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if rstemp.eof then
	response.redirect "msg.asp?message=98"
else
	gIDEvent=rstemp("pcEv_IDEvent")
	geName=rstemp("pcEv_Name")
	geDate=rstemp("pcEv_Date")
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
			gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
			gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if
	gType=rstemp("pcEv_Type")
	if gType<>"" then
	else
		gType=""
	end if
	geincGc=rstemp("pcEv_IncGcs")
	if geincGc<>"" then
	else
		geincGc="0"
	end if
end if

set rstemp=nothing

'Update Products Price
query="SELECT products.idproduct, pcEvProducts.pcEP_ID, pcEvProducts.pcEP_IDProduct, pcEvProducts.pcEP_OptionsArray, pcEvProducts.pcEP_IDConfig FROM products, pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND products.active<>0 ORDER BY products.Description ASC, pcEvProducts.pcEP_GC ASC"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
do while not rstemp.eof
	geID=rstemp("pcEP_ID")
	gIDProduct=rstemp("pcEP_IDProduct")
	
	pcv_strOptionsArray=rstemp("pcEP_OptionsArray")

	gIDConfig=rstemp("pcEP_IDConfig")
	if gIDConfig<>"" then
	else
		gIDConfig="0"
	end if

	gnewPrice=updPrices(gIDProduct,gIDConfig,pcv_strOptionsArray,0)

	query="update pcEvProducts set pcEP_Price=" & gnewPrice & " where pcEP_ID=" & geID
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rs=nothing

	rstemp.MoveNext
loop

set rstemp=nothing

'End of Update Product Prices

%>
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<%query="SELECT products.sku,products.description,products.imageUrl,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price,pcEP_OptionsArray FROM products,pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND products.active=-1 " & mOrder
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if%>
<div id="pcMain">
<form method="post" name="Form1" action="ggg_addEPtocart.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcMainTable">
<tr>
	<td colspan="7">
		<h1>"<%=geName%>"<%response.write dictLanguage.Item(Session("language")&"_viewGR_1")%></h1>
		<br>
		<br><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1c")%><%=geName%>
		<br><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1b")%><%=geDate%>
		<%if gType<>"" then%>
		<br><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1d")%><%=gType%>
		<%end if%>
	</td>
</tr>
<% If msg<>"" then %>
<tr> 
	<td colspan="7"><div class="pcErrorMessage"><%=msg%></div></td>
</tr>
<% end if %>			
	
<%IF rstemp.eof then%>
	<tr>
		<td colspan="7">
			<br>
			<p> 
				<%response.write dictLanguage.Item(Session("language")&"_viewGR_11")%>
			</p>
			<br>
			<a href="javascript:history.go(-1);"><img src="<%=rslayout("back")%>" border=0></a>
			<br>
		</td>
	</tr>
<%ELSE%>
    <tr>
        <td colspan="7" class="pcSpacer"></td>
    </tr>
	<tr>
		<th nowrap>
			<div style="text-align: left;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_2")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.sku&gSort=asc"><img src="images/sortasc.gif" border="0"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.sku&gSort=desc"><img src="images/sortdesc.gif" border="0"></a>
        	</div>    
        </th>
		<th></th>
		<th nowrap width="50%">
			<div style="text-align: left;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_3")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.description&gSort=asc"><img src="images/sortasc.gif" border="0"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=products.description&gSort=desc"><img src="images/sortdesc.gif" border="0"></a>
            </div>
        </th>
		<th nowrap>
			<div style="text-align: center;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_4")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Price&gSort=asc"><img src="images/sortasc.gif" border="0"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Price&gSort=desc"><img src="images/sortdesc.gif" border="0"></a>
            </div>
        </th>
		<th nowrap>
			<div style="text-align: center;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_5")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Qty&gSort=asc"><img src="images/sortasc.gif" border="0"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_Qty&gSort=desc"><img src="images/sortdesc.gif" border="0"></a>
            </div>
        </th>
		<th nowrap>
			<div style="text-align: center;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_6")%>&nbsp;<a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_HQty&gSort=asc"><img src="images/sortasc.gif" border="0"></a><a href="ggg_viewGR.asp?grCode=<%=grCode%>&gOrder=pcEvProducts.pcEP_HQty&gSort=desc"><img src="images/sortdesc.gif" border="0"></a>
            </div>
        </th>	
		<th nowrap>
			<div style="text-align: center;">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_12")%>
			</div>
		</th>
	</tr>
    <tr>
        <td colspan="7" class="pcSpacer"></td>
    </tr>
	<%
	Count=0
	ExList="**"
	LowList="**"
	LowName="**"
	LowQStock="**"
	do while not rstemp.eof
		gsku=rstemp("sku")
		gname=rstemp("description")

		'// Find product image
		gtimage1=rstemp("ImageUrl")
		gtimage2=rstemp("smallImageUrl")
		if gtimage2<>"" then
			gtimage=gtimage2
		else
			gtimage=gtimage1
		end if
		if trim(gtimage)="" then
			gtimage="no_image.gif"
		end if

		gstock=rstemp("stock")
		if gstock<>"" then
		else
			gstock="0"
		end if
		gnostock=rstemp("nostock")
		if gnostock<>"" then
		else
			gnostock="0"
		end if
		gservice=rstemp("ServiceSpec")
		if gservice<>"" then
		else
			gservice="0"
		end if
		geID=rstemp("pcEP_ID")
		gQty=rstemp("pcEP_Qty")
		if gQty<>"" then
		else
			gQty="0"
		end if
		gHQty=rstemp("pcEP_HQty")
		if gHQty<>"" then
		else
			gHQty="0"
		end if
		gGC=rstemp("pcEP_GC")
		gPrice=rstemp("pcEP_Price")
		if gPrice<>"" then
		else
			gPrice="0"
		end if
		
		pcv_strSelectedOptions=""
		pcv_strSelectedOptions = rstemp("pcEP_OptionsArray")
		pcv_strSelectedOptions=pcv_strSelectedOptions&""
		
		if gGC<>"1" then
		Count=Count+1%>
		<tr valign="top"> 
		<td nowrap valign="top"><%=gsku%></td>
		<% if gtimage<>"no_image.gif" then %>
        <td align="left">
            <div style="text-align: left; padding: 5px;"><img src="catalog/<%=gtimage%>" width="50" height="50" border="0"></div>
        </td>
        <% end if %>
		<td <% if gtimage="no_image.gif" then %> colspan="2"<%end if%> valign="top"><a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=geID%>"><%=gname%></a>
		<%Dim pPriceToAdd, pOptionDescrip, pOptionGroupDesc
		Dim pcArray_SelectedOptions, pcv_strOptionsArray, cCounter, xOptionsArrayCount
		Dim pcv_strOptionsPriceArray, pcv_strOptionsPriceArrayCur, pcv_strOptionsPriceTotal
	
		IF len(pcv_strSelectedOptions)>0 AND pcv_strSelectedOptions<>"NULL" THEN
	
		pcArray_SelectedOptions = Split(pcv_strSelectedOptions,chr(124))
		
		pcv_strOptionsArray = ""
		pcv_strOptionsPriceArray = ""
		pcv_strOptionsPriceArrayCur = ""
		pcv_strOptionsPriceTotal = 0
		xOptionsArrayCount = 0
		
		For cCounter = LBound(pcArray_SelectedOptions) TO UBound(pcArray_SelectedOptions)
			
			' SELECT DATA SET
			' TABLES: optionsGroups, options, options_optionsGroups
			query = 		"SELECT optionsGroups.optionGroupDesc, options.optionDescrip, options_optionsGroups.price, options_optionsGroups.Wprice "
			query = query & "FROM optionsGroups, options, options_optionsGroups "
			query = query & "WHERE idoptoptgrp=" & pcArray_SelectedOptions(cCounter) & " "
			query = query & "AND options_optionsGroups.idOption=options.idoption "
			query = query & "AND options_optionsGroups.idOptionGroup=optionsGroups.idoptiongroup "	
			
			set rsQ=server.CreateObject("ADODB.RecordSet")
			set rsQ=conntemp.execute(query)
			if err.number<>0 then
				'//Logs error to the database
				call LogErrorToDatabase()
				'//clear any objects
				set rs=nothing
				'//close any connections
				call closedb()
				'//redirect to error page
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if					
			
			if Not rsQ.eof then 
				
				xOptionsArrayCount = xOptionsArrayCount + 1
				
				pOptionDescrip=""
				pOptionGroupDesc=""
				pPriceToAdd=""
				pOptionDescrip=rsQ("optiondescrip")
				pOptionGroupDesc=rsQ("optionGroupDesc")
				
				If Session("customerType")=1 Then
					pPriceToAdd=rsQ("Wprice")
					If rsQ("Wprice")=0 then
						pPriceToAdd=rsQ("price")
					End If
				Else
					pPriceToAdd=rsQ("price")
				End If	
				
				'// Generate Our Strings
				if xOptionsArrayCount = 1 then%>
				<br>
				<%end if%>
				&nbsp;&nbsp;<%= pOptionGroupDesc & ": " & pOptionDescrip%>
				<%
				if pPriceToAdd="" or pPriceToAdd="0" then
					response.write "&nbsp;"
				else 
					response.write " (" & scCurSign&money(pPriceToAdd) & ")"
				end if%>
				<br>
			<%end if
			set rsQ=nothing
		Next
		
		END IF%>
		</td>
			<td nowrap align="center" valign="top"><%=scCurSign & money(gPrice)%></td>
			<td nowrap align="center" valign="top"><%=gQty%></td>
			<td nowrap align="center" valign="top"><%=clng(gHQty)%></td>
			<td align="right" nowrap valign="top">
				<% if clng(gQty)-clng(gHQty)<=0 then
				ExList=ExList & Count & "**"%>
					<input name="add<%=Count%>" value="0" type=hidden>
					<% '// Fullfilled
					response.write dictLanguage.Item(Session("language")&"_viewGR_7")%>
				<% else
					if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") or ((gservice=-1) and (iBTOOutofStockPurchase=0)) or (scOutofstockpurchase=0) then %>
						<input name="add<%=Count%>" value="0" type=text size="3" style="float: right; text-align:right">
					<% else
						if clng(gstock)>0 then
						LowList=LowList & Count & "**"
						LowName=LowName & gname & "**"
						LowQStock=LowQStock & gstock & "**"%>
						<input name="add<%=Count%>" value="0" type=text size="3" style="float: right; text-align:right"><br>
						<% '// In Stock < Wanted quantity
						response.write dictLanguage.Item(Session("language")&"_viewPrd_19") & gstock%>
						<%else
							ExList=ExList & Count & "**"%>
							<input name="add<%=Count%>" value="0" type=hidden>
							<% '// Out of Stock 
							response.write dictLanguage.Item(Session("language")&"_viewGR_8")%>
						<% end if
					end if
				end if %>
				<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
				<%if clng(gQty)-clng(gHQty)>clng(gstock) AND clng(gstock)>0 then%>
					<input name="remain<%=Count%>" value="<%=gstock%>" type=hidden>
				<%else%>
					<input name="remain<%=Count%>" value="<%if clng(gQty)-clng(gHQty)<0 then%>0<%else%><%=clng(gQty)-clng(gHQty)%><%end if%>" type=hidden>
				<%end if%>
			</td>
		</tr>
		<%end if
		rstemp.MoveNext
	loop
	set rstemp=nothing
		
	Count1=Count
		
	query="select products.sku,products.description,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price from products,pcEvProducts where products.active=-1 and pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_GC=1 and products.removed=0 " & mOrder
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	IF NOT rstemp.eof then%>
        <tr>
            <td colspan="7" class="pcSpacer"></td>
        </tr>
		<tr>
			<th colspan="7"><%response.write dictLanguage.Item(Session("language")&"_GRDetails_8")%></th>
		</tr>
        <tr>
            <td colspan="7" class="pcSpacer"></td>
        </tr>
		<%do while not rstemp.eof
			gsku=rstemp("sku")
			gname=rstemp("description")
			gtimage=rstemp("smallImageUrl")
			if gtimage<>"" then
			else
				gtimage="no_image.gif"
			end if
			gstock=rstemp("stock")
			if gstock<>"" then
			else
				gstock="0"
			end if
			gnostock=rstemp("nostock")
			if gnostock<>"" then
			else
				gnostock="0"
			end if
			gservice=rstemp("ServiceSpec")
			if gservice<>"" then
			else
				gservice="0"
			end if
			geID=rstemp("pcEP_ID")
			gQty=rstemp("pcEP_Qty")
			if gQty<>"" then
			else
				gQty="0"
			end if
			gHQty=rstemp("pcEP_HQty")
			if gHQty<>"" then
			else
				gHQty="0"
			end if
			gGC=rstemp("pcEP_GC")
			gPrice=rstemp("pcEP_Price")
			if gPrice<>"" then
			else
				gPrice="0"
			end if
			if ((gGC="1") and (geincgc="1")) or (clng(gHQty)>0) then
			Count=Count+1%>
			<tr> 
				<td nowrap><%=gsku%></td>
				<td nowrap><a href="ggg_viewEP.asp?grCode=<%=grCode%>&geID=<%=geID%>"><%=gname%></a></td>
				<td><% if gtimage<>"no_image.gif" then %><img src="catalog/<%=gtimage%>" width="50" height="50" border="0"><% end if %></td>
				<td nowrap align="right"><%=scCurSign & money(gPrice)%></td>
				<td nowrap align="right"></td>
				<td nowrap align="right"></td>
				<td align="right" nowrap>
					<%if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") OR (scOutofstockpurchase=0) then%>
						<input name="add<%=Count%>" type=text value="0" size="3" style="float: right; text-align:right">
					<%else%>
						<input name="add<%=Count%>" type=hidden value="0">
						<%response.write dictLanguage.Item(Session("language")&"_viewGR_8")%>
					<%end if%>
					<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
					<input name="remain<%=Count%>" value="99999" type=hidden>
				</td>
			</tr>
			<%end if
			rstemp.MoveNext
		loop
		set rstemp=nothing
	END IF 'Have GCs
	set rstemp=nothing%>
	<tr> 
		<td colspan="7"> 
			<p>
				<br><br>
				<a href="javascript:history.go(-1);"><img src="<%=rslayout("back")%>" border=0></a>&nbsp;
				<input type="image" id="submit" name="submit" value="<%response.write dictLanguage.Item(Session("language")&"_viewGR_9")%>" src="<%=rslayout("addtocart")%>" border="0">
				<input type=hidden name="grCode" value="<%=grCode%>">
				<input type=hidden name="Count" value="<%=Count%>">
				<input type=hidden name="IDEvent" value="<%=gIDEvent%>">
				<br>
			</p>
		</td>
	</tr>
<%END IF 'Have products
%>
</table>
</form>
</div>
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
<%For k=1 to Count
if Instr(ExList,"**" & k & "**")=0 then%>
	if (theForm.add<%=k%>.value != "")
  	{
	if (allDigit(theForm.add<%=k%>.value) == false)
	{
		alert("Please enter a valid number for this field.");
		theForm.add<%=k%>.focus();
    return (false);
	}
	
	<%if Instr(LowList,"**" & k & "**")=0 then%>
	if (eval(theForm.add<%=k%>.value) > eval(theForm.remain<%=k%>.value))
	{
		alert("Your entered a quantity greater than remaining quantity.");
		theForm.add<%=k%>.focus();
    return (false);
	}
	<%else
		tmp1=split(LowList,"**")
		tmp2=split(LowName,"**")
		tmp3=split(LowQStock,"**")
		prdName=""
		prdStock=""
		For l=lbound(tmp1) to ubound(tmp1)
			if tmp1(l)<>"" then
				if clng(tmp1(l))=clng(k) then
					prdName=tmp2(l)
					prdStock=tmp3(l)
					exit for
				end if
			end if
		Next
		%>
		if (eval(theForm.add<%=k%>.value) > eval(theForm.remain<%=k%>.value))
		{
			alert("<%response.write dictLanguage.Item(Session("language")&"_instPrd_2")%><%=prdName%><%response.write dictLanguage.Item(Session("language")&"_instPrd_3")%><%=prdStock%><%response.write dictLanguage.Item(Session("language")&"_instPrd_4")%>");
			theForm.add<%=k%>.focus();
    		return (false);
		}
	<%end if%>
}
<%end if%>
<%Next%>

return (true);
}
//-->
</script>
<%call closedb()%>
<!--#include file="footer.asp"-->