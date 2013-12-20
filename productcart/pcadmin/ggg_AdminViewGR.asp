<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Dim pageTitle, Section
pageTitle="Add Gift Registry Products to Order"
Section="orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../pc/ggg_inc_chkEPPrices.asp"-->
<%
dim conntemp

call openDb()

pIDOrder=getUserInput(request("ido"),0)
gIDEvent=request("IDEvent")
gOrder=getUserInput(request("gOrder"),0)
gSort=getUserInput(request("gSort"),0)
if gOrder="" then
	gOrder="products.Description"
end if
if gSort="" then
	gSort="ASC"
end if

mOrder=" ORDER by " & gOrder & " " & gSort

query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_IDEvent=" & gIDEvent & " and pcEv_Active=1"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)

if rstemp.eof then
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

query="select pcEP_ID,pcEP_IDProduct,pcEP_OptionsArray,pcEP_IDConfig from pcEvProducts where pcEP_IDEvent=" & gIDEvent
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)

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
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing
	rstemp.MoveNext
loop
set rstemp=nothing

'End of Update Product Prices
%>
<!--#include file="AdminHeader.asp"--> 
<% If msg<>"" then %>
<div class="pcCPmessage">
	<%=msg%>
</div>
<% end if %>

<%
query="select products.sku,products.description,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct " & mOrder
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
%>
           
<form method="post" name="Form1" action="ggg_AdminAddEPs.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
<tr>
	<td colspan="7">
		<h2>"<%=geName%>"<%response.write dictLanguage.Item(Session("language")&"_viewGR_1")%></h2>
		<strong><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1c")%></strong><%=geName%>&nbsp;|&nbsp;<strong><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1b")%></strong><%=geDate%><%if gType<>"" then%>&nbsp;|&nbsp;<strong><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1d")%></strong><%=gType%>
		<%end if%>
	</td>
</tr>			
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
    	<td colspan="7" class="pcCPspacer"></td>
    </tr>
	<tr>
		<th nowrap align="left">
			<b><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=products.sku&gSort=asc"><img src="../pc/images/sortasc.gif" border="0"></a><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=products.sku&gSort=desc"><img src="../pc/images/sortdesc.gif" border="0"></a>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_viewGR_2")%></b>
		</th>
		<th nowrap align="left">
			<b><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=products.description&gSort=asc"><img src="../pc/images/sortasc.gif" border="0"></a><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=products.description&gSort=desc"><img src="../pc/images/sortdesc.gif" border="0"></a>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_viewGR_3")%></b>
		</th>
		<th nowrap align="left"></th>
		<th nowrap align="right">
			<a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_Price&gSort=asc"><img src="../pc/images/sortasc.gif" border="0"></a><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_Price&gSort=desc"><img src="../pc/images/sortdesc.gif" border="0"></a>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_viewGR_4")%>
		</th>
		<th nowrap align="right">
			<a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_Qty&gSort=asc"><img src="../pc/images/sortasc.gif" border="0"></a><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_Qty&gSort=desc"><img src="../pc/images/sortdesc.gif" border="0"></a>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_viewGR_5")%>
		</th>
		<th nowrap align="right">
			<a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_HQty&gSort=asc"><img src="../pc/images/sortasc.gif" border="0"></a><a href="ggg_AdminViewGR.asp?IDEvent=<%=gIDEvent%>&gOrder=pcEvProducts.pcEP_HQty&gSort=desc"><img src="../pc/images/sortdesc.gif" border="0"></a>&nbsp;<%response.write dictLanguage.Item(Session("language")&"_viewGR_6")%>
		</th>	
		<th nowrap align="right" width="100">
			<%response.write dictLanguage.Item(Session("language")&"_viewGR_12")%>
		</th>
	</tr>
    <tr>
    	<td colspan="7" class="pcCPspacer"></td>
    </tr>
	<%
	Count=0
	do while not rstemp.eof
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
		if gGC<>"1" then
		Count=Count+1%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td nowrap><%=gsku%></td>
			<td nowrap><a href="ggg_AdminViewEP.asp?IDEvent=<%=gIDEvent%>&geID=<%=geID%>"><%=gname%></a></td>
			<td><a href="javascript:win('../pc/catalog/<%=gtimage%>');"><img src="../pc/catalog/<%=gtimage%>" width="35" height="35" border="0"></a></td>
			<td nowrap align="center"><%=scCurSign & money(gPrice)%></td>
			<td nowrap align="center"><%=gQty%></td>
			<td nowrap align="center"><%=clng(gHQty)%></td>
			<td align="center" nowrap>
				<%if clng(gQty)-clng(gHQty)<=0 then%>
					<input name="add<%=Count%>" value="0" type=hidden>
					<%response.write dictLanguage.Item(Session("language")&"_viewGR_7")%>
				<%else
					if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") or ((gservice=-1) and (iBTOOutofStockPurchase=0)) or (scOutofstockpurchase=0) then%>
						<input name="add<%=Count%>" value="0" type=text size="4" style="float: right; text-align:right">
					<%else%>
						<input name="add<%=Count%>" value="0" type=hidden>
						<%response.write dictLanguage.Item(Session("language")&"_viewGR_8")%>
					<%end if
				end if%>
				<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
				<input name="remain<%=Count%>" value="<%=clng(gQty)-clng(gHQty)%>" type=hidden>
			</td>
		</tr>
		<%
		end if
		rstemp.MoveNext
	loop
	set rstemp=nothing
		
	Count1=Count
		
	query="select products.sku,products.description,products.smallImageUrl,products.stock,products.nostock,products.ServiceSpec,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_Price from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_GC=1 and products.removed=0 " & mOrder
	set rstemp=connTemp.execute(query)
		
	IF NOT rstemp.eof then%>
		<tr>
			<td colspan="7">
				<b><%response.write dictLanguage.Item(Session("language")&"_GRDetails_8")%></b>
			</td>
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
				<td nowrap><a href="ggg_AdminViewEP.asp?IDEvent=<%=gIDEvent%>&geID=<%=geID%>"><%=gname%></a></td>
				<td><a href="javascript:win('../pc/catalog/<%=gtimage%>');"><img src="../pc/catalog/<%=gtimage%>" width="35" height="35" border="0"></a></td>
				<td nowrap align="center"><%=scCurSign & money(gPrice)%></td>
				<td nowrap align="center"></td>
				<td nowrap align="center"></td>
				<td align="center" nowrap>
					<%if (clng(gQty)-clng(gHQty)<=clng(gstock)) or (gnostock<>"0") or (scOutofstockpurchase=0) then%>
						<input name="add<%=Count%>" type=text value="0" size="4" style="float: right; text-align:right">
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
	END IF 'Have GCs%>
	<tr> 
		<td colspan="7"> 
			<p>
				<br><br>
				<input type="submit" name="submit" value=" Add Products to Order " class="submit2">&nbsp;
				<input type=button name=back value=" Back " onclick="javascript:history.go(-1);" class="iBtnGrey">
				<input type=hidden name="ido" value="<%=pIDOrder%>">
				<input type=hidden name="Count" value="<%=Count%>">
				<input type=hidden name="IDEvent" value="<%=gIDEvent%>">
				<br>
			</p>
		</td>
	</tr>
<%END IF 'Have products
set rstemp=nothing%>
</table>
</form>
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
<%For k=1 to Count%>
	if (theForm.add<%=k%>.value != "")
  	{
		if (allDigit(theForm.add<%=k%>.value) == false)
		{
			alert("Please enter a valid number for this field.");
			theForm.add<%=k%>.focus();
	    return (false);
		}
		
		if (eval(theForm.add<%=k%>.value) > eval(theForm.remain<%=k%>.value))
		{
			alert("Your entered a quantity greater than remaining quantity.");
			theForm.add<%=k%>.focus();
	    return (false);
		}
	}
<%Next%>

return (true);
}
//-->
</script>
<%call closedb()%><!--#include file="adminFooter.asp"-->