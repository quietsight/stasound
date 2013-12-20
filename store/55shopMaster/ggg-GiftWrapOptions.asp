<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Gift Wrapping Options" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
Dim connTemp,query,rstemp

call opendb()

if request("submit4")<>"" then
	pcv_intGGGShow=request("pcv_intGGGShow")
	if pcv_intGGGShow="" then
		pcv_intGGGShow="0"
	end if
	pcv_intGGGshowtextCart=request("pcv_intGGGshowtextCart")
	if pcv_intGGGshowtextCart="" then
		pcv_intGGGshowtextCart="0"
	end if
	pcv_intGGGshowtext=request("pcv_intGGGshowtext")
	if pcv_intGGGshowtext="" then
		pcv_intGGGshowtext="0"
	end if
	if pcv_intGGGShow="0" then
		pcv_intGGGshowtext="0"
		pcv_intGGGshowtextCart="0"
	end if
	pcv_details=pcf_ReplaceCharacters(request("details"))
	pcv_detailsCart=pcf_ReplaceCharacters(request("detailsCart"))
	
	query="select pcGWSet_ID from pcGWSettings"
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	if rstemp.eof then
		query="INSERT INTO pcGWSettings (pcGWSet_Show,pcGWSet_Overview,pcGWSet_HTML,pcGWSet_OverviewCart,pcGWSet_HTMLCart) VALUES (" & pcv_intGGGShow & "," & pcv_intGGGshowtext & ",'" & pcv_details & "'," & pcv_intGGGshowtextCart & ",'" & pcv_detailsCart & "')"
		set rstemp=connTemp.execute(query)
	else
		query="UPDATE pcGWSettings set pcGWSet_Show=" & pcv_intGGGShow & ",pcGWSet_Overview=" & pcv_intGGGshowtext & ",pcGWSet_HTML='" & pcv_details & "',pcGWSet_OverviewCart=" & pcv_intGGGshowtextCart & ",pcGWSet_HTMLCart='" & pcv_detailsCart & "';"
		set rstemp=connTemp.execute(query)
	end if
	set rstemp=nothing
end if

if request("submit1")<>"" then
	Count=request("Count")
	if Count="" then
		Count="0"
	end if
	
	For i=1 to Count
		if request("Opt" & i)="1" then
			OptName=request("OptName" & i)
			if OptName<>"" then
				OptName=replace(OptName,"'","''")
			end if	
			OptImg=request("OptImg" & i)
			OptPrice=replacecomma(request("OptPrice" & i))
			OptActive=request("OptActive" & i)
			if OptActive="" then
				OptActive=0
			end if
			OptOrder=request("OptOrder" & i)
			if OptOrder="" then
				OptOrder=0
			end if
			IDOpt=request("IDOpt" & i)
			query="update pcGWOptions set pcGW_OptName='" & OptName & "',pcGW_OptImg='" & OptImg & "',pcGW_OptPrice=" & OptPrice & ",pcGW_OptActive=" & OptActive & ",pcGW_OptOrder=" & OptOrder & " where pcGW_IDOpt=" & IDOpt
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
		end if
	Next
end if
		
if request("submit2")<>"" then
	Count=request("Count")
	if Count="" then
		Count="0"
	end if
	For i=1 to Count
		if request("Opt" & i)="1" then
			IDOpt=request("IDOpt" & i)
			query="update pcGWOptions set pcGW_Removed=1 where pcGW_IDOpt=" & IDOpt
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
		end if
	Next
end if

if request("submit3")<>"" then
	Count1=request("Count1")
	if Count1="" then
		Count1="0"
	end if
	For i=1 to Count1
		if request("Pro" & i)="1" then
			IDPro=request("IDPro" & i)
			query="delete from pcProductsExc where pcPE_IDProduct=" & IDPro
			set rstemp=connTemp.execute(query)
			set rstemp=nothing
		end if
	Next
end if
	
query="SELECT pcGWSet_Show,pcGWSet_Overview,pcGWSet_OverviewCart,pcGWSet_HTML,pcGWSet_HTMLCart FROM pcGWSettings"
set rstemp=connTemp.execute(query)
if not rstemp.eof then
	pcv_intGGGShow=rstemp("pcGWSet_Show")
	if pcv_intGGGShow="" then
		pcv_intGGGShow="0"
	end if
	pcv_intGGGshowtext=rstemp("pcGWSet_Overview")
	if pcv_intGGGshowtext="" then
		pcv_intGGGshowtext="0"
	end if
	pcv_intGGGshowtextCart=rstemp("pcGWSet_OverviewCart")
	if pcv_intGGGshowtextCart="" then
		pcv_intGGGshowtextCart="0"
	end if
	pcv_details=rstemp("pcGWSet_HTML")
	pcv_detailsCart=rstemp("pcGWSet_HTMLCart")
	pcv_details=pcf_PrintCharacters(pcv_details)
	pcv_detailsCart=pcf_PrintCharacters(pcv_detailsCart)
else
	pcv_intGGGShow="0"
	pcv_intGGGshowtext="0"
	pcv_intGGGshowtextCart="0"
	pcv_details=""
end if
set rstemp=nothing
%>

<script language="JavaScript">
<!--
function newWindow(file,window)
{
	msgWindow=open(file,window,'resizable=no,width=400,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}

function chgWin(file,window)
{
    msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
    if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>
<form name="hForm" method="post" action="ggg-GiftWrapOptions.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th colspan="2">General Settings</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td width="50%"><p>Show Gift Wrapping options during checkout?</p></td>
		<td width="50%">
			<input type="radio" name="pcv_intGGGShow" value="1" <%if pcv_intGGGShow="1" then%>checked<%end if%> class="clearBorder"> Yes 
			<input type="radio" name="pcv_intGGGShow" value="0" <%if pcv_intGGGShow<>"1" then%>checked<%end if%> class="clearBorder"> No
		</td>
	</tr>
	<tr> 
		<td><p>Show Gift Wrapping &quot;Overview&quot; on the shopping cart page?</p></td>
		<td>
			<input type="radio" name="pcv_intGGGshowtextCart" value="1" <%if pcv_intGGGshowtextCart="1" then%>checked<%end if%> class="clearBorder"> Yes 
			<input type="radio" name="pcv_intGGGshowtextCart" value="0" <%if pcv_intGGGshowtextCart<>"1" then%>checked<%end if%> class="clearBorder"> No
		</td>
	</tr>
	<tr> 
		<td><p>Show Gift Wrapping &quot;Instructions&quot; during checkout?</p></td>
		<td>
			<input type="radio" name="pcv_intGGGshowtext" value="1" <%if pcv_intGGGshowtext="1" then%>checked<%end if%> class="clearBorder"> Yes 
			<input type="radio" name="pcv_intGGGshowtext" value="0" <%if pcv_intGGGshowtext<>"1" then%>checked<%end if%> class="clearBorder"> No
		</td>
	</tr>
	<tr> 
		<td colspan="2">
			<p>Message to show on the shopping cart page:</p>
		</td>
    </tr>
    	<td colspan="2">
			<textarea name="detailsCart" cols="80" rows="5"><%=pcv_detailsCart%></textarea>&nbsp;
			<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=detailsCart','window2')">
		</td>
	</tr>
	<tr> 
		<td colspan="2">
			<p>Message to show on the checkout page:</p>
		</td>
    </tr>
    	<td colspan="2">
			<textarea name="details" cols="80" rows="5"><%=pcv_details%></textarea>&nbsp;
			<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=details','window2')">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td colspan="2">
			<input name="submit4" type="submit" class="submit2" value="Update Settings">
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Available Gift Wrapping Options</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
	<td colspan="2" valign="top">
		<table class="pcCPcontent" style="border: 1px dashed #CCCCCC;">
		<tr>
			<td><strong>Name</strong></td>
			<td nowrap="nowrap"><strong>Image</strong> (optional)</td>
			<td><strong>Price</strong></td>
			<td><strong>Active</strong></td>
			<td><strong>Order</strong></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="6" class="pcCPspacer"></td>
		</tr>
		<%query="SELECT pcGW_IDOpt,pcGW_OptName,pcGW_OptImg,pcGW_OptPrice,pcGW_OptActive,pcGW_OptOrder FROM pcGWOptions WHERE pcGW_Removed=0 ORDER BY pcGW_OptOrder ASC,pcGW_OptName ASC;"
		set rstemp=connTemp.execute(query)
		Count=0
		if rstemp.eof then%>
		<tr>
			<td colspan="6" style="padding: 10px;">No Items to display.</td>
		</tr>
		<%else
			do while not rstemp.eof
				Count=Count+1
				pcv_IDOpt=rstemp("pcGW_IDOpt")
				pcv_OptName=rstemp("pcGW_OptName")
				pcv_OptImg=rstemp("pcGW_OptImg")
				pcv_OptPrice=rstemp("pcGW_OptPrice")
				pcv_OptActive=rstemp("pcGW_OptActive")
				pcv_OptOrder=rstemp("pcGW_OptOrder")
				if pcv_OptOrder="" then
					pcv_OptOrder=0
				end if
				%>
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td><input name="OptName<%=Count%>" type=text value="<%=pcv_OptName%>"></td>
					<td><input name="OptImg<%=Count%>" type=text value="<%=pcv_OptImg%>"><a href="javascript:chgWin('../pc/imageDir.asp?ffid=OptImg<%=Count%>&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>&nbsp;<a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')"><img src="images/sortasc_blue.gif" alt="Upload Image"></a></td>
					<td><input name="OptPrice<%=Count%>" type="text" size="8" value="<%=money(pcv_OptPrice)%>"></td>
					<td><input name="OptActive<%=Count%>" type="checkbox" value="1" <%if pcv_OptActive<>"0" then%>checked<%end if%> class="clearBorder"></td>
					<td><input name="OptOrder<%=Count%>" type="text" size="4" value="<%=pcv_OptOrder%>"></td>
					<td>
						<input name="Opt<%=Count%>" type="checkbox" value="1" class="clearBorder">
						<input type=hidden name="IDOpt<%=Count%>" value="<%=pcv_IDOpt%>">
					</td>
				</tr>
				<%rstemp.MoveNext
			loop
		end if
		set rstemp=nothing%>
        <tr>
        	<td colspan="6" class="pcCPspacer"></td>
        </tr>
		<tr>
			<td colspan="6" class="cpLinksList" align="right">
			<%if Count>0 then%>
				<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
				<script language="JavaScript">
				<!--
				function checkAll()
				{
					for (var j = 1; j <= <%=count%>; j++) {
						box = eval("document.hForm.Opt" + j); 
						if (box.checked == false) box.checked = true;
					}
				}
	
				function uncheckAll()
				{
					for (var j = 1; j <= <%=count%>; j++) {
						box = eval("document.hForm.Opt" + j); 
						if (box.checked == true) box.checked = false;
					}
				}

				//-->
				</script>
			<%end if%>
			</td>
		</tr>
		<tr>
			<td colspan="6" style="padding: 10px;">
				<%if Count>0 then%>
					<input name="submit1" type="submit" value="Updated Selected" class="submit2">
					&nbsp;
					<input name="submit2" type="submit" value="Remove Selected" class="submit2">
					&nbsp;
				<%end if%>
				<input name="New" type="button" value="Add New" onclick="location='ggg_AddGWOpt.asp';">
				<input name="Count" type=hidden value="<%=Count%>">
			</td>
		</tr>
		</table>
	</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Product exclusions</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
	<td colspan="2" valign="top">
		<table class="pcCPcontent" style="border: 1px dashed #CCCCCC;">
		<tr>
			<td colspan="2" style="padding: 10px;">If no products are selected, all products can be gift wrapped.</td>
		</tr>
		<%
		query="select products.idproduct,products.description from products,pcProductsExc where products.idproduct=pcProductsExc.pcPE_IDProduct order by products.description"
		set rstemp=connTemp.execute(query)
		Count1=0
		do while not rstemp.eof
			Count1=Count1+1
			pIDProduct=rstemp("IDProduct")
			pName=rstemp("description")%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td style="padding-left: 10px;" width="95%"><a href="FindProductType.asp?id=<%=pIDProduct%>" target="_blank"><%=pName%></a></td>
				<td style="padding-right: 10px;" width="5%" align="left">
					<input type="checkbox" name="Pro<%=Count1%>" value="1" class="clearBorder">
					<input type="hidden" name="IDPro<%=Count1%>" value="<%=pIDProduct%>">
				</td>
			</tr>
			<%
			rstemp.MoveNext
		loop
		set rstemp=nothing
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
		<td colspan="2" class="cpLinksList" align="right" style="padding-right: 10px;">
			<%if Count1>0 then%>
				<a href="javascript:checkAllPrd();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAllPrd();">Uncheck All</a>
				<script language="JavaScript">
				<!--
				function checkAllPrd()
				{
					for (var j = 1; j <= <%=count1%>; j++) {
						box = eval("document.hForm.Pro" + j); 
						if (box.checked == false) box.checked = true;
					}
				}

				function uncheckAllPrd()
				{
					for (var j = 1; j <= <%=count1%>; j++) {
						box = eval("document.hForm.Pro" + j); 
						if (box.checked == true) box.checked = false;
					}
				}
				//-->
				</script>
			<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2" style="padding: 10px;">
			<%if Count1>0 then%>
				<input name="submit3" type="submit" value="Remove Selected" class="submit2">
				&nbsp;
			<%end if%>
			<input name="newpro" type="button" value="Add New" onclick="location='ggg_addPrdExc.asp';">
			<input type="hidden" name="Count1" value="<%=Count1%>">
		</td>
	</tr>
	</table>
	</td>
	</tr>
</table>
</form>
<script language="JavaScript">
<!--

function isDigit(s)
{
var test=""+s;
if(test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
	if (theForm.OptPrice<%=k%>.value != "")
  	{
	if (allDigit(theForm.OptPrice<%=k%>.value) == false)
	{
		alert("Please enter a valid number for this field.");
		theForm.OptPrice<%=k%>.focus();
    return (false);
	}
	}
	if (theForm.OptPrice<%=k%>.value == "")
  	{
  	theForm.OptPrice<%=k%>.value="0";
  	}
<%Next%>

return (true);
}
//-->
</script>
<%
call closedb()
%>
<!--#include file="AdminFooter.asp"-->