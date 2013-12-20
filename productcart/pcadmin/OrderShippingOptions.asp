<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Order Shipping Options" %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<script type="text/javascript" language="javascript"> 

var destMAX = 8; //maximum number of items in dest list 

function inDest(dest, text, value) { 
var opt, o = 0; 
while (opt = dest[o++]) if (opt.value == value && opt.text == text) return true; 
return o > destMAX; 
} 

function toDest(s, dest) { 
var opt, o = 0; 
while (opt = s[o++]) if (opt.selected && !inDest(dest, opt.text, opt.value)) 
dest.options[dest.length] = new Option(opt.text,opt.value); 
if (navigator.appName == 'Netscape' && 
(navigator.appVersion.indexOf('Win') != -1 || 
navigator.appVersion.indexOf('Mac') != -1)) 
history.go(0); 
} 

//courtesy M. Honnen ///////////// 
//http://www.faqts.com /////////// 
function moveSelected (select, down) { 
	if (select.selectedIndex != -1) { 
		if (down) { 
			if (select.selectedIndex != select.options.length - 1) 
				var x = select.selectedIndex + 1; 
			else 
				return; 
		} 
		else { 
			if (select.selectedIndex != 0) 
				var x = select.selectedIndex - 1; 
			else 
				return; 
		} 
	
		var swapOption = new Object(); 
		swapOption.text = select.options[select.selectedIndex].text; 
		swapOption.value = select.options[select.selectedIndex].value; 
		swapOption.selected = select.options[select.selectedIndex].selected; 
		swapOption.defaultSelected = select.options[select.selectedIndex].defaultSelected; 

		for (var property in swapOption) {
			select.options[select.selectedIndex][property] = select.options[x][property]; 
			select.options[x][property] = swapOption[property]; 			
		}

/*
		for (var property in swapOption) {
			select.options[select.selectedIndex][property] = select.options[x][property]; 
		}

		for (var property in swapOption) {
			select.options[x][property] = swapOption[property]; 
		}
*/
		
	} 
} 

function setHidden(f) { 
var destVals = new Array(), opt = 0, separator = '|', d = f.dest; 
while (d[opt]) destVals[opt] = d[opt++].value; 
f.destItems.value = separator + destVals.join(separator) + separator; 
//alert('destItems.value = ' + f.destItems.value); //demo only 
//f.submit();
return true; 
} 
</script>
<%
Dim connTemp, query, rs
set rs=server.CreateObject("ADODB.RecordSet")
if request.form("Submit")<>"" then
	call opendb()
	liststring=request.form("destItems")
	listarray=split(liststring,"|")
	pcv_provider=request.Form("provider")
	for i=1 to (ubound(listarray)-1)
		query="UPDATE shipService SET servicePriority="&i&" WHERE idshipservice="&listarray(i)
		set rs=connTemp.execute(query)
	next
	set rs=nothing
	call closedb()
	response.redirect "OrderShippingOptions.asp?provider="&pcv_provider
else
%>
<form name="data" action="OrderShippingOptions.asp" method="post" class="pcForms">
	<input name="provider" type="hidden" value="<%=request("provider")%>">
	<table class="pcCPcontent">
		<tr> 
			<td colspan="2">During the checkout process, your customers will see a list of available shipping options. Here you can set the order in which they are displayed. Make sure to click on the &quot;Save&quot; button when you are done.
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td width="60%" valign="top"> 
				<select name="dest" size="8" style="width:100%;height:400px;">
				<% tmpStr=""
				if request("provider")="x" then
					tmpStr=" AND serviceCode like 'C%'"
				end if
				if request("provider")="usps" then
					tmpStr=" AND serviceCode like '99%'"
				end if
				if request("provider")="ups" then
					tmpStr=" AND serviceDescription like 'UPS%' AND serviceCode not like 'C%'"
				end if
				if request("provider")="fedex" then
					tmpStr=" AND serviceDescription like 'FedEx%' AND serviceCode not like 'C%'"
				end if
				if request("provider")="cp" then
					tmpStr=" AND serviceDescription like '%anada Post%' AND serviceCode not like 'C%'"
				end if
				set rs=server.CreateObject("ADODB.RecordSet")
				call opendb()
				query="SELECT idshipservice, servicePriority, serviceDescription FROM shipService WHERE serviceActive=-1"& tmpStr &" ORDER BY servicePriority;"
				set rs=connTemp.execute(query)
				do until rs.eof %>
					<option value="<%=rs("idshipservice")%>"><%=rs("serviceDescription")%></option>
				<%
					rs.moveNext
					loop 
					set rs=nothing
					call closedb()
				%>
				</select>
			</td>
			<td width="40%" align="center" valign="middle"> 
				<input id="up" type="button" value="Move Up" onClick="moveSelected(dest,false)" />
				<br /><br />		
				<input id="down" type="button" value="Move Down" onClick="moveSelected(dest,true)" />
				<br /><br />
				<input type="hidden" name="destItems" />
				<input  name="submit" type="submit" onClick="return setHidden(this.form)" value="Save" class="submit2"/>
			</td>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>           
			<td align="center" colspan="2">
				<input type="button" value="Shipping Options Summary" onClick="location.href='viewShippingOptions.asp'">
				&nbsp;
				<input type="button" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->