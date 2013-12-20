<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Order Payment Options" %>
<% Section="paymntOpt" %>
<%PmAdmin=5%><!--#include file="adminv.asp"-->
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
<% Dim connTemp, query, rs
set rs=server.CreateObject("ADODB.RecordSet")
if request.form("Submit")<>"" then
	call opendb()
	liststring=request.form("destItems")
	listarray=split(liststring,"|")
	for i=1 to (ubound(listarray)-1)
		query="UPDATE payTypes SET paymentPriority="&i&" WHERE idPayment="&listarray(i)
		set rs=connTemp.execute(query)
	next
	set rs=nothing
	call closedb()
	response.redirect "OrderPaymentOptions.asp"
else
%>
<form name="data" action="OrderPaymentOptions.asp" method="post" class="pcForms">
	<table class="pcCPcontent">
		<tr> 
			<td colspan="2">During the checkout process, your customers will see a list of available payment options. Here you can set the order in which they are displayed in the dropdown list. Make sure to click on the &quot;Save&quot; button when you are done. Google Checkout and PayPal Express Checkout are <u>not shown</u> in the list because they appear separately from the other payment options. </td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td valign="top" width="50%" align="center">
				<select name="dest" size="8" style="width:250px;height:250px;">
				<% 
				set rs=server.CreateObject("ADODB.RecordSet")
				call opendb()
				query="SELECT idPayment, paymentDesc FROM payTypes WHERE paymentDesc NOT LIKE '%google checkout%' AND paymentDesc NOT LIKE '%express checkout%' AND active<>0 ORDER BY paymentPriority;"
				set rs=connTemp.execute(query)
				if rs.eof then
					set rs=nothing
					call closedb()
					response.Redirect("pcPaymentSelection.asp")
				end if
				do until rs.eof
					pidPayment=rs("idPayment")
					ppaymentDesc=trim(rs("paymentDesc"))
					%>
					<option value="<%=rs("idPayment")%>"><%=rs("paymentDesc")%></option>
					<%
                    rs.moveNext
				loop 
				call closedb()
				%>
				</select>
			</td>
			<td width="50%" align="left" valign="top"> 
            	<div style="margin-bottom: 10px;"><input id="up" type="button" value="Move Up" onClick="moveSelected(dest,false)" /></div>
				<div style="margin-bottom: 10px;"><input id="down" type="button" value="Move Down" onClick="moveSelected(dest,true)" /></div>
				<div style="margin-bottom: 10px;"><input type="hidden" name="destItems" />
				<input name="submit" type="submit" onClick="return setHidden(this.form)" value="Save" class="submit2"/></div>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="center" colspan="2">
			<input type="button" name="back" value="Back" onClick="javascript:history.back()">&nbsp;
			<input type="button" value="Payment Options Summary" onClick="location.href='paymentOptions.asp'">
			</td>
		</tr>
	</table>
</form>
<% end if %>
<!--#include file="AdminFooter.asp"-->