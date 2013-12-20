<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<%Dim connTemp,query,rs

' Check if the store is on. If store is turned off display store message
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If 
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
<form name="Form1" action="ggg_srcGRb.asp?action=search" method="POST" class="pcForms">
<table class="pcMainTable">
	<tr>
		<td>
			<table class="pcShowContent">
				<tr>
					<td width="100%" colspan="2"><h1><%response.write dictLanguage.Item(Session("language")&"_SrcGR_1")%></h1></td>
				</tr>
				<tr>
					<td width="100%" colspan="2"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_8")%></td>
				</tr>
				<tr>
					<td width="25%" nowrap><%response.write dictLanguage.Item(Session("language")&"_SrcGR_2")%></td>
					<td width="75%">
						<input type=text name="cname" value="" size="30">
					</td>
				</tr>
				<tr>
					<td width="25%" nowrap><%response.write dictLanguage.Item(Session("language")&"_SrcGR_3")%></td>
					<td width="75%">
						<input type=text name="clastname" value="" size="30">
					</td>
				</tr>
				<tr>
					<td width="25%" nowrap><%response.write dictLanguage.Item(Session("language")&"_SrcGR_9")%></td>
					<td width="75%">
						<input type=text name="cregname" value="" size="30">
					</td>
				</tr>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr>
					<td width="25%" nowrap>
						<%response.write dictLanguage.Item(Session("language")&"_SrcGR_4")%>
					</td>
					<td width="75%">
						<%response.write dictLanguage.Item(Session("language")&"_SrcGR_5")%>
						<select name="emonth">
							<option value="" selected><%response.write dictLanguage.Item(Session("language")&"_SrcGR_5")%></option>
							<option value="1"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_11")%></option>
							<option value="2"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_12")%></option>
							<option value="3"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_13")%></option>
							<option value="4"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_14")%></option>
							<option value="5"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_15")%></option>
							<option value="6"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_16")%></option>
							<option value="7"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_17")%></option>
							<option value="8"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_18")%></option>
							<option value="9"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_19")%></option>
							<option value="10"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_20")%></option>
							<option value="11"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_21")%></option>
							<option value="12"><%response.write dictLanguage.Item(Session("language")&"_SrcGR_22")%></option>
						</select>
						&nbsp;
						<%response.write dictLanguage.Item(Session("language")&"_SrcGR_6")%>
						<select name="eyear">
							<option value="" selected><%response.write dictLanguage.Item(Session("language")&"_SrcGR_6")%></option>
							<option value="<%=year(date())%>"><%=year(date())%></option>
							<option value="<%=year(date())+1%>"><%=year(date())+1%></option>
							<option value="<%=year(date())+2%>"><%=year(date())+2%></option>
							<option value="<%=year(date())+3%>"><%=year(date())+3%></option>
						</select>
					</td>
				</tr>
				<tr>
					<td colspan="2" class="pcSpacer"></td>
				</tr>
				<tr>
					<td colspan="2"><input type="image" id="submit" src="<%=rslayout("submit")%>" border="0" name="Submit"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</div>
<!--#include file="footer.asp"-->