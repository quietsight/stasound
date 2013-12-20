<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/openDb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/productcartinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="opc_contentType.asp" -->
<% On Error Resume Next
HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

dim pcCartArray, ppcCartIndex, f, cont

Call SetContentType()

IF HaveSecurity=0 THEN

	'*****************************************************************************************************
	'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
	%><!--#include file="pcVerifySession.asp"--><%
	pcs_VerifySession
	'*****************************************************************************************************
	'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
	
	pIdCustomer=session("idCustomer")
	ppcCartIndex=Session("pcCartIndex")
	pcCartArray=Session("pcCartSession")
	
	f=getUserInput(request("index"),0)
	
	if pcCartArray(f,34)="" then
		pcCartArray(f,34)="0"
	end if

	UpdateSuccess="0"
	tmpPrdList=""
	if request("action")="add" then
		
		if pcCartArray(f,10)=0 then
			GW=getUserInput(request("GW" & f),0)
			pcCartArray(f,34)=GW
			if GW<>"" AND GW<>"0" then
				if tmpPrdList<>"" then
					tmpPrdList=tmpPrdList & ","
				end if
				tmpPrdList=tmpPrdList & pcCartArray(f,0)
			end if
			GWText=URLDecode(getUserInput(request("GWText" & f),240))
			if GWText<>"" then
				GWText=replace(GWText,"''","'")
			end if
			pcCartArray(f,35)=GWText
		end if

		session("pcCartSession")=pcCartArray
		UpdateSuccess="1"
	end if

END IF

dim conntemp,query,rs
call openDb()
%>
<html>
<head>
<TITLE><%=dictLanguage.Item(Session("language")&"_opc_gwa_title")%></TITLE>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN  

Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")
%>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript">
<!--
	
	function Form1_Validator(theForm)
	{
		return (true);
	}
	function testchars(tmpfield,idx)
	{
		var tmp1=tmpfield.value;
		if (tmp1.length>240)
		{
			alert("<%response.write FixLang(dictLanguage.Item(Session("language")&"_GiftWrap_9"))%>");
			tmp1=tmp1.substr(0,240);
			tmpfield.value=tmp1;
			document.getElementById("countchar" + idx).innerHTML=240-tmp1.length;
			tmpfield.focus();
		}
		document.getElementById("countchar" + idx).innerHTML=240-tmp1.length;
	}
//-->
</script>
</head>
<body style="margin: 0;">
<div id="pcMain">
<form method="post" name="Form1" action="opc_GiftWrap.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
<input name="index" type="hidden" value="<%=f%>">
<table class="pcMainTable">
	<%IF HaveSecurity=1 THEN%>
        <tr>
            <td>
                <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_opc_gwa_1")%></div>
            </td>
        </tr>
    <%ELSE%>
    
        <%IF UpdateSuccess="1" THEN%>
            <tr>
                <td>
                    <div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_gwa_2")%></div>
                    <script>
                        var tmpPanelTest = parent.document.getElementById('GWMarker<%=pcCartArray(f,0)%>');
						if (tmpPanelTest == null) {
						} else {
							if ('<%=pcCartArray(f,34)%>' != '') {
								parent.document.getElementById('GWMarker<%=pcCartArray(f,0)%>').innerHTML = '<a href="javascript:;" onclick="GWAdd(\'<%=pcCartArray(f,0)%>\', \'<%=f%>\');"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_giftWrap_1"))%></a>';
							} else {
								parent.document.getElementById('GWMarker<%=pcCartArray(f,0)%>').innerHTML = '<a href="javascript:;" onclick="GWAdd(\'<%=pcCartArray(f,0)%>\', \'<%=f%>\');"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_giftWrap_2"))%></a>';				
							}
							parent.GetOrderInfo("","#GWframeloader",0,'');
						}
                        setTimeout(function(){parent.closeGWDialog()},1000);
                    </script>
                </td>
            </tr>        
		<%ELSE%>            
            <%
			if pcCartArray(f,10)=0 then
				pIDProduct=pcCartArray(f,0)
				pName=pcCartArray(f,1)
				%>
                <tr>
                	<td colspan="2"><h2><%response.write dictLanguage.Item(Session("language")&"_opc_giftWrap_3") & pName %></h2></td>
                </tr>
				<tr>
					<td width="60%" valign="top">
					<table class="pcShowContent">
					<td><input type="radio" name="GW<%=f%>" value="" class="clearBorder"></td>
					<td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_GiftWrap_3")%></td>
					<%query="SELECT pcGW_IDOpt,pcGW_OptName,pcGW_OptImg,pcGW_OptPrice from pcGWOptions WHERE pcGW_removed=0 AND pcGW_OptActive=1 ORDER BY pcGW_OptOrder ASC,pcGW_OptName ASC;"
					set rstemp=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rstemp=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					Count=0
					do while not rstemp.eof
					IDOpt=rstemp("pcGW_IDOpt")
					OptName=rstemp("pcGW_OptName")
					OptImg=rstemp("pcGW_OptImg")
					OptPrice=rstemp("pcGW_OptPrice")%>
					<tr>
						<td valign="top"><input type="radio" name="GW<%=f%>" value="<%=IDOpt%>" <%if (Count=0) or (cdbl(IDOpt)=cdbl(pcCartArray(f,34))) then%>checked<%Count=1%><%end if%> class="clearBorder"></td>
						<td valign="top" width="100%"><%=OptName%><br>
							<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_8")%><%if cdbl(OptPrice)=0 then%><b><%response.write dictLanguage.Item(Session("language")&"_GiftWrap_6")%></b><%else%><%=scCurSign & money(OptPrice)%><%end if%>
						</td>
						<td><%if OptImg<>"" then%><img src="catalog/<%=OptImg%>" border="0" align="top"><%end if%></td>
					</tr>
					<%rstemp.MoveNext
					loop%>
					</table>
					</td>
					<td width="40%" valign="top" style="border-left: 1px solid #EEE; padding-left: 10px;">
						<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_4")%><br>
						<textarea name="GWText<%=f%>" rows="6" cols="20" onKeyUp="javascript:testchars(this,'<%=f%>');" maxlength="240"><%=pcCartArray(f,35)%></textarea><br>
						<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5a")%><span id="countchar<%=f%>" name="countchar<%=f%>"style="font-weight: bold">240</span> <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%><br>
					</td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
				<%
			end if
        	%>	
            <tr>
                <td colspan="2" class="pcSpacer"></td>
            </tr>	
            <tr>
                <td colspan="2"> 
                    <input type="image" id="GWASubmit" name="GWASubmit" value="<%response.write dictLanguage.Item(Session("language")&"_GiftWrap_7")%>" src="<%=RSlayout("pcLO_Update")%>" border="0" class="clearBorder">
                </td>
            </tr>
        <%END IF%>
    <%END IF%>
</table>
</form>
</div>
</body>
</html>
<!--#include file="inc_SaveShoppingCart.asp"-->
<% 
call closedb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing 
%>