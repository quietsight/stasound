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
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/validation.asp"-->
<% 'SB S %>
<!--#include file="inc_sb.asp"--> 
<% 'SB E %> 
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

//-->
</script>
</head>
<!--#include file="../includes/pcServerSideValidation.asp" -->
  
<%
'on error resume next

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

dim pcCartArray, ppcCartIndex, f, cont

'*****************************************************************************************************
'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
%><!--#include file="pcVerifySession.asp"--><%
pcs_VerifySession
'*****************************************************************************************************
'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
'*****************************************************************************************************
ppcCartIndex=Session("pcCartIndex")

dim conntemp

call openDb()


if request("action")="verify" then

    aCnt = Request.form("aCnt")
	fCnt = 0
	
   for a = 1 to aCnt   
    recFound = 0
	for f=1 to pcCartIndex
	    pSubscriptionID = cstr(pcCartArray(f,38))
		if  pSubscriptionID = request.form("Agree_"&pSubscriptionID&"_"&a) then			
			 recFound = 1			
			 fCnt = fCnt + 1 
			 exit for		
		 end if 
	Next
	if recFound = 1 then
		  session("agree_"&pSubscriptionID&"_"&a) =  pSubscriptionID
	else
		  session("agree_"& pSubscriptionID&"_"&a) = ""
	end if
  Next
	session("pcCartSession")=pcCartArray
	session("pcAgreeCnt") = aCnt
	if cint(aCnt) = cint(fCnt) then
		%>
		<script>
            var tmpPanelTest = parent.document.getElementById('AgreeRegTerms');
            if (tmpPanelTest == null) {
            } else {
            	tmpPanelTest.checked=true;
            }
            setTimeout(function(){parent.closeRegDialog()},50);
        </script>
        <%
		session("pcAgreeAll") = True
		msg = " Thank You. "
	else
		session("pcAgreeAll") = False
		msg = " You must check all the  ""I Agree"" boxes to checkout."
	end if
	
end if

%>


<body style="margin: 0;">
<div id="pcMain">
<form method="post" name="Form1" action="sb_RegAgree.asp?action=verify" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcMainTable">

	<tr> 
		<td colspan="2"><h1><%response.write scSBLang1%></h1></td>
	</tr>
 <% if msg <> "" then %>
   <tr> 
		<td colspan="2"><div class="pcErrorMessage"><%=msg%></div></td>
	</tr>
 <%end if%>
	<%
	 	i=0
		call opendb()
	for f=1 to ppcCartIndex

	  pName=pcCartArray(f,1)
	  pSubscriptionID = pcCartArray(f,38)
		
	 if pSubscriptionID  > 0 then
		query="SELECT SB_Agree, SB_AgreeText FROM SB_Packages WHERE SB_PackageID=" & pSubscriptionID &" AND SB_Agree = 1;"
		set rss=server.CreateObject("ADODB.RecordSet")
		set rss=connTemp.execute(query)	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rss=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID 
		end if
		 pSubAgreeText = ""
		if not rss.eof then		
		    i = i + 1
			pSubAgreeText = rss("SB_AgreeText")
			if len(pSubAgreeText)=0 then
				pSubAgreeText = scSBAgreeText
			end if
			%>
			<tr>
				<td width="80%" valign="top">
				<p> 
					<b><%=pName%></b><br>
					<%response.write scSBLang2%>
					<br>
					<table width="100%" border="0" cellpadding="2" cellspacing="0">
                    	<tr>
							<td colspan="3">
                            	<%
								pSubAgreeText=replace(pSubAgreeText,"<br>",vbCrLf)
								pSubAgreeText=replace(pSubAgreeText,"<BR>",vbCrLf)								
								%>
                            	<textarea name="agreeText" readonly cols="60" rows="6" ><%=pSubAgreeText%></textarea>
                            </td>
						</tr>
						<tr>
							<td colspan="3" >
								<%response.write scSBLang3%> <input type="checkbox" name="Agree_<%=pSubscriptionID%>_<%=i%>" value="<%=pSubscriptionID%>" class="clearBorder" <% if session("Agree_"& pSubscriptionID&"_"& i) = cstr(pSubscriptionID) then %> checked <%end if%> >
								<%pcs_RequiredImageTag "Agree_"&pSubscriptionID&"_"&i, true %>
                            </td>
                        </tr>			
					</table>
				</p>
				</td>			
			</tr>
			<tr>
				<td colspan="2"><hr></td>
			</tr>
		<% else %>
		
        	<% 
			if scSBRegAgree then
            	i = i + 1
				pSubAgreeText = scSBAgreeText
				%>
				<tr>
					<td width="80%" valign="top">
					<p> 
						<b><%=pName%></b><br>
						<%response.write scSBLang2%>
						<br>
						<table width="100%" border="0" cellpadding="2" cellspacing="0">
                    		<tr>
								<td colspan="3">
                            		<%
									pSubAgreeText=replace(pSubAgreeText,"<br>",vbCrLf)	
									pSubAgreeText=replace(pSubAgreeText,"<BR>",vbCrLf)							
									%>
                                	<textarea name="agreeText" readonly cols="60" rows="6" ><%=pSubAgreeText%></textarea>
                                </td>
							</tr>
							<tr>
								<td colspan="3" >
									<%response.write scSBLang3%> <input type="checkbox" name="Agree_<%=pSubscriptionID%>_<%=i%>" value="<%=pSubscriptionID%>" class="clearBorder" <% if session("Agree_"& pSubscriptionID&"_"& i) = cstr(pSubscriptionID) then %> checked <%end if%> >
									<%pcs_RequiredImageTag "Agree_"&pSubscriptionID&"_"&i, true %>
                            	</td>
                        	</tr>			
						</table>
					</p>
					</td>			
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
                <%				
            end if 
			%>
		<%   
		end if 
	end if



	  
Next

call closedb()
%>	
	<tr>
		<td colspan="2" class="pcSpacer"></td>
	</tr>	
	<tr>
		<td colspan="2"> 
			<p>
			<input type="image" id="submit" name="submit" value="<%response.write scSBLang4%>" src="<%=RSlayout("submit")%>" border="0">
			<input type="hidden" name="aCnt" value="<%=i%>">
			
			
			</p>
		</td>
	</tr>
</table>
</form>
</div>
<%call closedb()%>
</body>
</html>