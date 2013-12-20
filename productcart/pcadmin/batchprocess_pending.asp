<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languages_ship.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" --> 
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="inc_GenDownloadInfo.asp"-->
<!--#include file="adminHeader.asp" -->
<% on error resume next
dim query, conntemp, rs

'// How many checkboxes?
checkboxCnt=request.Form("checkboxCnt")

call opendb()

'////////////////////////////////////////////////////
'// START: Process Selected Orders
'////////////////////////////////////////////////////
dim r
For r=1 to checkboxCnt
	if request.Form("checkOrd"&r)="YES" then
		pOrderStatus=request.Form("orderstatus"&r)
		pCheckEmail=request.Form("checkEmail"&r)
		pIdOrder=Request.Form("idOrder"&r)  & ""
		qry_ID=pIdOrder
		pcv_CustomerReceived=0
		pcv_AdmComments=""
		pcv_SubmitType=3
		
		'// START:  Process Order and Send Notification E-mails
		%>  <!--#include file="inc_ProcessOrder.asp"-->  <%
		'// END:  Process Order and Send Notification E-mails

		successCnt=successCnt+1
		successData=successData&"Order Number "& (int(pIdOrder)+scpre) &" was processed successfully<BR>" 
	end if
Next
'////////////////////////////////////////////////////
'// END: Process Selected Orders
'////////////////////////////////////////////////////

call closedb()
%>
<table class="pcCPcontent">
<tr>
	<td><div class="pcCPmessageSuccess"><%=successCnt%> records were successfully processed.</div>
		<% if successData<>"" then %>
			<br><%=successData%><br>
		<% end if %>
	</td>
</tr>
<tr>
	<td>
    	<p>&nbsp;</p>
	    <p><a href="resultsAdvancedAll.asp?B1=View%2BAll&dd=1">Manage Orders</a></p>
	</td>
</tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<% Public Function FixedField(ByVal Width, ByVal Justify, ByVal Text)

	Select Case True
		Case Width < Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Left(Text, Width)
				Case Justify="R"
					FixedField=Right(Text, Width)
				Case Else
			End Select
									
		Case Width=Len(Text)
			FixedField=Text

		Case Width > Len(Text)
			Select Case True
				Case Justify="L"
					FixedField=Text & String(Width - Len(Text), " ")
				Case Justify="R"
					FixedField=String(Width - Len(Text), " ") & Text
				Case Else
			End Select

	End Select

End Function %><!--#include file="adminFooter.asp" -->