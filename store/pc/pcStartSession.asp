<!--#include file="../includes/productcartinc.asp"-->
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************
' Check if store is ON or OFF
'*******************************
If scStoreOff="1" then
	response.redirect "msg.asp?message=83"
End If

Function SEOcheckAff()
Dim tmpStr1,tmpStr2,k,tmp1
	tmp1=Cint(1)
	if session("strSEOAffiliate")<>"" then
		if IsNumeric(session("strSEOAffiliate")) then
			tmp1=session("strSEOAffiliate")
		end if
	end if
	SEOcheckAff=tmp1
End Function

'*******************************
' START ProductCart Session
'*******************************
HaveNewSession=0
if session("idcustomer")="" then
	Dim pcv_intFlagNoLocal
	pcv_intFlagNoLocal=Cint(0)
	session("idPCStore")= scID
	session("idCustomer")=Cint(0)
	session("customerCategory")=Cint(0)
	session("customerType")=Cint(0)
	session("ATBCustomer")= Cint(0)
	session("ATBPercentOff")= Cint(0)
	session("idAffiliate")=SEOcheckAff()     
	session("language")=Cstr("english")
	session("pcCartIndex")=Cint(0)	
	dim pcCartArrayORG(100,45)
	session("pcCartSession")=pcCartArrayORG
	HaveNewSession=1
end if
if session("idPCStore")<>scID then
	session.Abandon()
	session("idPCStore")= scID
	session("idCustomer")=Cint(0)
	session("customerCategory")=Cint(0)
	session("customerType")=Cint(0)
	session("ATBCustomer")= Cint(0)
	session("ATBPercentOff")= Cint(0)
	session("idAffiliate")=SEOcheckAff()     
	session("language")=Cstr("english")
	session("pcCartIndex")=Cint(0)
	redim pcCartArrayORG(100,45)
	session("pcCartSession")=pcCartArrayORG
	HaveNewSession=1
end if
pcCartArray=session("pcCartSession")
'*******************************
' END ProductCart Session
'*******************************
%>
<!--#include file="../includes/pcAffConstants.asp"-->
<%
'*******************************
' AFFILIATE - START
'*******************************
	dim pcInt_IdAffiliate, pcv_SavedAffiliateID
	dim pcInt_UseAffiliate, pcv_AffiliateCookiePath
	
	session("pcInt_AllowedAffOrders")=scAllowedAffOrders

	IF scAffProgramActive="1" THEN

		'// Check for previously stored affiliate ID
		IF session("idAffiliate")=1 THEN
			pcv_SavedAffiliateID=getUserInput(Request.Cookies("SavedAffiliateID"),0)
			If validNum(pcv_SavedAffiliateID) then
				pcInt_IdAffiliate=pcv_SavedAffiliateID
			Else
				pcInt_IdAffiliate=1
			End If
			session("idAffiliate")=trim(pcInt_IdAffiliate)
	
			'// If none, check querystring
			If session("idAffiliate")=1 then
				pcInt_IdAffiliate = request.querystring("idAffiliate")
				if validNum(pcInt_IdAffiliate) then
					session("idAffiliate")=pcInt_IdAffiliate
				else
					session("idAffiliate")=1
				end if
			End if
		END IF
		
		'// Set cookie with Affiliate ID, if feature is active
		If scSaveAffiliate="1" and session("idAffiliate")<>1 Then
			Response.Cookies("SavedAffiliateID")=session("idAffiliate")
			pcInt_SaveAffiliateDays=Cint(scSaveAffiliateDays)
			if NOT validNum(pcInt_SaveAffiliateDays) then
				pcInt_SaveAffiliateDays=365
			end if
			Response.Cookies("SavedAffiliateID").Expires=Date() + pcInt_SaveAffiliateDays
		end if
		
	END IF

'*******************************
' AFFILIATE - END
'*******************************


'*******************************
' set Reward Points - referral
'*******************************
Dim pcvIntRefBy
pcvIntRefBy=getUserInput(request("refby"),10)
if not validNum(pcvIntRefBy) then pcvIntRefBy=""
If Session("ContinueRef")="" then
	If RewardsActive=1 then
		If (pcvIntRefBy <> "") And (RewardsReferral=1) Then
			Session("ContinueRef")=CLng(pcvIntRefBy)
		End If
	End if
End if

IF HaveNewSession=1 THEN%>
<!--#include file="inc_RestoreShoppingCart.asp"-->
<%END IF%>