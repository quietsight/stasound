<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="header.asp"-->

<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td> 
			<p>&nbsp;</p>
			<%
			msg=request.querystring("message")
				'Check that msg is a number
				if not validNum(msg) then
					msg = 0
					response.write dictLanguage.Item(Session("language")&"_techErr_1")
				end if
			on error resume next
			
			Dim pcStrClass
				pcStrClass="pcErrorMessage"			

			select case msg
			case 1
				pcStrClass="pcInfoMessage"
			end select				
			%>
			<div class="<%=pcStrClass%>">
			<%
			select case msg
			case 1
				response.write dictLanguage.Item(Session("language")&"_showcart_1")&"<br><br><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"    
			case 2  
				response.write dictLanguage.Item(Session("language")&"_forgotpassworderror") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 3
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 4
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 5
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 6
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 7
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 8
				dstr=replace(scStoreMsg,"''","~~~")
				dstr=replace(dstr,"'","""")
				dstr=replace(dstr,"~~~","'")
				response.write dstr
			case 9
				response.write dictLanguage.Item(Session("language")&"_checkout_1")&"<br><br><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"    
			case 10
				response.write dictLanguage.Item(Session("language")&"_msg_202")
			case 11
				response.write dictLanguage.Item(Session("language")&"_CustviewPastD_16")
			case 12
				response.write dictLanguage.Item(Session("language")&"_msg_12")
			case 13
				response.write dictLanguage.Item(Session("language")&"_msg_13")
			case 14
				response.write dictLanguage.Item(Session("language")&"_chooseShpmnt_1") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>" 
			case 15
				response.write dictLanguage.Item(Session("language")&"_chooseShpmnt_2") &"<br><br><a href='checkout.asp?cmode=2'><img src="""& rslayout("back") & """></a>"
			case 16
				response.write dictLanguage.Item(Session("language")&"_cRec_1")
			case 17
				response.write dictLanguage.Item(Session("language")&"_cRec_2") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 18
				response.write "" 
			case 19
				response.write dictLanguage.Item(Session("language")&"_cRemv") 
			case 20
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"       
			case 21
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>" 
			case 22
				response.write dictLanguage.Item(Session("language")&"_Custmodb_2")
			case 23
				response.write dictLanguage.Item(Session("language")&"_CustPastAdd_2")  
			case 24
				response.write dictLanguage.Item(Session("language")&"_CustPastAdd_3")
			case 25
				response.write dictLanguage.Item(Session("language")&"_CustPastAdd_4")
			case 26
				response.write dictLanguage.Item(Session("language")&"_CustPastAdd_1")
			case 27
				response.write dictLanguage.Item(Session("language")&"_additem_6")
			case 28
				response.write dictLanguage.Item(Session("language")&"_CustRegb_3")
			case 29
				response.write dictLanguage.Item(Session("language")&"_CustRegb_1")
			case 30
				response.write dictLanguage.Item(Session("language")&"_CustRegb_2")
			case 31
				dstr=replace(scStoreMsg,"''","~~~")
				dstr=replace(dstr,"'","""")
				dstr=replace(dstr,"~~~","'")
				response.write dstr
			case 32
				response.write dictLanguage.Item(Session("language")&"_Custvb_1") &"<br><br><a href=Checkout.asp?cmode=1&redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & "><img src="""& rslayout("back") & """></a>"    
			case 33
				response.write dictLanguage.Item(Session("language")&"_Custvb_2") &"<br><br><a href=Checkout.asp?cmode=1&redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & "><img src="""& rslayout("back") & """></a>" 
			case 34
				response.write dictLanguage.Item(Session("language")&"_CustviewPast_1")
			case 35
				response.write dictLanguage.Item(Session("language")&"_CustviewPastD_1")
			case 36
				response.write dictLanguage.Item(Session("language")&"_Custwl_1")
			case 37
				response.write dictLanguage.Item(Session("language")&"_CustwlRmv_1")
			case 38
				response.write dictLanguage.Item(Session("language")&"_msg_38")
			case 39
				response.write dictLanguage.Item(Session("language")&"_instPrd_B") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>"   
			case 40
				response.write dictLanguage.Item(Session("language")&"_instPrd_C")& "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>"     
			case 41
				response.write dictLanguage.Item(Session("language")&"_instPrd_D")    
			case 42
				response.write dictLanguage.Item(Session("language")&"_instPrd_E")
			case 43
				response.write dictLanguage.Item(Session("language")&"_instPrd_E")  
			case 44
				response.write dictLanguage.Item(Session("language")&"_instPrd_E")
			case 45
				response.write dictLanguage.Item(Session("language")&"_instPrd_E")
			case 46
				response.write dictLanguage.Item(Session("language")&"_instPrd_E")
			case 47
				response.write dictLanguage.Item(Session("language")&"_instPrd_A")
			case 48
				response.write ""     
			case 49
				response.write dictLanguage.Item(Session("language")&"_instPrd_C")
			case 50
				response.write dictLanguage.Item(Session("language")&"_instPrd_B") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>"  
			case 51
				response.write dictLanguage.Item(Session("language")&"_instPrd_C")& "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>" 
			case 52
				response.write dictLanguage.Item(Session("language")&"_instPrd_D")
			case 53
				response.write "" 
			case 54
			response.write dictLanguage.Item(Session("language")&"_login_2") &"<br><br><a href=checkout.asp><img src="""& rslayout("back") & """></a>" 
			case 55
				response.write dictLanguage.Item(Session("language")&"_login_2")&"<br><br><a href=checkout.asp><img src="""& rslayout("back") & """></a>"  
			case 56
				response.write dictLanguage.Item(Session("language")&"_login_3")&"<br><br><a href=checkout.asp><img src="""& rslayout("back") & """></a>"    
			case 57
				response.write ""  
			case 58
			response.write dictLanguage.Item(Session("language")&"_instPrd_C")
			case 59
				dstr=replace(scStoreMsg,"''","~~~")
				dstr=replace(dstr,"'","""")
				dstr=replace(dstr,"~~~","'")
				response.write dstr
			case 60
				response.write dictLanguage.Item(Session("language")&"_mainIndex_1")
			case 61
				response.write dictLanguage.Item(Session("language")&"_NewCust_1")&"<br><br><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"    
			case 62
				response.write dictLanguage.Item(Session("language")&"_orderverify_1")
			case 63
				response.write dictLanguage.Item(Session("language")&"_orderverify_2")
			case 64
				response.write dictLanguage.Item(Session("language")&"_paymntb_c_1")
			case 65
				response.write dictLanguage.Item(Session("language")&"_paymntb_c_2")
			case 66
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_1")
			case 67
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_6") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>" 
			case 68
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_5") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>" 
			case 69
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_8") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>" 
			case 70
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_7") & "<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &"""></a>"    
			case 71
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_4")&"<br><br><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"
			case 72
				response.write dictLanguage.Item(Session("language")&"_paymntb_o_4")
			case 73
				response.write dictLanguage.Item(Session("language")&"_msg_73")
			case 74
				response.write dictLanguage.Item(Session("language")&"_viewPrd_1")
			case 75
				response.write dictLanguage.Item(Session("language")&"_msg_75")&"<br><br><a href=javascript:history.go(-1)><img src="& rslayout("back")& "></a>"
			case 76
				response.write dictLanguage.Item(Session("language")&"_msg_76")
			case 77
				response.write dictLanguage.Item(Session("language")&"_orderverify_5")&"<br><br><a href=javascript:history.go(-1)><img src="& rslayout("back")& "></a>"
			case 78
				response.write dictLanguage.Item(Session("language")&"_advSrcb_1") &"<br><br><a href=search.asp><img src="""& rslayout("back") & """></a>"
			case 79
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp><img src="& rslayout("back")& "></a>"    
			case 80
				response.write ""
			case 81
				response.write ""
			case 82
				response.write dictLanguage.Item(Session("language")&"_updOrdStats_2")&"<br><br><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"    
			case 83
				dstr=replace(scStoreMsg,"''","~~~")
				'dstr=replace(dstr,"'","""")
				dstr=replace(dstr,"~~~","'")
				dstr=replace(dstr, "&lt;BR&gt;", "<br>")
				response.write dstr
			case 84
				dstr=replace(scStoreMsg,"''","~~~")
				'dstr=replace(dstr,"'","""")
				dstr=replace(dstr,"~~~","'")
				dstr=replace(dstr, "&lt;BR&gt;", "<br>")
				response.write dstr
			case 85
				response.write dictLanguage.Item(Session("language")&"_viewCat_P_1")
			case 86
				response.write dictLanguage.Item(Session("language")&"_viewCat_P_6")
			case 87
				response.write dictLanguage.Item(Session("language")&"_viewCat_P_1")
			case 88
				response.write dictLanguage.Item(Session("language")&"_viewPrd_2")
			case 89
				response.write dictLanguage.Item(Session("language")&"_viewSpc_1")
			case 90
				response.write dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=""viewBrands.asp""><img src="""& rslayout("back") & """></a>"
			case 91
				response.write dictLanguage.Item(Session("language")&"_AffLogin_10") &"<br><br><a href=AffiliateLogin.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & "><img src="""& rslayout("back") & """></a>"    
			case 92
				response.write dictLanguage.Item(Session("language")&"_AffLogin_11") &"<br><br><a href=AffiliateLogin.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & "><img src="""& rslayout("back") & """></a>" 
			case 93
				response.write dictLanguage.Item(Session("language")&"_viewNewArrivals_1")
			case 94
				response.write dictLanguage.Item(Session("language")&"_viewBestSellers_1")
			case 95
				response.write dictLanguage.Item(Session("language")&"_viewPrd_62")
				
			'GGG Add-on start
			case 96
				response.write dictLanguage.Item(Session("language")&"_msg_4") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 97
				response.write dictLanguage.Item(Session("language")&"_msg_5") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 98
				response.write dictLanguage.Item(Session("language")&"_msg_6") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 99
				response.write dictLanguage.Item(Session("language")&"_msg_7") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 100
				response.write dictLanguage.Item(Session("language")&"_msg_8") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 101
				response.write dictLanguage.Item(Session("language")&"_msg_9") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			case 102
				response.write dictLanguage.Item(Session("language")&"_msg_10") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """ border=0></a>"
			'GGG Add-on end
			
			case 130
				response.write bto_dictLanguage.Item(Session("language")&"_configurePrd_19") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 131
				response.write dictLanguage.Item(Session("language")&"_checkout_13") &"<br><br><a href=viewCart.asp><img src="""& rslayout("back") & """></a>"
			case 132
				response.write dictLanguage.Item(Session("language")&"_alert_12") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 133
				response.write dictLanguage.Item(Session("language")&"_alert_13") &"<br><br><a href=""repeatorder.asp?idOrder=" & request("idorder") & "&OrderRepeat=haveto""><img src="""& rslayout("submit") & """ border=""0""></a>&nbsp;<a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 134
				response.write dictLanguage.Item(Session("language")&"_alert_14") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 135
				response.write dictLanguage.Item(Session("language")&"_alert_15") &"<br><br><a href=""addsavedprdstocart.asp?OrderRepeat=haveto""><img src="""& rslayout("submit") & """ border=""0""></a>&nbsp;<a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 200
				response.write dictLanguage.Item(Session("language")&"_techErr_2")
			case 201
				response.write dictLanguage.Item(Session("language")&"_msg_201")
			case 202
				response.write dictLanguage.Item(Session("language")&"_sdsLogin_8") &"<br><br><a href=sds_Login.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & "><img src="""& rslayout("back") & """></a>"
			case 203  
				response.write dictLanguage.Item(Session("language")&"_sds_forgotpassworderror") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
				
			case 204
				' Error Description: Quantity being ordered is greater than quantity in stock
				' Set local variables and clear session variables
				pDescription = session("pcErrStrPrdDesc")
				session("pcErrStrPrdDesc") = ""
				pStock = session("pcErrIntStock")
				session("pcErrIntStock") = Cint(0)
				response.write dictLanguage.Item(Session("language")&"_instPrd_2")&pDescription&dictLanguage.Item(Session("language")&"_instPrd_3")&pStock&dictLanguage.Item(Session("language")&"_instPrd_4")&"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>"
				
			case 205
				' Error Description: Wholesale minimum not met, so customer cannot checkout
				response.write dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)&dictLanguage.Item(Session("language")&"_techErr_3") & "<BR><BR><a href='viewCart.asp'>"&dictLanguage.Item(Session("language")&"_mainIndex_5")&"</a>"
				
			case 206
				' Error Description: Retail minimum not met, so customer cannot checkout
				response.write dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scMinPurchase) & "<BR><BR><a href='viewCart.asp'>"&dictLanguage.Item(Session("language")&"_mainIndex_5")&"</a>"
				
			case 207
				' Error Description: The product ID could not be retrieved
				response.write dictLanguage.Item(Session("language")&"_PrdError_1")&"<br /><br/><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"
			case 208  
				response.write dictLanguage.Item(Session("language")&"_PayPal_2") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 209  
				response.write dictLanguage.Item(Session("language")&"_PayPal_3") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
				
			case 210
				' Error Description: generic error due to invalid product or other ID
				response.write dictLanguage.Item(Session("language")&"_msg_210")
				
			case 211
				' Error Description: Your session is invalid
				response.write dictLanguage.Item(Session("language")&"_validateform_9")&"<br /><br/><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"
			case 212
				' Error Description: This browser does not accept cookies.
				response.write dictLanguage.Item(Session("language")&"_PrdError_2")&"<br /><br/><a href=default.asp><img src="""& rslayout("continueshop") &"""></a>"
				
			'// ProductCart v4
			case 300
				' Content page cannot be accessed
				response.write dictLanguage.Item(Session("language")&"_viewPages_1") &"<br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") & """></a>"
			case 301
				' Content page cannot be accessed
				response.write dictLanguage.Item(Session("language")&"_ShowRecentRev_2")
			case 302
				' No brands available
				response.write dictLanguage.Item(Session("language")&"_msg_302")
			case 303
				' No subbrands available
				response.write dictLanguage.Item(Session("language")&"_msg_303")
			case 304
				' Did not setup header and footer properly
				dstr=replace(scStoreMsg,"''","~~~")
				dstr=replace(dstr,"~~~","'")
				dstr=replace(dstr, "&lt;BR&gt;", "<br>")
				response.write dstr
			case 305
				' Subscription product in cart
				response.write scSBLang5
			case 306
				' Subscription product not allowed
				response.write scSBLang6
			case 307
				' No payment methods available
				response.write dictLanguage.Item(Session("language")&"_EIG_17")
			case 308
				' Duplicate Order Detected
				response.write dictLanguage.Item(Session("language")&"_OPC_Alert_01")
			case 309
				' BTO items are low of stock
				response.write dictLanguage.Item(Session("language")&"_instConfQty_1")
			case 310
				'Suspended Account Checkout
				response.write dictLanguage.Item(Session("language")&"_opc_checkorv_3")
			end select 
			 %>
			 </div>
			 <p>&nbsp;</p>
			</td>
		</tr>
	</table>
</div>
<!--#include file="footer.asp"-->