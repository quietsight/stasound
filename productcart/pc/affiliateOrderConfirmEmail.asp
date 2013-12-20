<%
	AffiliateOrderEmail=""
	AffiliateOrderEmail=AffiliateOrderEmail & dictLanguage.Item(Session("language")&"_storeEmail_10") & AffiliateName &","
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_11") & scCompanyName & dictLanguage.Item(Session("language")&"_storeEmail_12")
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_13") & scCurSign&money(AffiliatePay)
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf
	AffiliateOrderEmail=AffiliateOrderEmail & VBcrLf & dictLanguage.Item(Session("language")&"_storeEmail_14") & scCompanyName
%>
