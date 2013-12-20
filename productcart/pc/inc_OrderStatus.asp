<%
'// Determine order processing and payment status and load corresponding icons
	select case porderstatus
	case "2"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_2") 
	case "3"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_3") 
	case "4"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_4")
	case "5"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_5") 
	case "6"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_6") 
	case "7"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_7") 
	case "8"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_8") 
	case "9"
	  response.write dictLanguage.Item(Session("language")&"_sds_viewpast_9") 
	  
	'// Google Checkout
	case "10"
	  response.write "Delivered"
	case "11"
	  response.write "Will not deliver"
	case "12"
	  response.write "Archived"
	end select
%>