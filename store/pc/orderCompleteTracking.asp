<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pOrderNumber = (int(pOID)+scpre)

'// HOW TO USE THIS PAGE
'// 1.) This file can help you add custom tracking code
'// 2.) The file is included at the bottom of orderComplete.asp, home.asp, and viewPrd.asp.
'// 3.) There are two sections in this page "TRACKING" and "LANDING": see below

	IF pcv_intOrderComplete=1 THEN	
	
		'////////////////////////////////////////////////////////////////////////////////////////
		'// TRACKING: enter below any code to track the successful completion of an order
		'////////////////////////////////////////////////////////////////////////////////////////

		' Advanced users: you use any of the following variables in your tracking code:
		' pOID:  This variable holds the order number from the ProductCart database
		' pOrderNumber: The order number shown to customer
		' ptotal: The order total amount
		' pOrderStatus: An integer representing the status of the order
		
		' ENTER your tracking code BELOW the ASP tag located on the next line.
		%>
		
    
		
		
		<%

	ELSE

		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		'// LANDING: enter below any code to track a customer's landing on the hone page
		'// and product details page
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

		' ENTER your tracking code BELOW the ASP tag located on the next line.
		%>
		
		
		
		
		<%

	END IF
%>