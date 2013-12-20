<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'*******************************
' START Verify Session
'*******************************
Dim pcCartIndex
Private Sub pcs_VerifySession
	Dim pcv_strCatcher
	pcv_strCatcher = Session("pcCartIndex")
	If pcv_strCatcher=0 Then
		pcv_strCatcher=""		
		pcv_strCheckSession = getUserInput(Request("cs"),1)
		if (len(pcv_strCheckSession)>0) AND (session("pcSessionID") <> Session.SessionID) then
		 	response.redirect "msg.asp?message=212" '// enable cookies
		else
			response.redirect "msg.asp?message=1" '// cart empty
		end if
	End If
	pcCartArray=Session("pcCartSession")
	pcCartIndex=Session("pcCartIndex")
End Sub
'*******************************
' START Verify Session
'*******************************
%>