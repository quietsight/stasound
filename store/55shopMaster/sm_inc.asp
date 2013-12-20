<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

Public Sub pcShowSaleIcon
	
	if pcSCID>"0" then
		query="SELECT pcSales_Completed.pcSC_ID,pcSales_Completed.pcSC_SaveName,pcSales_Completed.pcSC_SaveIcon FROM pcSales_Completed WHERE pcSales_Completed.pcSC_ID=" & pcSCID & ";"
		set rsS=Server.CreateObject("ADODB.Recordset")
		set rsS=conntemp.execute(query)

		if not rsS.eof then
			pcSCID=rsS("pcSC_ID")
			pcSCName=rsS("pcSC_SaveName")
			pcSCIcon=rsS("pcSC_SaveIcon") %>

			&nbsp;<a href="javascript:winSale('sm_showdetails.asp?id=<%=pcSCID%>')"><img src="../pc/catalog/<%=pcSCIcon%>" title="<%=pcSCName%>" alt="<%=pcSCName%>" style="vertical-align: middle"></a>
		<%end if
		set rsS=nothing
	end if
	
End Sub
%>