<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
' ************************************************
' * START quantity discounts in CrossSelling Item
' ************************************************
if pnoprices=0 then
	' check for discount per quantity
	query="SELECT idDiscountperquantity FROM discountsperquantity WHERE idproduct=" & cs_pidProduct
	if session("CustomerType")<>"1" then
		query=query & " and discountPerUnit<>0"
		else
		query=query & " and discountPerWUnit<>0"
	end if
	set rsDisc=Server.CreateObject("ADODB.Recordset")
	set rsDisc=conntemp.execute(query)
	if not rsDisc.eof then
		pDiscountPerQuantity=-1
	else
		pDiscountPerQuantity=0
	end if
	set rsDisc = nothing
end if
				
if pDiscountPerQuantity=-1 then 
%>
	<script language="JavaScript">
		<!--
			function win(fileName)
			{
			myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=300,height=250')
			myFloater.location.href=fileName;
			}
		//-->
	</script>
	<a href="javascript:win('priceBreaks.asp?type=<%=Session("customerType")%>&idproduct=<%=cs_pidProduct%>&type=1')"><img src="<%=rsIconObj("discount")%>" alt="<%=dictLanguage.Item(Session("language")&"_viewPrd_16")%>" style="vertical-align: middle"></a>
<% 
end if	
' ************************************************
' * END quantity discounts in CrossSelling Item
' ************************************************
%>