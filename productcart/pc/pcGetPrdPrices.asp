<%
if pBtoBPrice=0 then
	pBtoBPrice=pPrice
end if

dblpcCC_Price=0
pPrice1=pPrice
pBtoBPrice1=pBtoBPrice
If pnoprices<2 Then
		if pserviceSpec=false then
			pPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,0)
			pBtoBPrice1=CheckParentPrices(pidProduct,pPrice,pBtoBPrice,1)
		else
			query="SELECT pcProd_BTODefaultPrice,pcProd_BTODefaultWPrice FROM Products WHERE idproduct=" & pIdProduct & ";"
			set rsQ=server.CreateObject("ADODB.RecordSet")
			set rsQ=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsQ=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			if not rsQ.eof then
				pPrice=rsQ("pcProd_BTODefaultPrice")
				pBtoBPrice=rsQ("pcProd_BTODefaultWPrice")
				if pBtoBPrice=0 then
					pBtoBPrice=pPrice
				end if
			end if
			set rsQ=nothing
			pPrice1=pPrice
			pBtoBPrice1=pBtoBPrice
			if session("customerCategory")<>"" AND session("customerCategory")<>"0" then
				query="SELECT pcBDPC_Price FROM pcBTODefaultPriceCats WHERE idproduct=" & pIdproduct & " AND idCustomerCategory=" & session("customerCategory") & ";" 
				set rsQ=server.CreateObject("ADODB.RecordSet")
				set rsQ=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsQ=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rsQ.eof then
					pBtoBPrice1=rsQ("pcBDPC_Price")
					pPrice1=pBtoBPrice1
				end if
				set rsQ=nothing
			end if
			if pcPageStyle = "m" then
				call pcs_GetBTOConfigPrices(pCnt,pcPageStyle)
			end if
		end if
End if
		
		if session("customertype")=1 and pBtoBPrice1>0 then
			dblpcCC_Price=pBtoBPrice1
		else
			dblpcCC_Price=pPrice1
		end if
%>
