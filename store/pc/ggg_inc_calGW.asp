<%
'********************************************************************
'// Multiply Gift Wrapping charge times units of product purchased?
'// 1 = YES; 0 = NO
pcIntMultipleGiftWrap = 0
'********************************************************************

Function calGWTotal()

tmpTotal=0

pCart=Session("pcCartSession")
pCartIndex=Session("pcCartIndex")

	for f=1 to pCartIndex
		if pCart(f,10)=0 then
			GW=pCart(f,34)
			if (GW<>"") and (GW<>"0") then
				query="select pcGW_OptPrice from pcGWOptions where pcGW_IDOpt=" & GW
				set rsG=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsG=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rsG.eof then
					gOptPrice=rsG("pcGW_OptPrice")
				else
					gOptPrice=0
				end if
				set rsG=nothing
				if gOptPrice<>"" then
				else
					gOptPrice=0
				end if
				if pcIntMultipleGiftWrap <> 0 then
					tmpTotal=tmpTotal+cdbl(gOptPrice*cdbl(pCart(f,2)))
				else
					tmpTotal=tmpTotal+cdbl(gOptPrice)
				end if				
			end if
		end if
	Next

calGWTotal=tmpTotal
	
END Function %>
