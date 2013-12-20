<%

Function CheckOFS(chkID,chkQty,chkStock,exLine)
Dim rsQ,queryQ,pcCartArray,ppcCartIndex,ck,cm,tmpQty

pcCartArray	= Session("pcCartSession")
ppcCartIndex = Session("pcCartIndex")

tmpQty=chkQty

for ck=1 to ppcCartIndex
	'checking stock level
	if Clng(exLine)<>Clng(ck) then
		if (pcCartArray(ck,10)=0) AND (Clng(pcCartArray(ck,0))=Clng(chkID)) then         
			tmpQty=Clng(pcCartArray(ck,2))+Clng(tmpQty)
		end If
	end if
next

if Clng(tmpQty)>Clng(chkStock) then
	CheckOFS=1
	exit function
end if

for ck=1 to ppcCartIndex
	'checking stock level
	if Clng(exLine)<>Clng(ck) then
		if (pcCartArray(ck,10)=0) AND (pcCartArray(ck,16)<>"") AND (pcCartArray(ck,16)>"0") then         
			queryQ="SELECT stringProducts, stringQuantity FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(ck,16)) & " AND ((stringProducts like '" & chkID & ",%') OR (stringProducts like '%," & chkID & ",%'));"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tmpSP=split(rsQ("stringProducts"),",")
				tmpSQ=split(rsQ("stringQuantity"),",")
				set rsQ=nothing
				For cm=lbound(tmpSP) to ubound(tmpSP)
					if trim(tmpSP(cm))<>"" then
						if Clng(tmpSP(cm))=Clng(chkID) then
							tmpQty=Clng(tmpSQ(cm))*Clng(pcCartArray(ck,2))+Clng(tmpQty)
						end if
					end if
				Next	
			end if
			set rsQ=nothing
		end If
	end if
next

if Clng(tmpQty)>Clng(chkStock) then
	CheckOFS=1
else
	CheckOFS=0
end if

End Function

Sub CheckALLCartStock()
Dim rsQ,queryQ,pcCartArray,ppcCartIndex,cl,ci,queryC,rsC

pcCartArray	= Session("pcCartSession")
ppcCartIndex = Session("pcCartIndex")

if scOutofstockpurchase=-1 then
	call opendb()
	for cl=1 to ppcCartIndex
		if (pcCartArray(cl,10)=0) then
			if (pcCartArray(cl,16)="") OR (pcCartArray(cl,16)="0") OR (iBTOOutofstockpurchase=-1) then
				'Check Product Stock
				queryC="SELECT idProduct,stock,description FROM Products WHERE idProduct=" & pcCartArray(cl,0) & " AND (nostock=0) AND (pcProd_BackOrder=0);"
				set rsC=ConnTemp.execute(queryC)
				if not rsC.eof then
					tmpID=rsC("idProduct")
					tmpiStock=rsC("stock")
					tmpiDesc=rsC("description")
					set rsC=nothing
					if CheckOFS(tmpID,pcCartArray(cl,2),tmpiStock,cl)=1 then
						call closedb()
						response.Clear()
						response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&tmpiDesc&" that you are trying to order is greater than the quantity that we currently have in stock. We currently have "&tmpiStock&" unit(s) in stock.<br><br><br><a href=""javascript:history.go(-1)"">Back</a>" )
					end if
				end if
				set rsC=nothing
			end if
			if (pcCartArray(cl,16)<>"") AND (pcCartArray(cl,16)>"0") then         
			queryQ="SELECT stringProducts, stringQuantity FROM configSessions WHERE idconfigSession=" & trim(pcCartArray(cl,16)) & " AND idProduct=" & pcCartArray(cl,0) & ";"
			set rsQ=connTemp.execute(queryQ)
			if not rsQ.eof then
				tmpSP=split(rsQ("stringProducts"),",")
				tmpSQ=split(rsQ("stringQuantity"),",")
				set rsQ=nothing
				For ci=lbound(tmpSP) to ubound(tmpSP)
					if trim(tmpSP(ci))<>"" then
						'Check Product Stock
						queryC="SELECT idProduct,stock,description FROM Products WHERE idProduct=" & tmpSP(ci) & " AND (nostock=0) AND (pcProd_BackOrder=0);"
						set rsC=ConnTemp.execute(queryC)
						if not rsC.eof then
							tmpID=rsC("idProduct")
							tmpiStock=rsC("stock")
							tmpiDesc=rsC("description")
							set rsC=nothing
							if CheckOFS(tmpSP(ci),Clng(tmpSQ(ci))*Clng(pcCartArray(cl,2)),tmpiStock,cl)=1 then
								call closedb()
								response.Clear()
								response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&tmpiDesc&" that you are trying to order is greater than the quantity that we currently have in stock. We currently have "&tmpiStock&" unit(s) in stock.<br><br><br><a href=""javascript:history.go(-1)"">Back</a>" )
							end if
						end if
						set rsC=nothing
					end if
				Next	
			end if
			set rsQ=nothing
			end if
		end If
	next
end if
End Sub

%>