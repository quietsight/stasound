<%
Sub CheckMinMulQty(tmpID,tmpQty)
Dim queryQ,rsQ
Dim tmpValid,tmpMin,tmpMul,tmpDesc

call opendb()
queryQ="SELECT pcprod_qtyvalidate,pcprod_minimumqty,pcProd_multiQty,Description FROM Products WHERE idProduct=" & tmpID & ";"
set rsQ=connTemp.execute(queryQ)

if not rsQ.eof then
	tmpValid=rsQ("pcprod_qtyvalidate")
	if IsNull(tmpValid) OR tmpValid="" then
		tmpValid=0
	end if
	tmpMin=rsQ("pcprod_minimumqty")
	if IsNull(tmpMin) OR tmpMin="" then
		tmpMin=0
	end if
	tmpMul=rsQ("pcProd_multiQty")
	if IsNull(tmpMul) OR tmpMul="" then
		tmpMul=0
	end if
	tmpDesc=rsQ("Description")
	set rsQ=nothing
	
	if Clng(tmpMin)>0 then
		if Clng(tmpQty)<Clng(tmpMin) then
			call closedb()
			response.Clear()
			response.redirect "msgb.asp?message="&Server.Urlencode("The quantity of "&tmpDesc&" that you are trying to order is less than the minimum quantity customers can buy. You need to buy at least "&tmpMin&" unit(s).<br><br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>" )
		end if
	end if
	
	if (Clng(tmpMul)>0) AND (Clng(tmpValid)=1) then
		if (Clng(tmpQty) Mod Clng(tmpMul))>0 then
			call closedb()
			response.Clear()
			response.redirect "msgb.asp?message="&Server.Urlencode("The product "&tmpDesc&" can only be ordered in multiples of "&tmpMul&".<br><br><br><a href=""javascript:history.go(-1)""><img src="""& rslayout("back") &""" border=0></a>" )
		end if
	end if
end if
set rsQ=nothing
End Sub
%>
	
	