<%
query="SELECT pcOrd_GCDetails FROM Orders where idOrder=" & pIdOrder
set rsG=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsG=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
GCDetails=rsG("pcOrd_GCDetails")
set rsG=nothing

IF GCDetails<>"" THEN
	GCArr=split(GCDetails,"|g|")
	intGCCount=ubound(GCArr)
	For ik=0 to intGCCount
	IF GCArr(ik)<>"" THEN
	GCInfo=split(GCArr(ik),"|s|")
	
	query="SELECT pcGCOrdered.pcGO_ExpDate,pcGCOrdered.pcGO_Amount,pcGCOrdered.pcGO_Status,products.Description FROM pcGCOrdered,products WHERE pcGCOrdered.pcGO_GcCode='"&GCInfo(0)&"' AND products.idproduct=pcGCOrdered.pcGO_IDProduct"
	set rsG=Server.CreateObject("ADODB.Recordset")
	set rsG=conntemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rsG=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	IF rsG.eof then
		pDiscountCode=""
		set rsG=nothing
	ELSE
		mTest=0
		pGCExpDate=rsG("pcGO_ExpDate")
		pGCAmount=rsG("pcGO_Amount")
		if pGCAmount<>"" then
		else
			pGCAmount="0"
		end if
		pGCStatus=rsG("pcGO_Status")
		pDiscountDesc=rsG("Description")
	
		if cdbl(pGCAmount)<=0 then
			mTest=1
		end if
		if cint(pGCStatus)<>1 then
			mTest=1
		end if
		if year(pGCExpDate)<>"1900" then
			if Date()>pGCExpDate then
				mTest=1
			end if
		end if
	
		if mTest=0 then
		'Have Available Amount

			GCAmount=cdbl(GCInfo(2))
			if GCAmount<>"" then
			else
				GCAmount="0"
			end if

			pGCAmount=pGCAmount-GCAmount
			if pGCAmount<0 then
				pGCAmount=0
			end if
			if pGCAmount=0 then
				pGCStatus=0
			end if
	
			query="update pcGCOrdered set pcGO_Amount=" & pGCAmount & ",pcGO_Status=" & pGCStatus & " where pcGO_GcCode='"&GCInfo(0)&"'"
			set rs=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs=nothing
		end if
	END IF 'rsG.eof
	set rsG=nothing
	
	END IF 'GcArr(ik)<>""
	Next
END IF
%>
