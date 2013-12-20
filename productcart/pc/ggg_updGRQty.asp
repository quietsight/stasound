<%'GGG Addon start
query="SELECT idproduct,quantity,pcPO_EPID FROM ProductsOrdered WHERE idorder=" & pIdOrder & " AND pcPO_EPID>0;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=ConnTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

if not rs.eof then
	pcA=rs.GetRows()
	intCount=ubound(pcA,2)
	For i=0 to intCount
		query="Update pcEvProducts set pcEP_HQty=pcEP_HQty+" &pcA(1,i)& " where pcEP_ID=" & pcA(2,i)
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		set rs=nothing
	Next
end if
set rs=nothing

'GGG Add-on end%>