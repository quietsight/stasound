<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
dim initialize
initialize=0

dim query, conntemp, rsTemp

pIdDbSession=session("pcSFIdDbSession")
pRandomKey=session("pcSFRandomKey")

HaveToRefeshCustomerCache=""

' if dbSession was not defined
if pIdDbSession="" or pRandomKey="" then
	initialize=-1
end if

' check if current pcCustomerSessions is valid
if initialize=0 AND HaveToRefeshCustomerCache<>"1" then
	pcCustSession_Date=Date()
	if SQL_Format="1" then
		pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	else
		pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	end if

	call opendb()
	if scDB="Access" then
		TmpQuery="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND pcCustSession_Date=#" &pcCustSession_Date& "# ORDER BY pcCustomerSessions.idDbSession DESC;"
	else
		' SQL Server and other DBS use ' instead of # to filter by date
		TmpQuery="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND pcCustSession_Date='" &pcCustSession_Date& "' ORDER BY pcCustomerSessions.idDbSession DESC;"
 	end if
	set rsTmpObj=Server.CreateObject("ADODB.Recordset")
	set rsTmpObj=conntemp.execute(TmpQuery)
	if rsTmpObj.eof then
		' invalid pcCustomerSessions
		response.redirect "msg.asp?message=38"
	end if
	set rsTmpObj=nothing
	call closeDb()
end if

if initialize=-1 OR HaveToRefeshCustomerCache="1" then
	pRandomKey=randomNumber(99999999)
	session("pcSFRandomKey")=pRandomKey
	pcCustSession_Date=Date()
	if SQL_Format="1" then
		pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	else
		pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	end if
	call opendb()
	if scDB="Access" then
 		query="INSERT INTO pcCustomerSessions (randomKey, idCustomer, pcCustSession_Date) VALUES (" &pRandomKey& ","&session("idCustomer")&", #" &pcCustSession_Date& "#)"
	else
 		query="INSERT INTO pcCustomerSessions (randomKey, idCustomer, pcCustSession_Date) VALUES (" &pRandomKey& ","&session("idCustomer")&", '" &pcCustSession_Date& "')"
	end if

	set rs=Server.CreateObject("ADODB.Recordset")
 	set rs=conntemp.execute(query)

 	if err.number <> 0 then
		set rs=nothing
		call closeDb()
		response.redirect "techErr.asp?error="& Server.Urlencode("Error in DBsv 3: "&err.description) 
	end if

 	' get pcCustomerSessions 
	if scDB="Access" then
		query="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND idCustomer="&session("idCustomer")&" AND pcCustSession_Date=#" &pcCustSession_Date& "# ORDER BY idDbSession DESC;"
	else
		' SQL Server and other DBS use ' instead of # to filter by date
		query="SELECT idDbSession FROM pcCustomerSessions WHERE randomKey="&pRandomKey& " AND idCustomer="&session("idCustomer")&" AND pcCustSession_Date='" &pcCustSession_Date& "' ORDER BY idDbSession DESC;"
 	end if
 	set rs=conntemp.execute(query)
 	pIdDbSession=rs("idDbSession")
	session("pcSFIdDbSession")=pIdDbSession
	set rs=nothing
	call closedb()
end if

if session("idCustomer")>"0" then
	call opendb()
	query="UPDATE pcCustomerSessions SET idCustomer="&session("idCustomer")&" WHERE randomKey="&pRandomKey& " AND idDbSession=" & pIdDbSession & ";"
	set rs=connTemp.execute(query)
	set rs=nothing
	call closedb()
end if


' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+2
end function
%>