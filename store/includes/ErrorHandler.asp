<%

'**************************************************************
'* START - Retrieve Error Handler status from database
'* This variable is set via the Store Settings page
'**************************************************************

Call Opendb()
Dim pcIntErrorHandler
query="SELECT pcStoreSettings_ErrorHandler FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
pcIntErrorHandler=rs("pcStoreSettings_ErrorHandler")
	if pcIntErrorHandler="" then
		pcIntErrorHandler=1
	end if
set rs=nothing
Call Closedb()

'**************************************************************
'* END - Retrieve Error Handler status from database
'**************************************************************



'log file not yet used
pcStrErrFileName = "#"
'Create an id the customer can use when they call up.
pcStrCustRefID = Session.SessionID & "-" & Hour(Now) & Minute(Now) & Second(Now)

'//do not use yet...
Function LogErrorToFile ()
	Dim objFS
	Dim objFile

	On Error Resume Next
	LogError = False

	Set objFS = Server.CreateObject ("Scripting.FileSystemObject")
	If Err.number = 0 Then
		Set objFile = objFS.OpenTextFile (pcStrErrFileName, 8, True)
		If Err.number = 0 Then
			tErrDescription=	Replace(err.description,vbLf,vbCrLf)
			objFile.WriteLine "------------------------------------------------------"
			objFile.WriteLine "* Error At " & Now
			objFile.WriteLine "* CustomerRefID: "  & pcStrCustRefID			
			objFile.WriteLine "* Session ID: " & Session.SessionID
			objFile.WriteLine "* Error Number: " & err.number
			objFile.WriteLine "* Error Source: " & err.source
			objFile.WriteLine "* Error Description: " & tErrDescription
			objFile.WriteLine "* RequestMethod: " & Request.ServerVariables("REQUEST_METHOD")
			objFile.WriteLine "* ServerPort: " & Request.ServerVariables("SERVER_PORT")
			objFile.WriteLine "* HTTPS: " & Request.ServerVariables("HTTPS")
			objFile.WriteLine "* LocalAddr: "  & Request.ServerVariables("LOCAL_ADDR")
			objFile.WriteLine "* HostAddress :"  & Request.ServerVariables("REMOTE_ADDR")
			objFile.WriteLine "* UserAgent: " & Request.ServerVariables("HTTP_USER_AGENT")
			objFile.WriteLine "* URL: " &  Request.ServerVariables("URL")
			
			objFile.WriteLine "* FormData: " & Request.Form
			objFile.WriteLine "* HTTP Headers: " 
			objFile.WriteLine "*****************************"
			objFile.WriteLine Replace(Request.ServerVariables("ALL_HTTP"),vbLf,vbCrLf)
			objFile.WriteLine "*****************************"
			objFile.WriteLine "------------------------------------------------------" & vbCrLf
			objFile.Close
			
		End If
	End If
End Function

Function LogErrorToDatabase()
	
	tErrDescription = Replace(err.description,vbLf,vbCrLf)
	if instr(tErrDescription,"SQL") then
	
	else
		'// Append the query string for debugging
		if query <> "" then
			pcv_srtErrDescription = tErrDescription & "<p>" & "query=" & query & "</p>"
		else
			pcv_srtErrDescription = tErrDescription
		end if
		pcv_srtErrDescription = replace(pcv_srtErrDescription,"'","''")
		pcv_srtErrDescription = replace(pcv_srtErrDescription,"""","""""")
		
		Set conError = Server.CreateObject("ADODB.Connection")
		Set rstError = Server.CreateObject("ADODB.Recordset")
		
		conError.open scDSN
		
		if scDB="SQL" then
			strDtDelim="'"
		else
			strDtDelim="#"
		end if
		ErrQuery="INSERT INTO pcErrorHandler (pcErrorHandler_SessionID, pcErrorHandler_RequestMethod, pcErrorHandler_ServerPort, pcErrorHandler_HTTPS, pcErrorHandler_LocalAddr, pcErrorHandler_RemoteAddr, pcErrorHandler_UserAgent, pcErrorHandler_URL, pcErrorHandler_HttpHost, pcErrorHandler_HttpLang, pcErrorHandler_ErrNumber, pcErrorHandler_ErrSource, pcErrorHandler_ErrDescription, pcErrorHandler_InsertDate,pcErrorHandler_CustomerRefID) VALUES ('"&Session.SessionID&"','"&Request.ServerVariables("REQUEST_METHOD")&"','"&Request.ServerVariables("SERVER_PORT")&"','"&Request.ServerVariables("HTTPS")&"','"&Request.ServerVariables("LOCAL_ADDR")&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&Request.ServerVariables("URL")&"','"&Request.ServerVariables("HTTP_Host")&"','"&Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")&"', '"&err.number&"', '"&err.source&"', '"& pcv_srtErrDescription &"', "&strDtDelim&Date()&strDtDelim&", '"&pcStrCustRefID&"');"
		Set rstError = Server.CreateObject("ADODB.Recordset")
		Set rstError = conError.execute(ErrQuery)
		Set rstError = Nothing
		
		conError.close
		
		'If the Error Handler is turned on (1), hide the error. If it's turned off (0), show the error in the browser.
		if pcIntErrorHandler=0 then
			response.write "------------------------------------------------------<BR>"
			response.write "* Error At " & Now	&"<BR>"		
			response.write "* CustomerRefID: "  & pcStrCustRefID	&"<BR>"		
			response.write "* Session ID: " & Session.SessionID	&"<BR>"		
			response.write "* Error Number: " & err.number	&"<BR>"		
			response.write "* Error Source: " & err.source	&"<BR>"		
			response.write "* Error Description: " & tErrDescription	&"<BR>"		
			if query <> "" then
				pcv_srtErrDescription=query
				pcv_srtErrDescription = replace(pcv_srtErrDescription,"'","''")
				pcv_srtErrDescription = replace(pcv_srtErrDescription,"""","""""")
				response.write "* Last Query: " & pcv_srtErrDescription	&"<BR>"		
			end if
			response.write "* RequestMethod: " & Request.ServerVariables("REQUEST_METHOD")	&"<BR>"		
			response.write "* ServerPort: " & Request.ServerVariables("SERVER_PORT")	&"<BR>"		
			response.write "* HTTPS: " & Request.ServerVariables("HTTPS")	&"<BR>"		
			response.write "* LocalAddr: "  & Request.ServerVariables("LOCAL_ADDR")	&"<BR>"		
			response.write "* HostAddress :"  & Request.ServerVariables("REMOTE_ADDR")	&"<BR>"		
			response.write "* UserAgent: " & Request.ServerVariables("HTTP_USER_AGENT")	&"<BR>"		
			response.write "* URL: " &  Request.ServerVariables("URL")	&"<BR>"		
			response.write "* FormData: " & Request.Form	&"<BR>"		
			response.write "* HTTP Headers: " 	&"<BR>"		
			response.write "*****************************<BR>"		
			response.write Replace(Request.ServerVariables("ALL_HTTP"),vbLf,"<BR>")
			response.write "*****************************<BR>"
			response.write "------------------------------------------------------<BR>"
			
			response.end
		end if 
	end if
	err.clear
	
End Function
%>
