<%
pcv_CP_LogIP=""
pcv_CP_LogServerName=""
pcv_CP_LogReferer=""
pcv_CP_LogDate=""
pcv_CP_LogTime=""
pcv_filePath=""
pcv_Query=""
objCPFSO=""
objCPFile=""
pcStrFileName=""

'// Retrieve user info for tracking purposes
pcv_CP_LogIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If pcv_CP_LogIP = "" Then pcv_CP_LogIP = Request.ServerVariables("REMOTE_ADDR")
pcv_CP_LogSessionID=Session.SessionID
pcv_CP_LogServerName = Request.ServerVariables("SERVER_NAME")
pcv_CP_LogReferer = Request.ServerVariables("HTTP_REFERER")
pcv_CP_LogDate = Date()
pcv_CP_LogTime = Time()

'Get Path Info
pcv_filePath = Request.ServerVariables("PATH_INFO")

do while instr(pcv_filePath,"/")>0
	pcv_filePath = mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
loop

pcv_Query = Request.ServerVariables("QUERY_STRING")

if pcv_Query<>"" then
	pcv_filePath = pcv_filePath & "?" & pcv_Query
end if

'// Log current activity

Set objCPFSO = Server.CreateObject ("Scripting.FileSystemObject")
pcStrFileName = Server.Mappath ("CPLogs/"&replace(Date,"/","")&".txt")

Set objCPFile = objCPFSO.OpenTextFile (pcStrFileName, 8, True, 0)
objCPFile.WriteLine session("admin")&","&session("IDAdmin")&","&session("CUID")&","&session("PmAdmin")&","&pcv_CP_LogIP&","&pcv_CP_LogSessionID&","&pcv_CP_LogDate&","&pcv_CP_LogTime&","&pcv_CP_LogReferer&","&pcv_filePath&vbCrLf
objCPFile.Close
set objCPFSO = nothing
set objCPFile = nothing

if err.number<>0 then
	'response.redirect "techErr.asp?error="&Server.URLEncode("Permissions Not Set to Log")
	'//We will ignore permission errors so that the admin will not receive errors. We will add a message to the view logs page that alert them to a permissions issue if there are no logs to be read.
end if

err.clear
err.number=0
%>