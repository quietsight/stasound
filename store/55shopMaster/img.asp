<%
Dim SPath
SPath=Request.ServerVariables("PATH_INFO")
SPath=mid(SPath,1,InStrRev(SPath,"/")-1)
If UCase(Trim(Request.ServerVariables("HTTPS")))="OFF" then
	strSiteURL="http://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
Else
	strSiteURL="https://" & Request.ServerVariables("HTTP_HOST") & SPath & "/"
End if

	Randomize
	tmp1=fix(10*Rnd)
	tmp3=session("cp_num")
	tmp2=mid(tmp3,1,request("a")-1) & tmp1 & mid(tmp3,request("a")+1,len(tmp3))
	session("cp_num")=tmp2
	Response.ContentType = "image/gif"
	Set objHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP"&scXML)
	objHTTP.open "GET", strSiteURL & "images/" & tmp1 & ".gif",false
	objHTTP.send
	Response.BinaryWrite objHTTP.ResponseBody
	Set objHTTP = Nothing
%>
