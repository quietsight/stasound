<%
Server.ScriptTimeout = 5400
Response.Buffer = False
%>
<%
'Get Path Info
pcv_filePath=Request.ServerVariables("PATH_INFO")
do while instr(pcv_filePath,"/")>0
	pcv_filePath=mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
loop

pcv_Query=Request.ServerVariables("QUERY_STRING")

if pcv_Query<>"" then
pcv_filePath=pcv_filePath & "?" & pcv_Query
end if

' verifies if customer is logged, so as not send to login page
if Session("idCustomer")=0 then
	response.redirect "Checkout.asp?cmode=1&redirectUrl="&Server.URLEncode(pcv_filePath)
end if
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<%dim mySQL, connTemp, rsTemp    
call opendb()
DownloadID=replace(request("id"),"'","''")
if DownloadID<>"" then
 mySQL="select * from DPRequests where RequestSTR='" & DownloadID & "' and IDCustomer=" & Session("idCustomer")
 set rsTemp=connTemp.execute(mySQL)
 
 IF not rsTemp.eof then
 pIdOrder=rstemp("IDOrder")
 pIdProduct=rstemp("IDProduct")
 pProcessDate=rstemp("StartDate")
 mySQL="select * from DProducts where Idproduct=" & pIdProduct
 set rs=connTemp.execute(mySQL)
 ProductURL=rs("ProductURL")  
 pURLExpire=rs("URLExpire")
 pExpireDays=rs("ExpireDays")
 
 myTest=true
 myMsg=""
 if (pURLExpire<>"") and (pURLExpire="1") then
	 if date()-(CDate(pprocessDate)+pExpireDays)<0 then
	 else
	if date()-(CDate(pprocessDate)+pExpireDays)=0 then
	else
	myTest=false
	end if
	end if
 end if

if myTest=true then
FileN=ProductURL
FileName1=ProductURL
if (instr(ucase(FileN),"HTTP://")>0) or (instr(ucase(FileN),"HTTPS://")>0) or (instr(ucase(FileN),"FTP://")>0) then
response.redirect FileN
else
FileN=replace(FileN,"FILE:///","")
FileN=replace(FileN,"FILE://","")
FileN=replace(FileN,"file:///","")
FileN=replace(FileN,"file://","")
FileName1=FileN
if instr(FileN,"/")>0 then
myfilter="/"
else
if instr(FileN,"\")>0 then
myfilter="\"
end if
end if
Do while instr(FileName1,myfilter)>0
FileName1=mid(FileName1,instr(FileName1,myfilter)+1,len(FileName1))
loop
call downloadFile(FileN)
end if
end if
END IF
end if
%><%

function downloadFile(strFile)

' get full path of specified file
strFilename = strFile

' create stream
Set s = Server.CreateObject("ADODB.Stream")
s.Open

' set as binary
s.Type = 1

' load in the file
on error resume next


' check the file exists
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FileExists(strFilename) then
	Response.Write("<h1>Error:</h1>" & strFilename & " does not exist<p>")
	Response.End
end if


' get length of file
Set f = fso.GetFile(strFilename)
intFilelength = f.size

 
s.LoadFromFile(strFilename)
if err then
	Response.Write("<h1>Error: </h1>" & err.Description & "<p>")
	Response.End
end if

' send the headers to the users browser
Response.AddHeader "Content-Disposition", "attachment; filename=" & FileName1
'Response.AddHeader "Content-Length", intFilelength
Response.Charset = "UTF-8"
Response.ContentType = "application/octet-stream"

' output the file to the browser
Response.BinaryWrite s.Read

' tidy up
s.Close
Set s = Nothing

end function

%>