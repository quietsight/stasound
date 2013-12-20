<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<% 
on error resume next

File1=request.Querystring("file1")
if ucase(File1)="HTTP://" then
	File1=""
end if

	if File1<>"" then
		if (instr(ucase(File1),"HTTP://")>0) or (instr(ucase(File1),"HTTPS://")>0) or (instr(ucase(File1),"FTP://")>0) then
			response.redirect File1
		else
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			Err.number=0
			Set f = fso.OpenTextFile(File1, 1)
			myErr1="Tested successfully!"
			if Err.number>0 then
				myErr1= err.Description
				err.number=0
				err.Description=""
			end if
		end if
	end if
	
	%> 
    <html>
    <head>
    <title>Check File Location</title>
    </head>
    <body>
    <b><font face="Arial" size="4">Check File Locations</font></b><font face=Arial size=2><br><br>
	<%if File1<>"" then%>
	<strong>Downloadable File Location</strong>:<br>
	<%=File1%><br><br>
    <strong>Result</strong>:<br><%=myErr1%>
	<%end if%>
	</font>
    </body>
    </html>