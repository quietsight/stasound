<% dim strXML, xChecked, xErrorDesc, xErrorConnDesc, xErrorPostDesc, xTestURL, arrayVersion, spltArryVersion, xCnt, xVer, xErrorExist, x

strXML = "<?xml version=""1.0"" encoding=""UTF-16""?><cjb></cjb>"

xChecked = ""
xErrorDesc = "Installed"
xErrorConnDesc = ""
xErrorPostDesc = ""

xTestURL="http://www.earlyimpact.com/productcart/verify/test.asp"

arrayVersion="1, 2, 3, 4, 5, 6"
'//split array
spltArryVersion=split(arrayVersion, ", ")
xSetVer="0"
for xCnt=lbound(spltArryVersion) to ubound(spltArryVersion)
	select case spltArryVersion(xCnt)
		case "1"
			xVer=""
		
		case "2"
			xVer=".2.6"
		
		case "3"
			xVer=".3.0"
		
		case "4"
			xVer=".4.0"
		
		case "5"
			xVer=".5.0"
		
		case "6"
			xVer=".6.0"
		
	end select

	err.clear
	xErrorExist=0
	Set x = server.CreateObject("Msxml2.DOMDocument"&xVer)
	x.async = false 
	if x.loadXML(strXML) then
		xChecked="YES"
	end if
	
	set x=nothing
	
	if err.number<>0 then
		xErrorExist=1
		xErrorConnDesc=err.description
		xChecked=""
		err.clear
	else
		Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
		srvXmlHttp.open "POST", xTestURL, false
		if err.number<>0 then
			xErrorExist=1
			xErrorPostDesc=err.description
			err.clear
		else
			srvXmlHttp.send(xml)
			if err.number<>0 then
				xErrorExist=1
				xErrorSendDesc=err.description
				err.clear
			end if
		end if
		set srvXmlHttp=nothing
	end if

	if xChecked="YES" then
		if xErrorExist=0 then 
			'SET XML
			xSetVer=xVer
			Exit for
		end if
	end if
next

'response.write "xSetVer: "&xSetVer
	
%>               

