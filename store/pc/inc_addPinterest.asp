<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Pinterest Pin It Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

	'// SETTINGS-START
		'// Active: 1 = active, 0 = inactive
		if scPinterestDisplay&""="" then
			pcInterest=0
		else
			pcInterest=scPinterestDisplay
		end if
		
		'// Counter visibility- Options are: horizontal, vertical, none
		if scPinterestCounter&""="" then
			pcPinItCounter="none"
		else
			pcPinItCounter=scPinterestCounter
		end if
		
	'// SETTINGS-END
	
	
	tempURLp=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	tempURLp=replace(tempURLp,"http:/","http://")
	
	'// Product image
	if trim(pImageUrl)="" then
		pImageUrl=pSmallImageUrl
	end if
	
	if pcInterest=1 then
	%>
    <a href="http://pinterest.com/pin/create/button/?url=<%=Server.Urlencode(tempURLp&pcStrPrdLink)%>&media=<%=Server.Urlencode(tempURLp&pcv_tmpNewPath & "catalog/" & pImageUrl)%>&description=<%=Server.Urlencode(pDescription)%>" class="pin-it-button" count-layout="<%=pcPinItCounter%>" target="_blank"><img border="0" src="//assets.pinterest.com/images/PinExt.png" title="Pin It" /></a>
	<%
	end if
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Pinterest Pin It Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>