<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: AddThis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_AddThis
	Dim rs,query,pcStrAddThisCode
	query="SELECT pcStoreSettings_AddThisCode FROM pcStoreSettings WHERE (((pcStoreSettings_ID)=1));"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcStrAddThisCode=rs("pcStoreSettings_AddThisCode")
	end if
	if trim(pcStrAddThisCode)<>"" and not IsNull(pcStrAddThisCode) then
		if scAddThisDisplay=1 then
			response.write "<div class=""pcAddThisFloat"">" & pcStrAddThisCode & "</div>"
		else
			response.write "<div class=""pcAddThis"">" & pcStrAddThisCode & "</div>"
		end if
	end if
	set rs=nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: AddThis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>