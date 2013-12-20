
<%
Public Sub pcs_JavaTextField(FieldName,isRequiredField,ErrorMessage)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=FieldName	
	if isRequiredField then
		response.write "if (theForm."&pcv_strFieldName&".value == """")"&vbcrlf
		response.write "	{"&vbcrlf
		response.write "alert("""&ErrorMessage&""")"&vbcrlf
		response.write "theForm."&pcv_strFieldName&".focus();"&vbcrlf
		response.write "return (false);"&vbcrlf
		response.write "}"&vbcrlf
	end if	
End Sub


Public Sub pcs_JavaDropDownList(FieldName,isRequiredField,ErrorMessage)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=FieldName	
	if isRequiredField then
		response.write "if (theForm."&pcv_strFieldName&".selectedIndex == 0)"&vbcrlf
		response.write "{"&vbcrlf
		response.write "alert("""&ErrorMessage&""")"&vbcrlf
		response.write "theForm."&pcv_strFieldName&".focus();"&vbcrlf
		response.write "return (false);"&vbcrlf
		response.write "}"&vbcrlf
	end if	
End Sub

Public Sub pcs_JavaCheckedBox(FieldName,isRequiredField,ErrorMessage)
	Dim pcv_strFieldName
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=FieldName	
	if isRequiredField then
		response.write "if (theForm."&pcv_strFieldName&".checked == ""0"")"&vbcrlf
		response.write "{"&vbcrlf
		response.write "alert("""&ErrorMessage&""")"&vbcrlf
		response.write "theForm."&pcv_strFieldName&".focus();"&vbcrlf
		response.write "return (false);"&vbcrlf
		response.write "}"&vbcrlf
	end if	
End Sub
'if (chk.checked == 1)


Public Sub pcs_JavaCompare(FieldName,FieldName2,isRequiredField,ErrorMessage)
	Dim pcv_strFieldName, pcv_strFieldName2
	if isRequiredField = "" then
		isRequiredField = true
	end if
	pcv_strFieldName=FieldName	
	pcv_strFieldName2=FieldName2
	if isRequiredField then
		response.write "if (theForm."&pcv_strFieldName&".value !== theForm."&pcv_strFieldName2&".value)"&vbcrlf
		response.write "{ "&vbcrlf
		response.write "alert("""&ErrorMessage&""")"&vbcrlf
		response.write "theForm."&pcv_strFieldName&".focus()"&vbcrlf
		response.write "return (false);"&vbcrlf
		response.write "}"&vbcrlf
	end if	
End Sub
%>

