<%
Function pcv_GetMSXML()
	Dim i
	For i = 3 to 6
		On Error Resume Next
		Set pcv_GetMSXML = Server.CreateObject("Msxml2.DOMDocument." & CStr(i) & ".0")
		If Err.Number = 0 Then Exit For
	Next
End Function

Sub pcv_AddNode(ToNode, Name, Value, NodeType)
	Dim Element, sValue
	
	sValue = Null
	
	If IsNull(Value) Then
		sValue = ""
	Else
		sValue = CStr(Value)
	End If
	
	If Not IsNull(sValue) Then
		If NodeType = 3 Then
			Set Element = XMLDoc.createElement(Name)
			ToNode.appendChild Element
			Element.appendChild XMLDoc.createTextNode(sValue)
		ElseIf NodeType = 4 Then
			Set Element = XMLDom.createElement(Name)
			ToNode.appendChild Element
			Element.appendChild XMLDoc.createCDATASection(sValue)
		ElseIf NodeType = 2 Then
			ToNode.SetAttribute Name, sValue
		End If
	End If

End Sub

' Report error
Function pcv_GenXMLError()
	pcv_GenXMLError = "Error Number: " & Err.Number & "; " & _
		"Description: " & Err.Description & "; Source: " & Err.Source
End Function

%>