<%

' Use this class to concatenate strings in a much more
' efficient manner than simply concatenating a string
' (strVariable = strVariable & "your new string")
Class StringBuilder
	Dim arr 	'the array of strings to concatenate
	Dim growthRate  'the rate at which the array grows
	Dim itemCount   'the number of items in the array

	Private Sub Class_Initialize()
		growthRate = 50
		itemCount = 0
		ReDim arr(growthRate)
	End Sub

	'Append a new string to the end of the array. If the
	'number of items in the array is larger than the
	'actual capacity of the array, then "grow" the array
	'by ReDimming it.
	Public Sub Append(ByVal strValue)
		If itemCount > UBound(arr) Then
			ReDim Preserve arr(UBound(arr) + growthRate)
		End If

		arr(itemCount) = strValue
		itemCount = itemCount + 1
	End Sub

	'Concatenate the strings by simply joining your array
	'of strings and adding no separator between elements.
	Public Function ToString() 
		ToString = Join(arr, "")
	End Function
End Class
%>