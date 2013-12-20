<%
'//Sanitize Uploaded File
Function IsUploadAllowed(strFileName)
Dim tmpFileName
	tmpFileName=strFileName
	tmpFileName=Right(tmpFileName,Len(tmpFileName)-InstrRev(tmpFileName,"\"))
	IsUploadAllowed = True
	BlackList= array(";", ":", ">", "<", "/" ,"\", "..", "?", "%", "$", "#","&")
	TempStr = trim(tmpFileName)

	for i=lbound(BlackList) to ubound(BlackList)
 		if (instr(1,TempStr,BlackList(i),vbTextCompare)<>0) then
			
			IsUploadAllowed = false
			Exit Function
 		end if
 	next
End Function
%>

