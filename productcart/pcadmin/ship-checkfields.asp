<%
	SavedFile = "importlogs/ship-save.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	Flines = f.ReadAll
	f.close
	ALines=split(Flines,vbcrlf)
	
	Function CheckField(FDName)
	
		CorrectName=""
		TempStr=FDName & "*****"
		
		For dd=lbound(ALines) to ubound(ALines)
		
			if instr(ALines(dd),TempStr)>0 then
			
				Select Case dd+1
				Case 1: CorrectName="Order ID"
				Case 2: CorrectName="Ship"
				Case 3: CorrectName="Send Mail"
				Case 4: CorrectName="Ship Date"
				Case 5: CorrectName="Method"
				Case 6: CorrectName="Tracking Number"
				End Select
			
			exit for
			
			end if
		
		Next
		
		CheckField=CorrectName
	
	End Function
%>