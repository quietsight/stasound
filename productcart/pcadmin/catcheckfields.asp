<%
	SavedFile = "importlogs/catsave.txt"
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
 	
 	Case 1: CorrectName="Category Name"
	Case 2: CorrectName="Category ID"
	Case 3: CorrectName="Small Image"
	Case 4: CorrectName="Large Image"
	Case 5: CorrectName="Parent Category Name"
	Case 6: CorrectName="Parent Category ID"
	Case 7: CorrectName="Category Short Description"
	Case 8: CorrectName="Category Long Description"
	Case 9: CorrectName="Hide Category Description"
	Case 10: CorrectName="Display Sub-Categories"
	Case 11: CorrectName="Sub-Categories per Row"
	Case 12: CorrectName="Sub-Category Rows per Page"
	Case 13: CorrectName="Use Featured Sub-Category Image"
	Case 14: CorrectName="Display Products"
	Case 15: CorrectName="Products per Row"
	Case 16: CorrectName="Product Rows per Page"
	Case 17: CorrectName="Hide category"
	Case 18: CorrectName="Hide category from retail customers"
	Case 19: CorrectName="Product Details Page Display Option"
	Case 20: CorrectName="Category Meta Tags - Title"
	Case 21: CorrectName="Category Meta Tags - Description"
	Case 22: CorrectName="Category Meta Tags - Keywords"
	Case 23: CorrectName="Featured Sub-Category Name"
	Case 24: CorrectName="Featured Sub-Category ID"
	Case 25: CorrectName="Category Order"
	End Select
 
	 exit for

	 end if
	 
	 Next
	
	 CheckField=CorrectName
    	
	End Function
%>