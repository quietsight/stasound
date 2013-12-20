<%
	SavedFile = "importlogs/cwsave.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 1)
	Flines = f.ReadAll
	f.close
	ALines=split(Flines,vbcrlf)
	
	Function CheckField(FDName)
	
		CorrectName=""
		TempStr="*****" & FDName & "*****"
		
		For dd=lbound(ALines) to ubound(ALines)
	
			if instr(ALines(dd),TempStr)>0 then
				
				Select Case dd+1
					Case 1: CorrectName="E-mail Address"
					Case 2: CorrectName="Password"
					Case 3: CorrectName="Customer Type"
					Case 4: CorrectName="First Name"
					Case 5: CorrectName="Last Name"
					Case 6: CorrectName="Company"
					Case 7: CorrectName="Phone"
					Case 8: CorrectName="Address"
					Case 9: CorrectName="Address 2"
					Case 10: CorrectName="City"
					Case 11: CorrectName="State Code (US/Canada)"
					Case 12: CorrectName="Province"
					Case 13: CorrectName="Postal Code"
					Case 14: CorrectName="Country Code"
					Case 15: CorrectName="Shipping Company"
					Case 16: CorrectName="Shipping Address"
					Case 17: CorrectName="Shipping Address 2"
					Case 18: CorrectName="Shipping City"
					Case 19: CorrectName="Shipping State Code (US/Canada)"
					Case 20: CorrectName="Shipping Province"
					Case 21: CorrectName="Shipping Postal Code"
					Case 22: CorrectName="Shipping Country Code"
					Case 23: CorrectName="Current Reward Points Balance"
					Case 24: 
					if (CIView1="1") and (CILabel1<>"") then
					CorrectName="Current Reward Points Balance"
					else
					CorrectName=""
					end if
					Case 25:
					if (CIView2="1") and (CILabel2<>"") then
					CorrectName="Current Reward Points Balance"
					else
					CorrectName=""
					end if	
					Case 26: CorrectName="Newsletter Subscription"
					Case 27: CorrectName="Pricing Category ID"
					Case 28: CorrectName="Fax"
					Case 29: CorrectName="Shipping Email Address"
					Case 30: CorrectName="Shipping Phone"
					'MailUp-S
					Case 31: CorrectName="Opt-in MailUp List IDs"
					Case 32: CorrectName="Opt-out MailUp List IDs"
					'MailUp-E
				End Select
		 
			 exit for
		
			 end if
		 
		 Next
		
		 CheckField=CorrectName
    	
	End Function
%>