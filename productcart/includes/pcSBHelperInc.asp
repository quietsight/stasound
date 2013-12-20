<%
'///////////////////////////////////////////////////////////////////////////////////////
'// START: SB pcIsSubscription Function 
'///////////////////////////////////////////////////////////////////////////////////////
function findSubscription(pcCartArray, indexCart)
	Dim f,Subscription
		
	Subscription = False
	for f=1 to indexCart
		if  (pcCartArray(f,38)>0)  Then
			 Subscription = True 			
			 Exit For				   
		End if 				
    Next
  findSubscription = Subscription
end function
'///////////////////////////////////////////////////////////////////////////////////////
'// END: SB pcIsSubscription Function 
'///////////////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////////////////////////////////////
'// START: SB pcIsSubscription Function 
'///////////////////////////////////////////////////////////////////////////////////////
function IsCartLockable(pcCartArray, indexCart)
	Dim f,Subscription
		
	IsTaxOrShipping = False
	for f=1 to indexCart
		if  (pcCartArray(f,19)<>-1) OR (pcCartArray(f,20)<>-1) Then
			 IsTaxOrShipping = True 			
			 Exit For				   
		End if 				
    Next
  	IsCartLockable = IsTaxOrShipping	
end function
'///////////////////////////////////////////////////////////////////////////////////////
'// END: SB pcIsSubscription Function 
'///////////////////////////////////////////////////////////////////////////////////////		


'///////////////////////////////////////////////////////////////////////////////////////
'// START: SB getSubcription price Function  
'///////////////////////////////////////////////////////////////////////////////////////
Function getSubInstallVals(psubscriptionID)

	'// Open Private Connection String
	Set connSB=Server.CreateObject("ADODB.Connection")
	connSB.Open scDSN
	
	query="SELECT SB_BillingCycles, SB_IsTrial, SB_TrialAmount, SB_Type FROM SB_Packages WHERE SB_PackageID=" & psubscriptionID
	set rsSub=server.CreateObject("ADODB.RecordSet")
	set rsSub=connSB.execute(query)	

	if not rsSub.eof then

		'// Billing Cycles
		pcv_intBillingCycles = rsSub("SB_BillingCycles") '// rsSub("pcSubscription_TotalOccur")
		if isnull(pcv_intBillingCycles) or pcv_intBillingCycles ="" Then 
			pcv_intBillingCycles = "0"
		End if 
		
		'// Is Trial?
	    pcv_intIsTrial = rsSub("SB_IsTrial")	
	    if pcv_intIsTrial = "1" then				 
		  pcv_curTrialAmount = rsSub("SB_TrialAmount")
		  if isnull(pcv_curTrialAmount) or pcv_curTrialAmount = "" then 
			pcv_curTrialAmount = 0
		  end if 
		end if 
		
		pSubType = rsSub("SB_Type")
		if isnull(pSubType) or pSubType ="" Then 
			pSubType = 0
		End if 

	else
	  	set rsSub = nothing	  
	  	getSubInstallVals ="0,0,0,0"
	  	exit function 
	end if 

	getSubInstallVals = pcv_intBillingCycles & "," & pSubInstall & "," & pcv_intIsTrial & "," & pcv_curTrialAmount
	
	'// Close Private Connection String
	connSB.Close
	Set connSB=nothing
	
End Function

'///////////////////////////////////////////////////////////////////////////////////////
'// START: SB getSubcription price Function 
'///////////////////////////////////////////////////////////////////////////////////////
%>



<%
'// Flat File Methods
Function writeFile(pageName,pText)	

		if PPD="1" then
 			findit=Server.MapPath("/"&scPcFolder&"/Includes/"&PageName)
		else
			findit=Server.MapPath("../Includes/"&PageName)
		end if
        
		
		on error resume next

		Set fso=server.CreateObject("Scripting.FileSystemObject")
				
		if fso.FileExists(findit) then
			Set f=fso.OpenTextFile(findit, 2, True)
			f.Write(pText)			
			if Err.number>0 then
				response.redirect "techErr.asp?error="&Server.URLEncode("Error Updating File " &pageName &  "." & err.description)
			end if
			
		Else
			Set f = fso.CreateTextFile(findit,true)			
			f.Write(pText)
			if Err.number>0 then
				response.redirect "techErr.asp?error="&Server.URLEncode("Error Writting File " &pageName &  ".")
			end if			
		End if 
		f.close
		
		set f=nothing
		set fso =nothing
End function

Function removeFile(pageName)

		if PPD="1" then
 			findit=Server.MapPath("/"&scPcFolder&"/Includes/"&PageName)
		else
			findit=Server.MapPath("../Includes/"&PageName)
		end if
		
		on error resume next

		Set fso=server.CreateObject("Scripting.FileSystemObject")
		
		if fso.FileExists(findit) then			
			fso.DeleteFile(findit)
			if Err.number>0 then
				response.redirect "techErr.asp?error="&Server.URLEncode("Error Removing File " &pageName &  ". ")
			end if
			
		End if 
		
		set fso=nothing

End Function

Function getFile(pgName)
     Dim fText, fso, f

    
     	if PPD="1" then
 			findit=Server.MapPath("/"&scPcFolder&"/Includes/"&PgName)
		else
			findit=Server.MapPath("../Includes/"&PgName)
		end if	
		
		on error resume next
   
		Set fso=server.CreateObject("Scripting.FileSystemObject")
		
		if fso.FileExists(findit) then		
			Set f=fso.OpenTextFile(findit, 1, false,0)
			 fText= f.readall
			 f.close  
			 set f=nothing
			if Err.number>0 then			  
				response.redirect "techErr.asp?error="&Server.URLEncode("Error Reading File " &pageName &  ". ")
			end if			
		End if 
	    set fso = nothing
		getFile = fText		

End Function



Function getSubDate(subUnit,dFormat,dInc,dtDate,dfunc)
	' subunite:  Period of days or months will set sbdInterval to yy defualt is m 
	' dformat: Date format
	' dInc: Number of Cycles
	' dtDate: Start Date
	' dfunct: Tell us to add date other wise just format and return 
	Dim rtnDate 
	if not isdate(dtDate) then 
		dtDate = Date() 
	end if 
	if year(dtDate)=1990  Then 
		 dtDate = Date() 
	end if   


	If Ucase(dFunc) = "ADD" then
		'// sbdInterval defaults to months m
		sbdInterval ="m"
	 	if SubUnit = "days" then
			sbdInterval  =  "d" 
	 	End if 
		'response.Write(sbdInterval & "<br />")
		'response.Write(dInc & "<br />")
		'response.Write(dtDate & "<br />")
		'response.End()
	 	rtnDate  = dateadd(sbdInterval, dInc, dtDate) 
	Else
		rtnDate = dtDate
	End if 	 


	 if dFormat="DD/MM/YY" then
		 rtnDate=day(rtnDate) & "/" & month(rtnDate) & "/" & year(rtnDate)					
	 else
		 rtnDate=month(rtnDate) & "/" & day(rtnDate) & "/" & year(rtnDate)
	 end if		 
	
	 
	 getSubDate = rtnDate
	
 End Function
%>