<%				
'SB S

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START - DEFAULTS
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// Plan - Regular
pcv_intBillingFrequency = 0 '// The Subscription Length
pcv_strBillingPeriod = "" '// Subscription Billing Period
pcv_intBillingCycles = 0 '// Number of Subscription Occurances

'// Plan - Trial
pcv_intTrialFrequency = 0 '// The Trial Subscription Length
pcv_strTrialPeriod = "" '// Trial Subscription Billing Period
pcv_intTrialCycles = 0 '// Number of Trial Subscription Occurances

'// Plan - Misc.
pcv_intIsTrial = 0
pSubStartImmed = 0

'// Package
'pcv_curAmount = 0  '// Price (comes from cart)
pcv_curTrialAmount = 0 '// Trial Price

'// Installment (coming later)
'pSubStartFromPurch = ""
'pSubStart = ""

'// Agreement
pSubAgree = "1"

'// Display Setings
pShowTrialPrice = 0
pTrialDesc = "" 
pShowFreeTrial = 0
pShowStartDate = 0
pStartDateDesc = "" 
pShowReoccurenceDate = 0 
pReoccurenceDesc = "" 
pShowEOSDate = 0
pEOSDesc = ""
pShowTrialDate = 0 
pFreeTrialDesc = "" 
			
'// Installment (coming later)
'pShowInstallment=""
'pInstallmentDesc=""

'// Link ID
pcv_strLinkID = ""

'// Type
'// 0 = Infinite
'// 1 = Fixed 1-999
'// 2 = Installment
pSubType = 0	

'// Is Installment?
pSubInstall = 0

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END - DEFAULTS
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// START - GET PACKAGE DATA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If pSubscriptionID <> "0" Then

	query="SELECT * FROM SB_Packages WHERE SB_PackageID=" & pSubscriptionID
	set rsSub=server.CreateObject("ADODB.RecordSet")
	set rsSub=connTemp.execute(query)	
		
	if not rsSub.eof then	 
	    
		pIsLinked = rsSub("SB_IsLinked")

		If pIsLinked="0" Then
		
			'// Plan - Regular
			pcv_intBillingFrequency = rsSub("SB_BillingFrequency") '// The Subscription Length
			pcv_strBillingPeriod = rsSub("SB_BillingPeriod") '// Subscription Billing Period
			pcv_intBillingCycles = rsSub("SB_BillingCycles") '// Number of Subscription Occurances

			'// Plan - Trial
			pcv_intTrialFrequency = rsSub("SB_BillingFrequency") '// The Trial Subscription Length
			pcv_strTrialPeriod = rsSub("SB_BillingPeriod") '// Trial Subscription Billing Period
			pcv_intTrialCycles = rsSub("SB_TrialBillingCycles") '// Number of Trial Subscription Occurances
			
		  	'// Plan - Misc.
			pcv_intIsTrial = rsSub("SB_IsTrial") '// Is Trial - True OR False		  
		  	'pSubStartImmed = 0 '// rsSub("pcSubscription_StartImmed")   '// Start Immediately (TO DO)

		  	'// Package
			'pcv_curAmount = 0  '// Price (comes from cart)
			pcv_curTrialAmount = rsSub("SB_TrialAmount") '// Trial Price

			'// Installment (coming later)
		  	'pSubStartFromPurch=rsSub("pcSubscription_StartFromPurch")
		  	'pSubStart=rsSub("pcSubscription_StartDate")
			
			'// Agreement
			pSubAgree = "1" '// rss("SB_Agree")
			'pRegAgree = rstemp("SB_Agree")
			'pAgreeText = rstemp("SB_AgreeText")

			'// Display Setings
			pShowTrialPrice = rsSub("SB_ShowTrialPrice")
			pTrialDesc = rsSub("SB_TrialDesc")
			pShowFreeTrial = rsSub("SB_ShowFreeTrial")
			pShowStartDate = rsSub("SB_ShowStartDate")
			pStartDateDesc = rsSub("SB_StartDateDesc")
			pShowReoccurenceDate = rsSub("SB_ShowReoccurenceDate")
			pReoccurenceDesc = rsSub("SB_ReoccurenceDesc")
			pShowEOSDate = rsSub("SB_ShowEOSDate")
			pEOSDesc = rsSub("SB_EOSDesc")
			pShowTrialDate = rsSub("SB_ShowTrialDate")
			pFreeTrialDesc = rsSub("SB_FreeTrialDesc")
			
			'// Installment (coming later)
			'pShowInstallment=rsSub("pcSubscription_ShowInstallment")
			'pInstallmentDesc=rsSub("pcSubscription_InstallmentDesc")

		Else
		
			'// Link ID
			pcv_strLinkID = rsSub("SB_LinkID") '// The Link ID from SubscriptionBridge.com
			
		  	'// Plan - Misc.
			pcv_intIsTrial = rsSub("SB_IsTrial") '// Is Trial - True OR False			  
		  	'pSubStartImmed = 0 '// rsSub("pcSubscription_StartImmed")   '// Start Immediately (TO DO)
			
		  	'// Package
			'pcv_curAmount = 0  '// Price (comes from cart)
			pcv_curTrialAmount = rsSub("SB_TrialAmount") '// Trial Price

			pSubAgree = "1"

			'// Display Setings
			pShowTrialPrice = rsSub("SB_ShowTrialPrice")
			pTrialDesc = rsSub("SB_TrialDesc")
			pShowFreeTrial = rsSub("SB_ShowFreeTrial")
			pShowStartDate = rsSub("SB_ShowStartDate")
			pStartDateDesc = rsSub("SB_StartDateDesc")
			pShowReoccurenceDate = rsSub("SB_ShowReoccurenceDate")
			pReoccurenceDesc = rsSub("SB_ReoccurenceDesc")
			pShowEOSDate = rsSub("SB_ShowEOSDate")
			pEOSDesc = rsSub("SB_EOSDesc")
			pShowTrialDate = rsSub("SB_ShowTrialDate")
			pFreeTrialDesc = rsSub("SB_FreeTrialDesc")

		End If
		
		'// 0 = Infinite
		'// 1 = Fixed 1-999
		'// 2 = Installment
		pSubType = rsSub("SB_Type")
		if isnull(pSubType) or pSubType="" Then 
			pSubType = "0"
		end if 
		
	Else

		pcv_intBillingFrequency = 0
		pcv_strBillingPeriod = ""
		pcv_intBillingCycles = 0 
		pSubStartImmed = 0
		pSubStartFRomPurch = 0 
		pSubStartDate = ""
		pcv_intIsTrial = 0
		pcv_intTrialCycles = 0 
		pcv_curTrialAmount = 0		
		pSubType = 0 

	end if


	If pcv_intIsTrial = "0" Then	
		 pcv_intTrialCycles = "0"
		 pcv_curTrialAmount = "0"	
	End if 


	if pSubType = 2 Then
		pSubInstall = 1
	End if


	'// Global Preferences
	If pShowTrialPrice = 0 Then
		pShowTrialPrice = scSBShowTrialPrice  '// "1" or "0"
	End If
	If pTrialDesc = "" OR isNULL(pTrialDesc) Then
		pTrialDesc = scSBTrialDesc  '// "1" or "0"
	End If
	If pShowFreeTrial = 0 Then
		pShowFreeTrial = scSBShowFreeTrial  '// "1" or "0"
	End If
	If pShowStartDate = 0 Then
		pShowStartDate = scSBShowStartDate
	End If
	If pStartDateDesc = "" OR isNULL(pStartDateDesc) Then
		pStartDateDesc = scSBStartDateDesc
	End If
	If pShowReoccurenceDate = 0 Then
		pShowReoccurenceDate = scSBShowReoccurenceDate
	End If
	If pReoccurenceDesc = "" OR isNULL(pReoccurenceDesc) Then
		pReoccurenceDesc = scSBReoccurenceDesc
	End If
	If pShowEOSDate = 0 Then
		pShowEOSDate = scSBShowEOSDate
	End If
	If pEOSDesc = "" OR isNULL(pEOSDesc) Then
		pEOSDesc = scSBEOSDesc
	End If
	If pShowTrialDate = 0 Then
		pShowTrialDate = scSBShowTrialDate
	End If
	If pFreeTrialDesc = "" OR isNULL(pFreeTrialDesc) Then
		pFreeTrialDesc = scSBFreeTrialDesc
	End If

	'// Only do math if not linked
	If pIsLinked="0" Then
	
		'// Start Date
		'// 0 = Immediate
		'// 1 = Use time above?
		'// 2 = Use specific date
		if pSubStartImmed = "" or pSubStartImmed= "0" Then		
			pSubStartDate = getSubDate("", scDateFrmt, "", date(), "format")
		ElseIf pSubStartImmed ="1" Then				   
			pSubStartDate = getSubDate(pcv_strBillingPeriod, scDateFrmt, pSubStartFromPurch, date(), "add")			    
		ElseIf pSubStartImmed ="2" Then	
			pSubStartDate = getSubDate("", scDateFrmt, "", pSubStart, "format")
		End if 	
		
		'// Reoccurance Date
		if pSubType <> 2 Then
			'// Not Installment
			pSubReOccur = getSubDate(pcv_strBillingPeriod, scDateFrmt, pcv_intBillingFrequency, pSubStartDate, "add")
			pcv_intBillingCyclesUntDate = getSubDate(pcv_strBillingPeriod, scDateFrmt, (pcv_intBillingCycles*pcv_intBillingFrequency), pSubStartDate, "add")		
		else
			'// Installment (coming later)
			'pSubStartDate = getSubDate(pcv_strBillingPeriod,scDateFrmt,(pcv_intTrialCycles*pcv_intBillingFrequency),pSubStartDate,"add")
			'pSubReOccur = pSubStartDate 'getSubDate(pcv_strBillingPeriod,scDateFrmt,pcv_intBillingFrequency,pSubStartDate,"add")
			'pcv_intBillingCyclesUntDate = getSubDate(pcv_strBillingPeriod,scDateFrmt,((pcv_intBillingCycles*pcv_intBillingFrequency)-1),pSubStartDate,"add")	
		end if 
		
		'// Trial End Date
		pcv_intTrialCyclesUntDate = getSubDate(pcv_strBillingPeriod, scDateFrmt, (pcv_intTrialCycles*pcv_intBillingFrequency), pSubStartDate, "add")
	
	Else
		pSubStartDate = date()	
	End If
	
	set rs=nothing
 
 End if  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END - GET PACKAGE DATA
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'SB E
%>