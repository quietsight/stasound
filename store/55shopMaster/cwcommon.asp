<%
emailID=-1
passID=-1
ctypeID=-1
fnameID=-1
lnameID=-1
comID=-1
phoneID=-1
faxID=-1
addrID=-1
addr2ID=-1
cityID=-1
statcodeID=-1
statID=-1
zipID=-1
councodeID=-1
sComID=-1
sAddrID=-1
sAddr2ID=-1
sCityID=-1
sStatCodeID=-1
sStatID=-1
sZipID=-1
sCounCodeID=-1
RewID=-1
NewsID=-1
PriceCatID=-1
sEmailID=-1
sPhoneID=-1
'MailUp-S
OptInMUID=-1
OptOutMUID=-1
OptInMU=""
OptOutMU=""
'MailUp-E

CTemail=""
CTpass=""
CTctype=0
CTfname=""
CTlname=""
CTcom=""
CTphone=""
CTfax=""
CTaddr=""
CTaddr2=""
CTcity=""
CTstatcode=""
CTstat=""
CTzip=""
CTcouncode=""
CTsCom=""
CTsAddr=""
CTsAddr2=""
CTsCity=""
CTsStatCode=""
CTsStat=""
CTsZip=""
CTsCounCode=""
CTRew=0
CTNews=0
PriceCat=0
sEmail=""
sPhone=""

TempProducts=""
ErrorsReport=""%>

<!--#include file="cwcheckfields.asp"-->

<%

For i=1 to request("validfields")
	
	BLine=0
	Select Case request("T" & i)
	Case "E-mail Address": emailID=request("P" & i)
	BLine=1
	Case "Password": passID=request("P" & i)
	BLine=2
	Case "Customer Type": ctypeID=request("P" & i)
	BLine=3
	Case "First Name": fnameID=request("P" & i)
	BLine=4
	Case "Last Name": lnameID=request("P" & i)	
	BLine=5
	Case "Company": comID=request("P" & i)
	BLine=6
	Case "Phone": phoneID=request("P" & i)
	BLine=7
	Case "Address": addrID=request("P" & i)
	BLine=8
	Case "Address 2": addr2ID=request("P" & i)
	BLine=9
	Case "City": cityID=request("P" & i)
	BLine=10
	Case "State Code (US/Canada)": statcodeID=request("P" & i)
	BLine=11
	Case "Province": statID=request("P" & i)
	BLine=12
	Case "Postal Code": zipID=request("P" & i)
	BLine=13
	Case "Country Code": councodeID=request("P" & i)
	BLine=14
	Case "Shipping Company": sComID=request("P" & i)
	BLine=15
	Case "Shipping Address": sAddrID=request("P" & i)
	BLine=16
	Case "Shipping Address 2": sAddr2ID=request("P" & i)			
	BLine=17
	Case "Shipping City": sCityID=request("P" & i)
	BLine=18
	Case "Shipping State Code (US/Canada)": sStatCodeID=request("P" & i)
	BLine=19
	Case "Shipping Province": sStatID=request("P" & i)
	BLine=20
	Case "Shipping Postal Code": sZipID=request("P" & i)		
	BLine=21
	Case "Shipping Country Code": sCounCodeID=request("P" & i)
	BLine=22
	Case "Current Reward Points Balance": RewID=request("P" & i)
	BLine=23
	Case "Newsletter Subscription": NewsID=request("P" & i)
	BLine=26
	Case "Pricing Category ID": PriceCatID=request("P" & i)
	BLine=27
	Case "Fax": faxID=request("P" & i)
	BLine=28
	Case "Shipping Email Address" : sEmailID=request("P" & i)
	BLine=29
	Case "Shipping Phone" : sPhoneID=request("P" & i)		
	BLine=30
	'MailUp-S
	Case "Opt-in MailUp List IDs": OptInMUID=request("P" & i)
	BLine=31
	Case "Opt-out MailUp List IDs": OptOutMUID=request("P" & i)
	BLine=32
	'MailUp-E
	Case Else:
		'Start Special Customer Fields
		if not IsNull(session("cp_cw_custfields")) then
			pcArr=session("cp_cw_custfields")
			For k=0 to ubound(pcArr,2)
				if request("T" & i)=pcArr(1,k) then
					pcArr(2,k)=request("P" & i)
					session("cp_cw_HaveCustField")="1"
				end if
			Next
			if session("cp_cw_HaveCustField")="1" then
				session("cp_cw_custfields")=pcArr
			end if
		end if
		'End of Special Customer Fields
	End Select
	
	if BLine>0 then
	TempStr=request("F" & i) & "*****"
	if instr(ALines(BLine-1),TempStr)=0 then
	ALines(BLine-1)=ALines(BLine-1) & TempStr
	end if
	BLine=0
	end if
Next

	SavedFile = "importlogs/cwsave.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 2)
	For dd=lbound(ALines) to ubound(ALines)
	f.WriteLine ALines(dd)
	Next
	f.close
%>