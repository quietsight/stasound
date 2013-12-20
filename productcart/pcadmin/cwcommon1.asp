<%

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
sEmail =""
sPhone=""
if emailID<>-1 then
CTemail=trim(CSVRecord(emailID))
end if

if passID<>-1 then
CTpass=trim(CSVRecord(passID))
else
CTpass=""
end if

if ctypeID<>-1 then
CTctype=trim(CSVRecord(ctypeID))
end if

'Start Special Customer Fields
	if session("cp_cw_HaveCustField")="1" then
		pcArr=session("cp_cw_custfields")
		For k=0 to ubound(pcArr,2)
			pcArr(3,k)=""
			if cint(pcArr(2,k))>-1 then
				pcArr(3,k)=trim(CSVRecord(cint(pcArr(2,k))))
			end if
		Next
		session("cp_cw_custfields")=pcArr
	end if
'End of Special Customer Fields

if session("append")<>"1" then
		if CTpass="" then
			Dim z
			For z=1 to 16
	 		randomize
	 		CTpass=CTpass & Cstr(Fix(rnd*10))
	 		Next
		end if
		if CTctype<>"" then
			if instr(ucase(CTctype),"WHOLESALE")>0 then
				CTctype="1"
			else
				if (CTctype<>"1") and (CTctype<>"-1") then
					CTctype="0"
				else
					CTctype="1"
				end if	
			end if
		else
			CTctype="0"
		end if
end if

if fnameID<>-1 then
CTfname=trim(CSVRecord(fnameID))
if CTfname<>"" then
else
CTfname="NA during import"
end if
end if

if lnameID<>-1 then
CTlname=trim(CSVRecord(lnameID))
if CTlname<>"" then
else
CTlname="NA during import"
end if
end if

if comID<>-1 then
CTcom=trim(CSVRecord(comID))
end if

if phoneID<>-1 then
CTphone=trim(CSVRecord(phoneID))
end if

if faxID<>-1 then
CTfax=trim(CSVRecord(faxID))
end if

if addrID<>-1 then
CTaddr=trim(CSVRecord(addrID))
end if

if addr2ID<>-1 then
CTaddr2=trim(CSVRecord(addr2ID))
end if

if cityID<>-1 then
CTcity=trim(CSVRecord(cityID))
end if

if statcodeID<>-1 then
CTstatcode=trim(CSVRecord(statcodeID))
end if

if statID<>-1 then
CTstat=trim(CSVRecord(statID))
end if

if zipID<>-1 then
CTzip=trim(CSVRecord(zipID))
end if

if councodeID<>-1 then
CTcouncode=trim(CSVRecord(councodeID))
end if

if sComID<>-1 then
CTsCom=trim(CSVRecord(sComID))
end if

if sAddrID<>-1 then
CTsAddr=trim(CSVRecord(sAddrID))
end if

if sAddr2ID<>-1 then
CTsAddr2=trim(CSVRecord(sAddr2ID))
end if

if sCityID<>-1 then
CTsCity=trim(CSVRecord(sCityID))
end if

if sStatCodeID<>-1 then
CTsStatCode=trim(CSVRecord(sStatCodeID))
end if

if sStatID<>-1 then
CTsStat=trim(CSVRecord(sStatID))
end if

if sZipID<>-1 then
CTsZip=trim(CSVRecord(sZipID))
end if

if sCounCodeID<>-1 then
CTsCounCode=trim(CSVRecord(sCounCodeID))
end if

if RewID<>-1 then
CTRew=trim(CSVRecord(RewID))
end if

if NewsID<>-1 then
CTNews=trim(CSVRecord(NewsID))
end if

'MailUp-S
if OptInMUID<>-1 then
OptInMU=trim(CSVRecord(OptInMUID))
end if

if OptOutMUID<>-1 then
OptOutMU=trim(CSVRecord(OptOutMUID))
end if
'MailUp-E

if PriceCatID<>-1 then
PriceCat=trim(CSVRecord(PriceCatID))
end if
if PriceCat="" then
	PriceCat="0"
end if
if IsNumeric(PriceCat)=false then
	PriceCat="0"
end if

			
if sEmailID<>-1 then
sEmail=trim(CSVRecord(sEmailID))
end if

if sPhoneID<>-1 then
sEmail=trim(CSVRecord(sPhoneID))
end if

		if CTRew="" then
		CTRew="0"
		else
		if IsNumeric(CTRew)=false then
		CTRew="0"
		end if
		end if
		if CTNews="" then
		CTNews="0"
		else
		if IsNumeric(CTNews)=false then
		CTNews="0"
		end if
		end if
			
		
		if CTemail="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not include an E-mail Address." & vbcrlf
		RecordError=true
		end if
		if session("append")<>"1" then
		if CTpass="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not include a Password." & vbcrlf
		RecordError=true
		end if
		if CTctype="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not include a Customer Type." & vbcrlf
		RecordError=true
		end if
		if CTfname="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not include a First Name." & vbcrlf
		RecordError=true
		end if
		if CTlname="" then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": does not include a Last Name." & vbcrlf
		RecordError=true
		end if				
		end if
		if isNumeric(CTRew)=false then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Current Reward Points Balance is not a number." & vbcrlf
		RecordError=true
		end if
		if isNumeric(CTNews)=false then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": The Newsletter Subscription field value is not a number." & vbcrlf
		RecordError=true
		end if
		
		if scDecSign="," then
			CTRew=replace(CTRew,".","")
		else
			CTRew=replace(CTRew,",","")
		end if
		
		CTRew=replace(CTRew,scCurSign,"")
		
		if CTpass<>"" then
			CTpass=enDeCrypt(CTpass, scCrypPass)
			CTpass=replace(CTpass,chr(34),"**DD**")
		end if
%>