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
CTemail=trim(rsExcel.Fields.Item(int(emailID)).Value)
end if

if passID<>-1 then
CTpass=trim(rsExcel.Fields.Item(int(passID)).Value)
else
CTpass=""
end if

if ctypeID<>-1 then
CTctype=trim(rsExcel.Fields.Item(int(ctypeID)).Value)
end if

'Start Special Customer Fields
	if session("cp_cw_HaveCustField")="1" then
		pcArr=session("cp_cw_custfields")
		For k=0 to ubound(pcArr,2)
			pcArr(3,k)=""
			if cint(pcArr(2,k))>-1 then
				pcArr(3,k)=trim(rsExcel.Fields.Item(int(pcArr(2,k))).Value)
			end if
		Next
		session("cp_cw_custfields")=pcArr
	end if
'End of Special Customer Fields

if session("append")<>"1" then
		if CTpass<>"" then
		else
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
CTfname=trim(rsExcel.Fields.Item(int(fnameID)).Value)
if CTfname<>"" then
else
CTfname="NA during import"
end if
end if

if lnameID<>-1 then
CTlname=trim(rsExcel.Fields.Item(int(lnameID)).Value)
if CTlname<>"" then
else
CTlname="NA during import"
end if
end if

if comID<>-1 then
CTcom=trim(rsExcel.Fields.Item(int(comID)).Value)
end if

if phoneID<>-1 then
CTphone=trim(rsExcel.Fields.Item(int(phoneID)).Value)
end if

if faxID<>-1 then
CTfax=trim(rsExcel.Fields.Item(int(faxID)).Value)
end if

if addrID<>-1 then
CTaddr=trim(rsExcel.Fields.Item(int(addrID)).Value)
end if

if addr2ID<>-1 then
CTaddr2=trim(rsExcel.Fields.Item(int(addr2ID)).Value)
end if

if cityID<>-1 then
CTcity=trim(rsExcel.Fields.Item(int(cityID)).Value)
end if

if statcodeID<>-1 then
CTstatcode=trim(rsExcel.Fields.Item(int(statcodeID)).Value)
end if

if statID<>-1 then
CTstat=trim(rsExcel.Fields.Item(int(statID)).Value)
end if

if zipID<>-1 then
CTzip=trim(rsExcel.Fields.Item(int(zipID)).Value)
end if

if councodeID<>-1 then
CTcouncode=trim(rsExcel.Fields.Item(int(councodeID)).Value)
end if

if sComID<>-1 then
CTsCom=trim(rsExcel.Fields.Item(int(sComID)).Value)
end if

if sAddrID<>-1 then
CTsAddr=trim(rsExcel.Fields.Item(int(sAddrID)).Value)
end if

if sAddr2ID<>-1 then
CTsAddr2=trim(rsExcel.Fields.Item(int(sAddr2ID)).Value)
end if

if sCityID<>-1 then
CTsCity=trim(rsExcel.Fields.Item(int(sCityID)).Value)
end if

if sStatCodeID<>-1 then
CTsStatCode=trim(rsExcel.Fields.Item(int(sStatCodeID)).Value)
end if

if sStatID<>-1 then
CTsStat=trim(rsExcel.Fields.Item(int(sStatID)).Value)
end if

if sZipID<>-1 then
CTsZip=trim(rsExcel.Fields.Item(int(sZipID)).Value)
end if

if sCounCodeID<>-1 then
CTsCounCode=trim(rsExcel.Fields.Item(int(sCounCodeID)).Value)
end if

if RewID<>-1 then
CTRew=trim(rsExcel.Fields.Item(int(RewID)).Value)
end if

if NewsID<>-1 then
CTNews=trim(rsExcel.Fields.Item(int(NewsID)).Value)
end if

'MailUp-S
if OptInMUID<>-1 then
OptInMU=trim(rsExcel.Fields.Item(int(OptInMUID)).Value)
end if

if OptOutMUID<>-1 then
OptOutMU=trim(rsExcel.Fields.Item(int(OptOutMUID)).Value)
end if
'MailUp-E

if PriceCatID<>-1 then
PriceCat=trim(rsExcel.Fields.Item(int(PriceCatID)).Value)
end if
if PriceCat="" then
	PriceCat="0"
end if
if IsNumeric(PriceCat)=false then
	PriceCat="0"
end if

			
if sEmailID>-1 then
sEmail=trim(rsExcel.Fields.Item(int(sEmailID)).Value)
end if

if sPhoneID>-1 then
sPhone=trim(rsExcel.Fields.Item(int(sPhoneID)).Value)
end if


		if CTRew<>"" then
		if IsNumeric(CTRew)=false then
		CTRew="0"
		end if
		else
		CTRew="0"
		end if
		if CTNews<>"" then
		if IsNumeric(CTNews)=false then
		CTNews="0"
		end if
		if CTNews="-1" then
		CTNews="1"
		end if
		else
		CTNews="0"
		end if
			
		
		if CTemail<>"" then
		else
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not include an E-mail Address." & vbcrlf
		RecordError=true
		end if
		if session("append")<>"1" then
		if CTpass<>"" then
		else
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not include a Password." & vbcrlf
		RecordError=true
		end if
		if CTctype<>"" then
		else
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not include a Customer Type." & vbcrlf
		RecordError=true
		end if
		if CTfname<>"" then
		else
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not include a First Name." & vbcrlf
		RecordError=true
		end if
		if CTlname<>"" then
		else
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": does not have a Last Name." & vbcrlf
		RecordError=true
		end if				
		end if
		if isNumeric(CTRew)=false then
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": The Current Reward Points Balance is not a number." & vbcrlf
		RecordError=true
		end if
		if isNumeric(CTNews)=false then
		ErrorsReport=ErrorsReport & "Record " & TotalXLSlines & ": The Newsletter Subscription field value is not a number." & vbcrlf
		RecordError=true
		end if
		
		if scDecSign="," then
			CTRew=replace(CTRew,".","")
		else
			CTRew=replace(CTRew,",","")
		end if
		
		CTRew=replace(CTRew,scCurSign,"")
		
		if CTfname<>"" then
			CTfname=replace(CTfname,"'","''")
		end if
		if CTlname<>"" then
			CTlname=replace(CTlname,"'","''")
		end if
		if CTcom<>"" then
			CTcom=replace(CTcom,"'","''")
		end if
		if CTaddr<>"" then
			CTaddr=replace(CTaddr,"'","''")
		end if
		if CTaddr2<>"" then
			CTaddr2=replace(CTaddr2,"'","''")
		end if
		if CTcity<>"" then
			CTcity=replace(CTcity,"'","''")
		end if
		if CTstatcode<>"" then
			CTstatcode=replace(CTstatcode,"'","''")
		end if
		if CTstat<>"" then
			CTstat=replace(CTstat,"'","''")
		end if
		if CTcouncode<>"" then
			CTcouncode=replace(CTcouncode,"'","''")
		end if
		if CTsCom<>"" then
			CTsCom=replace(CTsCom,"'","''")
		end if
		if CTsAddr<>"" then
			CTsAddr=replace(CTsAddr,"'","''")
		end if
		if CTsAddr2<>"" then
			CTsAddr2=replace(CTsAddr2,"'","''")
		end if
		if CTsCity<>"" then
			CTsCity=replace(CTsCity,"'","''")
		end if
		if CTsStatCode<>"" then
			CTsStatCode=replace(CTsStatCode,"'","''")
		end if
		if CTsStat<>"" then
			CTsStat=replace(CTsStat,"'","''")
		end if
		if CTsCounCode<>"" then
			CTsCounCode=replace(CTsCounCode,"'","''")
		end if
		
		if CTpass<>"" then
			CTpass=enDeCrypt(CTpass, scCrypPass)
			CTpass=replace(CTpass,chr(34),"**DD**")
		end if
		
%>