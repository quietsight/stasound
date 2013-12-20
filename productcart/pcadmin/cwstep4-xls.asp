<% pageTitle = "Import Results" %>
<% section = "mngAcc" %>
<%PmAdmin=7%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="cwcommon.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
on error resume next
Server.ScriptTimeout = 5400

dim f, query, conntemp, rstemp, rstemp1,TopRecord(100), IDcustom(2), Customcontent(2)

call openDb()
	Append=session("append")
	if PPD="1" then
		FileXLS = "/"&scPcFolder&"/pc/catalog/" & session("importfile")
	else
		FileXLS = "../pc/catalog/" & session("importfile")
	end if
	Set cnnExcel = Server.CreateObject("ADODB.Connection")
	cnnExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(FileXLS) & ";Extended Properties=Excel 8.0;"
	Set rsExcel = Server.CreateObject("ADODB.Recordset")
	rsExcel.open "SELECT * FROM IMPORT;", cnnExcel 
	
	if Err.number<>0 then
		session("importfilename")=""
		response.redirect "msg.asp?message=30"
	end if
	TotalXLSlines=0
	ImportedRecords=0
	fields=session("totalfields")
	iCols = rsExcel.Fields.Count
		
	if rsExcel.EOF then
		session("importfilename")=""
		response.redirect "msg.asp?message=32"
	end if
	
	'Get previous information before import/update customers
	query="Select IDCustomer from customers order by IDCustomer desc"
	set rstemp4=connTemp.execute(query)
	
	if not rstemp4.eof then
	PreIDCustomer="" & rstemp4("IDCustomer")
	else
	PreIDCustomer="0"
	end if
		
	if session("append")="1" then
	UpdateType="UPDATE"
	else
	UpdateType="IMPORT"
	end if
	PreRecords=""
		
	MailError=0
	
	'MailUp-S
	if session("append")<>"1" then
		query="SELECT idcustomer FROM customers ORDER by idcustomer DESC;"
		set rsQ=connTemp.execute(query)
		tmp_StartNum=0
		if not rsQ.eof then
			tmp_StartNum=rsQ("idcustomer")
		end if
		set rsQ=nothing
	end if
	'MailUp-E
	
	Do While not rsExcel.EOF
	
	RecordError=false
	TotalXLSlines=TotalXLSlines+1
		
	 	
if RecordError=False then%>
<!--#include file="cwcommon2.asp"-->
<%end if%>
<%
if RecordError=false then 'STEP 1

	PriceCType=-1
	if clng(PriceCat)>0 then
		query="SELECT pcCC_WholesalePriv FROM pcCustomerCategories WHERE idCustomerCategory=" & PriceCat
		set rstemp4=connTemp.execute(query)
		if not rstemp4.eof then
			PriceCType=clng(rstemp4("pcCC_WholesalePriv"))
			CTctype=PriceCType
		end if
		set rstemp4=nothing
	end if

		query="Select * from Customers where email='" & CTemail & "'"
		set rstemp4=connTemp.execute(query)
		testMail=0
		if not rstemp4.eof then
			testMail=1
			IDCustomer=rstemp4("idCustomer")
		end if	

		
	pAppend=0
	if session("append")="1" then
		query="Select * from Customers where email='" & CTemail & "'"
		set rstemp4=connTemp.execute(query)
		IF not rstemp4.eof then 'EXISTING E-MAIL
			
			temp4=""
			
if passID>-1 then
temp4=temp4 & ",Password='" & CTpass & "'"
end if

if (ctypeID>-1) or (PriceCType<>-1) then
temp4=temp4 & ",customerType=" & CTctype
end if

if fnameID>-1 then
temp4=temp4 & ",Name='" & CTfname & "'"
end if

if lnameID>-1 then
temp4=temp4 & ",lastName='" & CTlname & "'"
end if

if comID>-1 then
temp4=temp4 & ",customerCompany='" & CTcom & "'"
end if

if phoneID>-1 then
temp4=temp4 & ",phone='" & CTphone & "'"
end if

if faxID>-1 then
temp4=temp4 & ",fax='" & CTfax & "'"
end if

if addrID>-1 then
temp4=temp4 & ",address='" & CTaddr & "'"
end if

if addr2ID>-1 then
temp4=temp4 & ",address2='" & CTaddr2 & "'"
end if

if cityID>-1 then
temp4=temp4 & ",city='" & CTcity & "'"
end if

if statcodeID>-1 then
temp4=temp4 & ",stateCode='" & CTstatcode & "'"
end if

if statID>-1 then
temp4=temp4 & ",state='" & CTstat & "'"
end if

if zipID>-1 then
temp4=temp4 & ",zip='" & CTzip & "'"
end if

if councodeID>-1 then
temp4=temp4 & ",countryCode='" & CTcouncode & "'"
end if

if sComID>-1 then
temp4=temp4 & ",shippingCompany='" & CTsCom & "'"
end if

if sAddrID>-1 then
temp4=temp4 & ",shippingaddress='" & CTsAddr & "'"
end if

if sAddr2ID>-1 then
temp4=temp4 & ",shippingAddress2='" & CTsAddr2 & "'"
end if

if sCityID>-1 then
temp4=temp4 & ",shippingcity='" & CTsCity & "'"
end if

if sStatCodeID>-1 then
temp4=temp4 & ",shippingStateCode='" & CTsStatCode & "'"
end if

if sStatID>-1 then
temp4=temp4 & ",shippingState='" & CTsStat & "'"
end if

if sZipID>-1 then
temp4=temp4 & ",shippingZip='" & CTsZip & "'"
end if

if sCounCodeID>-1 then
temp4=temp4 & ",shippingCountryCode='" & CTsCounCode & "'"
end if

if RewID>-1 then
temp4=temp4 & ",iRewardPointsAccrued=" & CTRew
end if

if NewsID>-1 then
temp4=temp4 & ",RecvNews=" & CTNews
end if

if PriceCatID>-1 then
temp4=temp4 & ",idCustomerCategory=" & PriceCat
end if
			
if sEmailID>-1 then
temp4=temp4 & ",ShippingEmail='" & sEmail &"'"
end if
if sPhoneID>-1 then
temp4=temp4 & ",ShippingPhone='" & sPhone &"'"
end if

			
			'Get customer information before update
			query="select * from Customers where email='" & CTemail & "'"
			set rstemp=conntemp.execute(query)
			
			IF not rstemp.eof THEN

			PreRecord1=""
			PreRecord1=PreRecord1 & rstemp("idCustomer") & "****"
			
			iCols = rstemp.Fields.Count
		    for dd=1 to iCols-1
		    FType="" & Rstemp.Fields.Item(dd).Type
		    if (Rstemp.Fields.Item(dd).Name="dtRewardsStarted") then
		    FType="DYDL"
		    end if
		    if (Ftype="202") or (Ftype="203") or (FType="DYDL") then
		    PTemp=Rstemp.Fields.Item(dd).Value
		    if PTemp<>"" then
		    PTemp=replace(PTemp,"'","''")
		    PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
		    end if
		    if FType="DYDL" then
		    if scDB="Access" then
		    myStr11="#"
		    else
		    myStr11="'"
		    end if
		    else
		    myStr11="'"
		    end if
		    if dd=1 then
		    PreRecord1=PreRecord1 & myStr11 & PTemp & myStr11
		    else
		    PreRecord1=PreRecord1 & "@@@@@" & myStr11 & PTemp & myStr11
		    end if
		    else
		    PTemp="" & Rstemp.Fields.Item(dd).Value
		    if PTemp<>"" then
		    else
		    PTemp="0"
		    end if
		    if dd=1 then
		    PreRecord1=PreRecord1 & PTemp
		    else
		    PreRecord1=PreRecord1 & "@@@@@" & PTemp
		    end if
		    end if
			next
			PreRecords=PreRecords & PreRecord1 & vbcrlf
			END IF

			query="update customers set email='" & CTemail & "'" & temp4 & " where email='" & CTemail & "'"
			query=replace(query,chr(34),"&quot;")
			query=replace(query,"**DD**",chr(34))
			set rstemp=conntemp.execute(query)	
			pAppend=1
			
			query="SELECT idcustomer FROM Customers WHERE email='" & CTemail & "'"
			set rstemp=conntemp.execute(query)
 
			pIdCustomer = rstemp("idCustomer")
			
			call updCustEditedDate(pIdCustomer)
				
		ELSE 'Do not have existing EMAIL
			MailError=1
			ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": This E-mail Address: " & CTemail & " is not in the database." & vbcrlf
			RecordError=true
		END IF
	else 'Append=0
	
	if testMail=1 then
		ErrorsReport=ErrorsReport & "Record " & TotalCSVlines & ": E-mail Address: " & CTemail & " could not be imported because it already exists." & vbcrlf
		RecordError=true
	else
		query="INSERT INTO customers (email,Password,customerType,Name,lastName,customerCompany,phone,fax,address,address2,city,stateCode,state,zip,countryCode,shippingCompany,shippingaddress,shippingAddress2,shippingcity,shippingStateCode,shippingState,shippingZip,shippingCountryCode,iRewardPointsAccrued,RecvNews,idCustomerCategory, ShippingEmail, ShippingPhone ) VALUES ("
		query=query & "'" & CTemail & "','" & CTpass & "'," & CTctype & ",'" & CTfname & "','" & CTlname & "','" & CTcom & "','" & CTphone & "','" & CTfax & "','" & CTaddr & "','" & CTaddr2 & "','" & CTcity & "','" & CTstatcode & "','" & CTstat & "','" & CTzip & "','" & CTcouncode & "','" & CTsCom & "','" & CTsAddr & "','" & CTsAddr2 & "','" & CTsCity & "','" & CTsStatCode & "','" & CTsStat & "','" & CTsZip & "','" & CTsCounCode & "'," & CTRew & "," & CTNews & "," & PriceCat &",'"&sEmail&"','"& sPhone& "')"
		query=replace(query,chr(34),"&quot;")
		query=replace(query,"**DD**",chr(34))
		set rstemp=conntemp.execute(query)
		
		query="SELECT idcustomer FROM Customers WHERE email='" & CTemail & "'"
		set rstemp=conntemp.execute(query)
 
		pIdCustomer = rstemp("idCustomer")
		
		call updCustCreatedDate(pIdCustomer)
		
	end if
		
	end if 'Update/Import
	
	IF RecordError=false THEN
	
	query="SELECT idcustomer FROM Customers WHERE email='" & CTemail & "'"
	set rstemp=conntemp.execute(query)
 
	pIdCustomer = rstemp("idCustomer")
	
	'MailUp-S
	'Opt-in
	tmp_HaveOptIn=0
	IF OptInMUID<>-1 AND trim(OptInMU)<>"" THEN
		tmpArr=split(OptInMU,"|")
		For k=lbound(tmpArr) to ubound(tmpArr)
			if trim(tmpArr(k))<>"" then
				query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmpArr(k) & ";"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					tmpMUList=rsQ("pcMailUpLists_ID")
					set rsQ=nothing
					query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idcustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
					set rsQ=connTemp.execute(query)
					dtTodaysDate=Date()
					if SQL_Format="1" then
						dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
					else
						dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
					end if
					if not rsQ.eof then
						if scDB="SQL" then
							query="UPDATE pcMailUpSubs SET idCustomer=" & pIdCustomer & ",pcMailUpLists_ID=" & tmpMUList & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=0 WHERE idCustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
						else
							query="UPDATE pcMailUpSubs SET idCustomer=" & pIdCustomer & ",pcMailUpLists_ID=" & tmpMUList & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=0 WHERE idCustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
						end if
					else
						if scDB="SQL" then
							query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pIdCustomer & "," & tmpMUList & ",'" & dtTodaysDate & "',1,0);"
						else
							query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pIdCustomer & "," & tmpMUList & ",#" & dtTodaysDate & "#,1,0);"
						end if
					end if
					set rsQ=nothing
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
					tmp_HaveOptIn=1
				end if
				set rsQ=nothing
			end if
		Next
	
	query="UPDATE Customers SET RecvNews=" & tmp_HaveOptIn & " WHERE idcustomer=" & pIdCustomer & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	
	END IF
	
	'Opt-out
	IF OptOutMUID<>-1 AND trim(OptOutMU)<>"" THEN
		tmpArr=split(OptOutMU,"|")
		For k=lbound(tmpArr) to ubound(tmpArr)
			if trim(tmpArr(k))<>"" then
				query="SELECT pcMailUpLists_ID FROM pcMailUpLists WHERE pcMailUpLists_ListID=" & tmpArr(k) & ";"
				set rsQ=connTemp.execute(query)
				if not rsQ.eof then
					tmpMUList=rsQ("pcMailUpLists_ID")
					set rsQ=nothing
					query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idcustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
					set rsQ=connTemp.execute(query)
					dtTodaysDate=Date()
					if SQL_Format="1" then
						dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
					else
						dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
					end if
					if not rsQ.eof then
						if scDB="SQL" then
							query="UPDATE pcMailUpSubs SET idCustomer=" & pIdCustomer & ",pcMailUpLists_ID=" & tmpMUList & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=1 WHERE idCustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
						else
							query="UPDATE pcMailUpSubs SET idCustomer=" & pIdCustomer & ",pcMailUpLists_ID=" & tmpMUList & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=1,pcMailUpSubs_Optout=1 WHERE idCustomer=" & pIdCustomer & " AND pcMailUpLists_ID=" & tmpMUList & ";"
						end if
					else
						if scDB="SQL" then
							query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pIdCustomer & "," & tmpMUList & ",'" & dtTodaysDate & "',1,1);"
						else
							query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & pIdCustomer & "," & tmpMUList & ",#" & dtTodaysDate & "#,1,1);"
						end if
					end if
					set rsQ=nothing
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				end if
				set rsQ=nothing
			end if
		Next
	END IF
	
	'MailUp-E
	
	END IF	
	
	'Start Special Customer Fields
	if RecordError=false then
		if session("cp_cw_HaveCustField")="1" then
			pcArr=session("cp_cw_custfields")
			For k=0 to ubound(pcArr,2)
				if cint(pcArr(2,k))>-1 then
					query="SELECT pcCField_ID FROM pcCustomerFieldsValues WHERE idcustomer=" & pIdCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						query="UPDATE pcCustomerFieldsValues SET pcCFV_Value='" & pcArr(3,k) & "' WHERE idcustomer=" & pIdCustomer & " AND pcCField_ID=" & pcArr(0,k) & ";"
					else
						query="INSERT INTO pcCustomerFieldsValues (idcustomer,pcCField_ID,pcCFV_Value) VALUES (" & pIdCustomer & "," & pcArr(0,k) & ",'" & pcArr(3,k) & "');"
					end if
					set rsQ=nothing
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				end if
			Next
		end if
	end if
	'End of Special Customer Fields
	
	if RecordError=false then
	ImportedRecords=ImportedRecords+1
	end if
	
end if 'END STEP 1
	
	rsExcel.MoveNext
	
	Loop
	
	set rsexcel=nothing
	cnnExcel.close
	set cnnExcel=nothing

	'Delete Import File
	'Set fso = server.CreateObject("Scripting.FileSystemObject")
	'Set f = fso.GetFile(Server.MapPath(FileXLS))
	'f.Delete
	'Set fso = nothing
	'Set f = nothing
	
	'MailUp-S
	if session("append")<>"1" then
	query="SELECT idcustomer FROM customers ORDER by idcustomer DESC;"
	set rsQ=connTemp.execute(query)
	tmp_EndNum=0
	if not rsQ.eof then
		tmp_EndNum=rsQ("idcustomer")
	end if
	set rsQ=nothing
	end if
	
	if clng(tmp_StartNum)<>clng(tmp_EndNum) then
		query="UPDATE pcMailUpSettings SET pcMailUpSett_LastCustomerID='" & tmp_StartNum & "," & tmp_EndNum & "';"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if
	'MailUp-E
	
	call closeDB()
	
	if ImportedRecords>0 then
	
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set afi=fs.CreateTextFile(server.MapPath(".") & "\importlogs\custlogs.txt",True)
		
	afi.Writeline(UpdateType)
	afi.Writeline(PreIDCustomer)
	afi.Writeline(PreRecords)
	afi.Close
	
	end if
	
	session("importfile")=""
	session("totalfields")=0
	
	if MailError=1 then
	ErrorsReport="One of the records you are importing does not currently exist in the database. The Update feature is strictly for modifying existing customer information. Please correct the error and try again." &vbcrlf&vbcrlf &ErrorsReport
	end if

%>

<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
	    <table class="pcCPcontent">
	    <tr>
	        <td colspan="2"><h2>Steps:</h2></td>
	    </tr>
	    <tr>
	        <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
	        <td width="95%"><font color="#A8A8A8">Select product data file</font></td>
	    </tr>
	    <tr>
	        <td align="right"><img border="0" src="images/step2.gif"></td>
	        <td><font color="#A8A8A8">Map fields</font></td>
	    </tr>
	    <tr>
	        <td align="right"><img border="0" src="images/step3.gif"></td>
	        <td><font color="#A8A8A8">Confirm mapping</font></td>
	    </tr>
	    <tr>
	        <td align="right"><img border="0" src="images/step4a.gif"></td>
	        <td><strong><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</strong></td>
	    </tr>
	    </table>
		<br>

		Total <b><%=ImportedRecords%></b> records were <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
		<%if TotalXLSlines-ImportedRecords>0 then%>
			<br>
			Total <b><font color="#FF0000"><%=TotalXLSlines-ImportedRecords%></font></b> records could not be <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!
		<%end if%>
		<br>
		<br>
		<%if ErrorsReport<>"" then%> 
			<table class="pcCPcontent">
				<tr>
				<th>Error(s) Report</th>
				</tr>
				<tr>
					<td>
						<textarea rows="7" name="S1" cols="62" style="font-family: Arial; font-size: 10pt"><%=ErrorsReport%></textarea>
					</td>
				</tr>
			</table>
		<%end if%>
		<p align="center">
        <br>
        <br>
			<input type="button" name="mainmenu" value="Back to Main menu" onClick="location='menu.asp';" class="ibtnGrey">
		</p>
	</td>
</tr>
</table>
<%
session("append")=0
'Start Special Customer Fields
session("cp_cw_custfields")=""
session("cp_cw_HaveCustField")=""
'End of Special Customer Fields
%>
<!--#include file="AdminFooter.asp"-->