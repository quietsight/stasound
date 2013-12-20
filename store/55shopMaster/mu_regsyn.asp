<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=0%><!--#include file="adminv.asp"-->   
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/MailUpFunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<% pageTitle="MailUp - Bulk Registration/Synchronization" %>
<% section="mngAcc"
Server.ScriptTimeout = 5400
Dim connTemp,query,rs
Dim tmp_setup
Dim tmpCanNotStart
Dim pcMU_sendconfirm
pcMU_sendconfirm=""
tmpCanNotStart=0

	tmp_setup=0
	tmp_bulk=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	call opendb()
	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess,pcMailUpSett_BulkRegister FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("CP_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("CP_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("CP_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		tmp_bulk=rs("pcMailUpSett_BulkRegister")
		if IsNull(tmp_bulk) or tmp_bulk="" then
			tmp_bulk=0
		end if
	end if
	set rs=nothing
call closedb()

if tmp_setup=0 then
	response.redirect "mu_manageNewsWiz.asp"
end if

msg=""

'Post Back
IF request("action")="run" THEN
call opendb()
	'Initial Setup
	IF request("submit1")<>"" then
		tmpSuccess=1
		tmpSuccessCount=0
		tmpCanNotStart=0
		tmpIDProList=""
		tmpIDLists=""
		tmpXMLDoc=""
		conFirmEmail=request("confirm1")
		if conFirmEmail="" then
			conFirmEmail=0
		end if
		
		query="SELECT idcustomer,email,[name],lastName,customerCompany FROM Customers WHERE RecvNews=1 AND customerType<3 AND suspend=0;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			custArr=rs.getRows()
			custCount=ubound(custArr,2)
			set rs=nothing
			For k=0 to custCount
				if (custArr(2,k)<>"") OR (custArr(3,k)<>"") OR (custArr(4,k)<>"") then
					tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & custArr(1,k) & """ Prefix="""" Number="""">"
					if (custArr(2,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(custArr(2,k)) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (custArr(3,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(custArr(3,k)) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (custArr(4,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(custArr(4,k)) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
				else
					tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & custArr(1,k) & """ />"
				end if
			Next
			tmpXMLDoc="<subscribers>" & tmpXMLDoc & "</subscribers>"
			query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_Active=1 AND pcMailUpLists_Removed=0;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				listArr=rs.getRows()
				listCount=ubound(listArr,2)
				set rs=nothing
				tmpListID=""
				tmpListGuid=""
				tmpGroupID=""
				For k=0 to listCount
					if tmpListID<>"" then
						tmpListID=tmpListID & ";"
						tmpListGuid=tmpListGuid & ";"
						tmpGroupID=tmpGroupID & ";"
					end if
					tmpListID=tmpListID & listArr(1,k)
					tmpListGuid=tmpListGuid & listArr(2,k)
					tmpGroupID=tmpGroupID & "0"
				Next
				
				tmpMUResult=MUImport(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),tmpListID,tmpListGuid,tmpXMLDoc,tmpGroupID,0,0,conFirmEmail)
				
				if tmpMUResult=0 then
					tmpSuccess=0
				else
					tmpReturnIDs=split(session("CP_MU_ReturnIDs"),";")
					tmpReturnList=split(session("CP_MU_ReturnList"),";")
					tmpReturnCode=split(session("CP_MU_ReturnCode"),";")
					For k=0 to listCount
						tmpCorrect=0
						For m=0 to listCount
							if clng(listArr(1,k))=clng(tmpReturnList(m)) then
								tmpMUResult1=tmpReturnIDs(m)
								tmpCorrect=m
								exit for
							end if
						Next
							
						if clng(tmpMUResult1)=0 then
							tmpSuccess=0
							tmpCanNotStart=tmpCanNotStart+1
						end if
						tmpIDProList=tmpIDProList & tmpMUResult1 & "*" & tmpReturnCode(tmpCorrect) & "||"
						tmpIDLists=tmpIDLists & listArr(1,k) & "||"
						IF clng(tmpMUResult1)>0 THEN
						tmpSuccessCount=tmpSuccessCount+1
						For l=0 to custCount
							query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idcustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
							set rsQ=connTemp.execute(query)
							dtTodaysDate=Date()
							if SQL_Format="1" then
								dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
							else
								dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
							end if
							if not rsQ.eof then
								if scDB="SQL" then
									query="UPDATE pcMailUpSubs SET idCustomer=" & custArr(0,l) & ",pcMailUpLists_ID=" & listArr(0,k) & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
								else
									query="UPDATE pcMailUpSubs SET idCustomer=" & custArr(0,l) & ",pcMailUpLists_ID=" & listArr(0,k) & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
								end if
							else
								if scDB="SQL" then
									query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & custArr(0,l) & "," & listArr(0,k) & ",'" & dtTodaysDate & "',0,0);"
								else
									query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & custArr(0,l) & "," & listArr(0,k) & ",#" & dtTodaysDate & "#,0,0);"
								end if
							end if
							set rsQ=nothing
							set rsQ=connTemp.execute(query)
							set rsQ=nothing
						Next
						END IF
					Next
				end if
				if tmpSuccess=1 then
					msg="1A"
				else
					if tmpSuccess=0 AND tmpSuccessCount>0 then
						msg="1B"
					else
						msg="1C"
					end if
				end if
			end if 'Have Active Lists
			set rs=nothing
		end if ' Have Opt-in Customers
		set rs=nothing
		if tmpIDProList<>"" then
			query="UPDATE pcMailUpSettings SET pcMailUpSett_LastIDList='" & tmpIDLists & "',pcMailUpSett_LastIDProcess='" & tmpIDProList & "';"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
		if msg="1A" then
			query="UPDATE pcMailUpSettings SET pcMailUpSett_BulkRegister=1;"
			set rs=connTemp.execute(query)
			set rs=nothing
			tmp_bulk=1
		end if
	END IF
	
	'Imported Customers
	IF request("submit2")<>"" then
		tmpSuccess=1
		tmpSuccessCount=0
		tmpCanNotStart=0
		tmpIDProList=""
		tmpIDLists=""
		tmpXMLDoc=""
		conFirmEmail=request("confirm2")
		if conFirmEmail="" then
			conFirmEmail=0
		end if
		
		query="SELECT pcMailUpSett_LastCustomerID FROM pcMailUpSettings;"
		set rsQ=connTemp.execute(query)
		tmpCustomers=""
		if not rsQ.eof then
			tmpCustomers=rsQ("pcMailUpSett_LastCustomerID")
		end if
		set rsQ=nothing
		
		tmp1=split(tmpCustomers,",")
		
		query="SELECT idcustomer,email,[name],lastName,customerCompany FROM Customers WHERE customerType<3 AND suspend=0 AND idcustomer>" & tmp1(0) & " AND idcustomer<=" & tmp1(1) & ";"
		set rs=connTemp.execute(query)
		if not rs.eof then
			custArr=rs.getRows()
			custCount=ubound(custArr,2)
			set rs=nothing
			For k=0 to custCount
				if (custArr(2,k)<>"") OR (custArr(3,k)<>"") OR (custArr(4,k)<>"") then
					tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & custArr(1,k) & """ Prefix="""" Number="""">"
					if (custArr(2,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo1>" & Server.HTMLEncode(custArr(2,k)) & "</campo1>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo1></campo1>"
					end if
					if (custArr(3,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo2>" & Server.HTMLEncode(custArr(3,k)) & "</campo2>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo2></campo2>"
					end if
					if (custArr(4,k)<>"") then
						tmpXMLDoc=tmpXMLDoc & "<campo3>" & Server.HTMLEncode(custArr(4,k)) & "</campo3>"
					else
						tmpXMLDoc=tmpXMLDoc & "<campo3></campo3>"
					end if
					tmpXMLDoc=tmpXMLDoc & "</subscriber>"
				else
					tmpXMLDoc=tmpXMLDoc & "<subscriber email=""" & custArr(1,k) & """ />"
				end if
			Next
			tmpXMLDoc="<subscribers>" & tmpXMLDoc & "</subscribers>"
			query="SELECT pcMailUpLists_ID,pcMailUpLists_ListID,pcMailUpLists_ListGuid FROM pcMailUpLists WHERE pcMailUpLists_Active=1 AND pcMailUpLists_Removed=0;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				listArr=rs.getRows()
				listCount=ubound(listArr,2)
				set rs=nothing
				tmpListID=""
				tmpListGuid=""
				tmpGroupID=""
				For k=0 to listCount
					if tmpListID<>"" then
						tmpListID=tmpListID & ";"
						tmpListGuid=tmpListGuid & ";"
						tmpGroupID=tmpGroupID & ";"
					end if
					tmpListID=tmpListID & listArr(1,k)
					tmpListGuid=tmpListGuid & listArr(2,k)
					tmpGroupID=tmpGroupID & "0"
				Next
				
				tmpMUResult=MUImport(session("CP_MU_APIUser"),session("CP_MU_APIPassword"),session("CP_MU_URL"),tmpListID,tmpListGuid,tmpXMLDoc,tmpGroupID,0,0,conFirmEmail)
				
				if tmpMUResult=0 then
					tmpSuccess=0
				else
					tmpReturnIDs=split(session("CP_MU_ReturnIDs"),";")
					tmpReturnList=split(session("CP_MU_ReturnList"),";")
					tmpReturnCode=split(session("CP_MU_ReturnCode"),";")
					For k=0 to listCount
						tmpCorrect=0
						For m=0 to listCount
							if clng(listArr(1,k))=clng(tmpReturnList(m)) then
								tmpMUResult1=tmpReturnIDs(m)
								tmpCorrect=m
								exit for
							end if
						Next
							
						if clng(tmpMUResult1)=0 then
							tmpSuccess=0
							tmpCanNotStart=tmpCanNotStart+1
						end if
						tmpIDProList=tmpIDProList & tmpMUResult1 & "*" & tmpReturnCode(tmpCorrect) & "||"
						tmpIDLists=tmpIDLists & listArr(1,k) & "||"
						IF clng(tmpMUResult1)>0 THEN
						tmpSuccessCount=tmpSuccessCount+1
						For l=0 to custCount
							query="SELECT pcMailUpSubs_ID FROM pcMailUpSubs WHERE idcustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
							set rsQ=connTemp.execute(query)
							dtTodaysDate=Date()
							if SQL_Format="1" then
								dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
							else
								dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
							end if
							if not rsQ.eof then
								if scDB="SQL" then
									query="UPDATE pcMailUpSubs SET idCustomer=" & custArr(0,l) & ",pcMailUpLists_ID=" & listArr(0,k) & ",pcMailUpSubs_LastSave='" & dtTodaysDate & "',pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
								else
									query="UPDATE pcMailUpSubs SET idCustomer=" & custArr(0,l) & ",pcMailUpLists_ID=" & listArr(0,k) & ",pcMailUpSubs_LastSave=#" & dtTodaysDate & "#,pcMailUpSubs_SyncNeeded=0,pcMailUpSubs_Optout=0 WHERE idCustomer=" & custArr(0,l) & " AND pcMailUpLists_ID=" & listArr(0,k) & ";"
								end if
							else
								if scDB="SQL" then
									query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & custArr(0,l) & "," & listArr(0,k) & ",'" & dtTodaysDate & "',0,0);"
								else
									query="INSERT INTO pcMailUpSubs (idCustomer,pcMailUpLists_ID,pcMailUpSubs_LastSave,pcMailUpSubs_SyncNeeded,pcMailUpSubs_Optout) VALUES (" & custArr(0,l) & "," & listArr(0,k) & ",#" & dtTodaysDate & "#,0,0);"
								end if
							end if
							set rsQ=nothing
							set rsQ=connTemp.execute(query)
							set rsQ=nothing
						Next
						END IF
					Next
				end if
				if tmpSuccess=1 then
					msg="2A"
				else
					if tmpSuccess=0 AND tmpSuccessCount>0 then
						msg="2B"
					else
						msg="2C"
					end if
				end if
			end if 'Have Active Lists
			set rs=nothing
		end if ' Have Imported Customers
		set rs=nothing
		if tmpIDProList<>"" then
			query="UPDATE pcMailUpSettings SET pcMailUpSett_LastIDList='" & tmpIDLists & "',pcMailUpSett_LastIDProcess='" & tmpIDProList & "';"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
	END IF
	
	'Synchronization
	IF request("submit3")<>"" then
		tmpSuccess=1
		tmpSuccessCount=0
		conFirmEmail=request("confirm3")
		if conFirmEmail="" then
			conFirmEmail=0
		end if
		if conFirmEmail="0" then
			pcMU_sendconfirm="0"
		end if
		if conFirmEmail="1" then
			pcMU_sendconfirm="1"
		end if
		query="SELECT customers.idcustomer,customers.email,pcMailUpLists.pcMailUpLists_ListID,pcMailUpLists_ListGuid,pcMailUpSubs.pcMailUpSubs_Optout FROM customers,pcMailUpLists,pcMailUpSubs WHERE customers.idcustomer=pcMailUpSubs.idcustomer AND pcMailUpLists.pcMailUpLists_ID=pcMailUpSubs.pcMailUpLists_ID AND pcMailUpSubs_SyncNeeded=1;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			custArr=rs.getRows()
			custCount=ubound(custArr,2)
			set rs=nothing
			For k=0 to custCount
				if custArr(4,k)="0" then
					tmpMUResult=UpdUserReg(custArr(0,k),custArr(1,k),custArr(2,k),custArr(3,k),session("CP_MU_URL"),1)
					if tmpMUResult=0 then
						tmpSuccess=0
					else
						tmpSuccessCount=tmpSuccessCount+1
					end if
				else
					tmpMUResult=UnsubUser(custArr(0,k),custArr(1,k),custArr(2,k),custArr(3,k),session("CP_MU_URL"),1)
					if tmpMUResult=0 then
						tmpSuccess=0
					else
						tmpSuccessCount=tmpSuccessCount+1
					end if
				end if
			Next
			if tmpSuccess=1 then
				msg="3A"
			else
				if tmpSuccess=0 AND tmpSuccessCount>0 then
					msg="3B"
				else
					msg="3C"
				end if
			end if
		end if ' Have Imported Customers
		set rs=nothing
	END IF
call closedb()
END IF
'End of Post Back
%>
<!--#include file="AdminHeader.asp"-->
<script>
function newWindow(file,window) {
	msgWindow=open(file,window,'scrollbars=yes,resizable=yes,,width=530,height=150');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
</script>
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%tmpNeedSync=0%>
<form name="form1" action="mu_regsyn.asp?action=run" method="post" class="pcForms" onsubmit="javascript:pcf_Open_MailUp();">
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<%if msg<>"" then%>
	<tr>
		<td>
			<div class="pcCPmessage">
				<%Select Case msg
				Case "1A":%>
					All opt-in customers have been successfully exported to your MailUp console.
					<%if tmpCanNotStart>"0" then%>
						<br>However, some of the import processes have not yet been completed. So the exported contacts may not yet appear among the recipients listed in your MailUp console.
					<%end if%>
				<%Case "1B":%>
					<img src="images/pcadmin_note.gif"> Some opt-in customers could not be exported to your MailUp console.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
					<%if tmpCanNotStart>"0" then%>
						<br>In addition, some of the import processes have not yet been completed. So the exported contacts may not yet appear among the recipients listed in your MailUp console.
					<%end if%>
				<%Case "1C":%>
					<img src="images/pcadmin_note.gif"> The system could not register any opt-in customers.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
				<%Case "2A":%>
					<img src="images/pcadmin_successful.gif"> All imported customers have been successfully exported to your MailUp console.
					<%if tmpCanNotStart>"0" then%>
						<br>However, some of the import processes have not yet been completed. So the exported contacts may not yet appear among the recipients listed in your MailUp console.
					<%end if%>
				<%Case "2B":%>
					<img src="images/pcadmin_note.gif"> Some imported customers could not be exported to your MailUp console.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
					<%if tmpCanNotStart>"0" then%>
						<br>In addition, some of the import processes have not yet been completed. So the exported contacts may not yet appear among the recipients listed in your MailUp console.
					<%end if%>
				<%Case "2C":%>
					<img src="images/pcadmin_note.gif"> The system could not register any imported customers.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
				<%Case "3A":%>
					<img src="images/pcadmin_successful.gif"> All records have been synchronized successfully.
				<%Case "3B":%>
					<img src="images/pcadmin_note.gif"> Some records could not be synchronized.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
				<%Case "3C":%>
					<img src="images/pcadmin_note.gif"> The system could not synchronize any records.<%if MU_ErrMsg<>"" then%>&nbsp;Error Message: <%=MU_ErrMsg%><%end if%>
				<%End Select%>
			</div>
		</td>
	</tr>
	<%end if%>
	<%
	call opendb()
	query="SELECT pcMailUpSett_LastIDList,pcMailUpSett_LastIDProcess FROM pcMailUpSettings WHERE pcMailUpSettings.pcMailUpSett_LastIDProcess<>'';"
	set rs=connTemp.execute(query)
	if not rs.eof then%>
		<tr>
			<th>Last Import Process Status</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<a onclick="javascript:pcf_Open_MailUp();" href="mu_history.asp">View import status &gt;&gt;</a>
			</td>
		</tr>
		<%
	end if
	set rs=nothing
	call closedb()%>
	<%IF tmp_bulk="0" THEN
	tmpNeedSync=1%>
	<tr>
		<th>Initial Setup</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
		<%
		call opendb()
		query="SELECT count(*) AS TotalCustomers FROM Customers WHERE RecvNews=1 AND customerType<3 AND suspend=0;"
		set rs=connTemp.execute(query)
		totalOptInCusts=0
		if not rs.eof then
			totalOptInCusts=rs("TotalCustomers")
		end if
		set rs=nothing
		call closedb()%>
		<%if totalOptInCusts="0" then%>
		<tr>
			<td>
				<div class="pcCPmessage">Your store does not have any opt-in customers</div>
			</td>
		</tr>
		<%else%>
		<tr>
			<td>
				<div>Your store has: <b><%=totalOptInCusts%></b> opt-in customer(s).</div>
				<div style="padding: 10px 0 5px 0">The customer(s) that opted to receive e-mail messages from the store will be registered as <strong>opted into all active lists</strong> (<a href="mu_settings.asp">define which lists are active &gt;&gt;</a>). They will then be able to individually opt out of any list either by editing their profile in the storefront, or by clicking on the <em>unsubscribe</em> link on any of the messages you will send.</div>
				<div style="padding: 15px 0 5px 0">Would you like to <strong>ask customers to confirm</strong> that they want to subscribe to the lists that they will be registered in?</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm1" value="1" class="clearBorder" checked="checked"> Yes, <strong>send Subscription Confirmation Request</strong> message.</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm1" value="0" class="clearBorder"> No, customers don't need to confirm their subscription.</div>
				<div>If you select &quot;Yes&quot;, MailUp will send customers a subscription confirmation request e-mail for each of the lists that the customer is added to. You can <strong>review and edit</strong> the copy of the message through your MailUp console (<em>Settings > E-mail Confirmation Request</em>).</div>
				<div style="padding: 15px 0 20px 0">
				<input type="submit" name="submit1" class="submit2" value="Register Opt-in Customers">
				</div>
			</td>
		</tr>
		<%end if%>
	<%END IF%>
		<%
		call opendb()
		totalImportedCusts=0
		query="SELECT pcMailUpSett_LastCustomerID FROM pcMailUpSettings;"
		set rs=connTemp.execute(query)
		tmpCustomers=""
		if not rs.eof then
			tmpCustomers=rs("pcMailUpSett_LastCustomerID")
		end if
		set rs=nothing
		
		if tmpCustomers<>"" then
			tmp1=split(tmpCustomers,",")
			query="SELECT count(*) AS TotalCustomers FROM Customers WHERE idCustomer>" & tmp1(0) & " AND idCustomer<=" & tmp1(1) & " AND customerType<3 AND suspend=0;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				totalImportedCusts=rs("TotalCustomers")
			end if
			set rs=nothing
		end if
		call closedb()
		%>
		<%if totalImportedCusts>"0" then
		tmpNeedSync=1%>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<th>Imported Customers</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				<div>Your store has: <b><%= totalImportedCusts%></b> imported customer(s).</div>
				<div style="padding: 10px 0 5px 0">All of these customer(s) will be registered as <strong>opted into all active lists</strong> (<a href="mu_settings.asp">define which lists are active &gt;&gt;</a>). They will then be able to individually opt out of any list either by editing their profile in the storefront, or by clicking on the <em>unsubscribe</em> link on any of the messages you will send.</div>
				<div style="padding: 15px 0 5px 0">You should <strong>ask customers to confirm their subscription</strong> since they did not opt into these lists on your Web site.</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm2" value="1" class="clearBorder" checked="checked"> Yes, <strong>send Subscription Confirmation Request</strong> message.</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm2" value="0" class="clearBorder"> No, customers don't need to confirm their subscription.</div>
				<div>If you select &quot;Yes&quot;, MailUp will send customers a subscription confirmation request e-mail for each of the lists that the customer is added to. You can <strong>review and edit</strong> the copy of the message through your MailUp console (<em>Settings > E-mail Confirmation Request</em>).</div>
				<div style="padding: 15px 0 20px 0">
				<input type="submit" name="submit2" class="submit2" value="Register Imported Customers">
				</div>
			</td>
		</tr>
		<%end if%>
		<%
		call opendb()
		query="SELECT count(*) AS TotalRecords FROM pcMailUpSubs WHERE pcMailUpSubs_SyncNeeded<>0;"
		set rs=connTemp.execute(query)
		totalRecords=0
		if not rs.eof then
			totalRecords=rs("TotalRecords")
		end if
		set rs=nothing
		call closedb()%>
		<%if totalRecords>"0" then
		tmpNeedSync=1%>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<th>Synchronization</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
				Your store has: <b><%=totalRecords%></b> record(s) that are waiting for synchronization. (<a href="javascript:newWindow('mu_synlist.asp','window1');">See who they are</a>)<br>
				Click the button below to synchronize them with your MailUp console.<br><br>
				<div style="padding: 15px 0 5px 0">You should <strong>ask customers to confirm their subscription</strong> since they have not yet done so.</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm3" value="1" class="clearBorder" checked="checked"> Yes, <strong>send Subscription Confirmation Request</strong> message.</div>
				<div style="padding: 0 0 5px 0"><input type="radio" name="confirm3" value="0" class="clearBorder"> No, customers don't need to confirm their subscription.</div>
				<div>If you select &quot;Yes&quot;, MailUp will send customers a subscription confirmation request e-mail for each of the lists that the customer is added to. You can <strong>review and edit</strong> the copy of the message through your MailUp console (<em>Settings > E-mail Confirmation Request</em>).</div>
				<div style="padding: 15px 0 20px 0">
				<input type="submit" name="submit3" class="submit2" value="Synchronize">
				</div>
				<br /><br /><br />
			</td>
		</tr>
		<%end if%>
		<%if tmpNeedSync=0 then%>
		<tr>
		<td colspan="2">
			<br>
			<br>
			<div class="pcCPmessage">There is no data to synchronize at this time.</div>
			<br>
			<br>
		</td>
		</tr>
		<%end if%>
		<tr>
		<td colspan="2">
			<input type="button" name="Back" value="Back" onClick="location='mu_manageNewsWiz.asp';" class="ibtnGrey">
		</td>
	</tr>
</table>
</form>
<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote2"),"MailUp", 300))%><!--#include file="AdminFooter.asp"-->