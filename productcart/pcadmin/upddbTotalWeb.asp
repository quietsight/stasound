<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "Ogone Gateway - Database Update" %>
<% Section="paymntOpt" %>
<%PmAdmin=5%><!--#include file="adminv.asp"-->  
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<% 
dim f, mySQL, conntemp, rstemp, rs, iCnt
dim pcArr, intCount,i,j,query,pcArr1, intCount1,pcArr2, intCount2

Response.Cookies("AgreeLicense") = ""
IF request("action")="sql" then
	if request.querystring("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			response.redirect "upddbTotalWeb.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Driver={SQL Server};Server="&SSIP&";Address="&SSIP&",1433;Network=DBMSSOCN;Database="&SSDB&";Uid="&UID&";Pwd="&PWD
		if err.number <> 0 then
			response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			response.redirect "upddbTotalWeb.asp?mode=1"
			response.End
		end if
		call openDb()
	end if
	
	iCnt=0
	ErrStr=""		

	'**** Create table pcPay_TotalWeb ***************************************
	query="CREATE TABLE [dbo].[pcPay_TotalWeb] ("
	query=query&"[pcPay_TW_ID] [int] NULL  DEFAULT (1),"
	query=query&"[pcPay_TW_MerchantID] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_TW_CurCode] [nvarchar] (3) NULL,"
	query=query&"[pcPay_TW_TestMode] [int] NULL DEFAULT (0) "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_TotalWeb")
	else
		query="INSERT INTO pcPay_TotalWeb (pcPay_TW_ID,pcPay_TW_TestMode) VALUES (1,1);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if

	set rs=nothing
	set connTemp=nothing
	
	Function TrapSQLError(varTableName)		
		'// -2147217900 = Table 'x' already exists.
		'// -2147217887 = Field 'x' already exists in table 'x'.
		if ((Err.Number=-2147217900) OR (Err.Number=-2147217887)) then
			Err.Description=""
			err.number=0
		else
			ErrStr = ErrStr & "Error Creating Table "&varTableName&": "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	End Function

	If iCnt>0 then
		mode="errors"
	else
		mode="complete"
	end if

'// ************ START ACCESS **********************************
ELSE
	if request("action")="access" then
		iCnt=0
		ErrStr=""
		if instr(ucase(scDSN),"DBQ=") then
			dim tempDSN, tempStr, tempDS, tempStr2, B, C, tempPWD
			tempDSN=scDSN
			tempDSN=replace(tempDSN,"dbq=","DBQ=")
			tempDSN=replace(tempDSN,"Dbq=","DBQ=")
			tempDSN=replace(tempDSN,"DBq=","DBQ=")
			tempDSN=replace(tempDSN,"dBq=","DBQ=")
			tempDSN=replace(tempDSN,"dBQ=","DBQ=")
			tempDSN=replace(tempDSN,"dbQ=","DBQ=")
			tempDSN=replace(tempDSN,"pwd=","PWD=")
			tempDSN=replace(tempDSN,"Pwd=","PWD=")
			tempDSN=replace(tempDSN,"PWd=","PWD=")
			tempDSN=replace(tempDSN,"pWd=","PWD=")
			tempDSN=replace(tempDSN,"pWD=","PWD=")
			tempDSN=replace(tempDSN,"pwD=","PWD=")
			tempDSN=replace(tempDSN,"Password=","password=")
			tempDSN=replace(tempDSN,"PASSWORD=","password=")
			tempPWD=""
			A=split(tempDSN,"DBQ=")
			tempStr=A(0)
			tempDS=A(1)
			if instr(tempDS,";") then
				B=split(tempDS,";")
				tempDS=B(0)
				tempStr=B(1)
				tempStr=replace(tempStr,"pwd=","PWD=")
				tempStr=replace(tempStr,"Pwd=","PWD=")
				tempStr=replace(tempStr,"PWd=","PWD=")
				tempStr=replace(tempStr,"pWd=","PWD=")
				tempStr=replace(tempStr,"pWD=","PWD=")
				tempStr=replace(tempStr,"pwD=","PWD=")
				tempStr=replace(tempStr,"Password=","password=")
				tempStr=replace(tempStr,"PASSWORD=","password=")
				if instr(tempStr,"PWD=") then
					C=split(tempStr,"=")
					if ubound(C)=1 then
						tempPwd=C(1)
					end if
				end if
				if instr(tempStr,"password=") then
					C=split(tempStr,"=")
					if ubound(C)=1 then
						tempPwd=C(1)
					end if
				end if
				if ubound(B)>1 then
					tempStr2=B(2)
					response.write "tempPwd2="&tempStr2&"<BR><BR>"
					if instr(tempStr2,"PWD=") then
						C=split(tempStr2,"PWD=")
						tempPwd=C(1)
					end if
				end if
				if ubound(B)>1 then
					tempStr2=B(2)
					if instr(tempStr2,"password=") then
						C=split(tempStr2,"password=")
						tempPwd=C(1)
					end if
				end if
			end if
			scDSN1="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & tempDS
			if tempPwd<>"" then
				scDSN1=scDSN1&";Jet OLEDB:Database Password="&tempPwd&"; "&pUserId &" "& pUserPassword2 &""
			end if
		else
			if instr(ucase(scDSN), "DATA SOURCE") then
				scDSN1=scDSN
			else
				call openDB()
				myMdbFile = connTemp.Properties("Current Catalog") & ".mdb"
				call closeDB()
				scDSN1="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & myMdbFile
				
				if instr(ucase(scDSN),"PWD") then
					A=split(scDSN,"PWD=")
					scDSN1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&myMdbFile&";Jet OLEDB:Database Password="&A(1)&"; "&pUserId &" "& pUserPassword2 &""
				end if
			end if
		end if

		err.clear
		set connTemp1=server.createobject("adodb.connection")
		connTemp1.Open scDSN1  
		
		if err.number <> 0 then
			response.write err.description
			response.end
		end if

		'**** Create table pcPay_TotalWeb ***************************************
		query="CREATE TABLE pcPay_TotalWeb ("
		query=query&"pcPay_TW_ID INTEGER NOT NULL DEFAULT 1 ,"
		query=query&"pcPay_TW_MerchantID VarChar(100) NULL ,"		
		query=query&"pcPay_TW_CurCode VarChar(3) NULL ,"		
		query=query&"pcPay_TW_TestMode INTEGER NOT NULL DEFAULT 0 "
		query=query&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp1.execute(query)
		if err.number <> 0 then
			TrapError("pcPay_TotalWeb")
		else
			query="INSERT INTO pcPay_TotalWeb (pcPay_TW_ID,pcPay_TW_TestMode) VALUES (1,1);"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp1.execute(query)
		end if

		set rs=nothing
		set connTemp1=nothing
		
		Function TrapError(varTableName)
			'// -2147217900 = Table 'x' already exists.
			'// -2147217887 = Field 'x' already exists in table 'x'.
			if ((Err.Number=-2147217900) OR (Err.Number=-2147217887)) then
				Err.Description=""
				err.number=0
			else
				ErrStr = ErrStr & "Error Creating Table "&varTableName&": "&Err.Description&"<BR>"
				err.number=0
				iCnt=iCnt+1
			end if
		End Function
		
		If iCnt>0 then
			mode="errors"
		else
			mode="complete"
		end if
	end if
End If

%>
<!--#include file="Adminheader.asp"-->
<form action="upddbTotalWeb.asp" method="get" name="form1" class="pcForms">
<%
if mode="complete" then 
	response.redirect "AddModRTPayment.asp?gwchoice=63"
	response.end
else %>
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if mode="errors" then %>
			<tr>
				<td align="center">
					<div class="pcCPmessage">The following errors occurred while trying to update your database:<br><br>
				<%=ErrStr%></div>
				</td>
			</tr>
		<% end if %>

		<tr>
			<td>
                <div class="pcCPnotes" style="padding:10px;">
				<span style="font-weight: bold">You must update your database before you can add Ogone Gateway.<u></u></span><br>
				In order to activate Ogone real-time payment gateway you will need to update your ProductCart database to add the required table.</div>				<p><strong><br>
				    <br>
				  You are about to update the store database. Please read the following carefully before proceeding.</strong></p>
				<p style="padding-top:10px;">Although we have tested this update script in a variety of environments, there is always the possibility of something going wrong. Make sure to <span style="font-weight: bold">backup your database</span> prior to executing this update. To do so:
				<ul>
				<li>If you are using an Access database, simply download a copy of the database to your local system.</li>
				<li>If you are using a SQL database, depending on how the database has been setup, you may be able to either perform the backup yourself or have your Web hosting company do it for you. Note: your SQL database is likely being automatically backed up every day: confirm that this is the case by asking your Web host when the last back up occurred.</li>
				</ul>
				</p>
				<br />
			<table class="pcCPcontent" width="80%">
			<% if scDB="Access" then %>
					<tr>
						<td align="center">
						  <input type="button" name="access" value="Update Your ProductCart MS Access Database" onClick="location='upddbTotalWeb.asp?action=access';" class="submit2">
						</td>
					</tr>
			<% 
				else
					if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
				<tr>
					<td>
					It appears that you are using a DSN connection to connect to your SQL server. In order to complete this update, please enter your SQL Server Information below:
					<% if request.querystring("mode")="1" then %>
						<br>
						<strong>*All fields are required.</strong>
					<% end if %>					</td>
				</tr>
				<tr>
					<td>Server Domain/IP:	<input name="SSIP" type="text" id="SSIP" size="30"></td>
				</tr>
				<tr>
					<td>Database	Name:	<input name="SSDB" type="text" id="SSDB" size="30"></td>
				</tr>
				<tr>
					<td>User ID: <input name="UID" type="text" id="UID" size="30"></td>
				</tr>
				<tr>
					<td>Password: <input name="PWD" type="password" id="PWD" size="30"></td>
				</tr>
				<input name="hmode" type="hidden" id="hmode" value="2">
				<input name="action" type="hidden" id="action" value="sql">
			<% end if %>
				<tr>
					<td align="center">
					<% if request.querystring("mode")="1" OR request.querystring("mode")="3" then %>
							<input name="access2" type="submit" id="access2" value="Update Your ProductCart MS SQL Database" class="submit2">
						<% else %>
					  <input type="button" name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbTotalWeb.asp?action=sql';" class="submit2">
						<% end if %>
					</td>
			</tr>
			<% end if %>
			</table>
		</td>
	</tr>
</table>
<% end if %>
</form>
<!--#include file="AdminFooter.asp"-->