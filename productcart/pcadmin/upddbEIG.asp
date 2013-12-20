<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "NetSource Commerce Gateway - Database Update" %>
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
			response.redirect "upddbEIG.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Driver={SQL Server};Server="&SSIP&";Address="&SSIP&",1433;Network=DBMSSOCN;Database="&SSDB&";Uid="&UID&";Pwd="&PWD
		if err.number <> 0 then
			response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			response.redirect "upddbEIG.asp?mode=1"
			response.End
		end if
		call openDb()
	end if
	
	iCnt=0
	ErrStr=""		

	'// Create table pcPay_EIG
	query="CREATE TABLE [dbo].[pcPay_EIG] ("
	query=query&"[pcPay_EIG_ID] [int] NULL  DEFAULT (1),"
	query=query&"[pcPay_EIG_Username] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Password] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Key] [nvarchar] (100) NULL ,"
	query=query&"[pcPay_EIG_Type] [nvarchar] (50) NULL ,"
	query=query&"[pcPay_EIG_Version] [nvarchar] (4) NULL ,"
	query=query&"[pcPay_EIG_Curcode] [nvarchar] (4) NULL ,"
	query=query&"[pcPay_EIG_CVV] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_SaveCards] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_UseVault] [int] NULL DEFAULT(0) ,"
	query=query&"[pcPay_EIG_TestMode] [int] NULL DEFAULT(0)"
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG")
	else
		query="INSERT INTO pcPay_EIG (pcPay_EIG_ID, pcPay_EIG_CVV, pcPay_EIG_TestMode) VALUES (1,0,1);"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
	
	'// Create table pcPay_EIG_Vault
	query="CREATE TABLE [dbo].[pcPay_EIG_Vault] ("
	query=query&"[pcPay_EIG_Vault_ID] [INT] IDENTITY (1, 1) NOT FOR REPLICATION NOT NULL ,"
	query=query&"[idOrder] [int] NULL  DEFAULT (1),"
	query=query&"[idCustomer] [int] NULL  DEFAULT (1),"
	query=query&"[IsSaved] [int] NULL  DEFAULT (1),"
	query=query&"[pcPay_EIG_Vault_CardNum] [nvarchar] (25) NULL ,"
	query=query&"[pcPay_EIG_Vault_CardType] [nvarchar] (10) NULL ,"
	query=query&"[pcPay_EIG_Vault_CardExp] [nvarchar] (10) NULL ,"
	query=query&"[pcPay_EIG_Vault_Token] [nvarchar] (50) NULL "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG_Vault")
	end if		
	
	'// Create table pcPay_EIG_Authorize
	query="CREATE TABLE [dbo].[pcPay_EIG_Authorize] ("
	query=query&"[idauthorder] [INT] IDENTITY (1, 1) NOT FOR REPLICATION NOT NULL ,"
	query=query&"[idOrder] [int] NULL  DEFAULT (1),"
	query=query&"[idCustomer] [int] NULL  DEFAULT (1),"
	query=query&"[captured] [int] NULL  DEFAULT (1),"
	query=query&"[pcSecurityKeyID] [int] NULL  DEFAULT (1),"
	query=query&"[vaultToken] [nvarchar] (50) NULL ,"
	query=query&"[amount] [money] NULL DEFAULT(0) ,"	
	query=query&"[paymentmethod] [nvarchar] (250) NULL ,"	
	query=query&"[transtype] [nvarchar] (250) NULL ,"
	query=query&"[authcode] [nvarchar] (250) NULL ,"
	query=query&"[ccnum] [nvarchar] (250) NULL ,"
	query=query&"[ccexp] [nvarchar] (10) NULL ,"
	query=query&"[cctype] [nvarchar] (25) NULL ,"	
	query=query&"[fname] [nvarchar] (250) NULL ,"	
	query=query&"[lname] [nvarchar] (250) NULL ,"	
	query=query&"[address] [nvarchar] (250) NULL ,"
	query=query&"[zip] [nvarchar] (25) NULL ,"	
	query=query&"[trans_id] [nvarchar] (250) NULL "
	query=query&");"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		TrapSQLError("pcPay_EIG_Vault")
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

		'// Create table pcPay_EIG
		query="CREATE TABLE pcPay_EIG ("
		query=query&"pcPay_EIG_ID INTEGER NOT NULL DEFAULT 1 ,"
		query=query&"pcPay_EIG_Username VarChar(100) NULL ,"
		query=query&"pcPay_EIG_Password VarChar(100) NULL ,"
		query=query&"pcPay_EIG_Key VarChar(100) NULL ,"
		query=query&"pcPay_EIG_Type VarChar(50) NULL ,"
		query=query&"pcPay_EIG_Version VarChar(4) NULL ,"
		query=query&"pcPay_EIG_Curcode VarChar(4) NULL ,"
		query=query&"pcPay_EIG_CVV INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"pcPay_EIG_SaveCards INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"pcPay_EIG_UseVault INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"pcPay_EIG_TestMode INTEGER NOT NULL DEFAULT 0 "
		query=query&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp1.execute(query)
		if err.number <> 0 then
			TrapError("pcPay_EIG")
		else
			query="INSERT INTO pcPay_EIG (pcPay_EIG_ID, pcPay_EIG_CVV, pcPay_EIG_TestMode) VALUES (1,0,1);"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp1.execute(query)
		end if

		'// Create table pcPay_EIG_Vault
		query="CREATE TABLE pcPay_EIG_Vault ("
		query=query&"pcPay_EIG_Vault_ID Counter NOT NULL PRIMARY KEY UNIQUE ,"	
		query=query&"idOrder  INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"idCustomer  INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"IsSaved  INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"pcPay_EIG_Vault_CardNum VarChar(25) NULL ,"
		query=query&"pcPay_EIG_Vault_CardType VarChar(10) NULL ,"
		query=query&"pcPay_EIG_Vault_CardExp VarChar(10) NULL ,"
		query=query&"pcPay_EIG_Vault_Token VarChar(50) NULL ,"
		query=query&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp1.execute(query)
		if err.number <> 0 then
			TrapError("pcPay_EIG_Vault")
		end if
		
		'// Create table pcPay_EIG_Authorize
		query="CREATE TABLE pcPay_EIG_Authorize ("
		query=query&"[idauthorder] Counter NOT NULL PRIMARY KEY UNIQUE ,"
		query=query&"[idOrder] INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"[idCustomer] INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"[captured] INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"[pcSecurityKeyID] INTEGER NOT NULL DEFAULT 0 ,"
		query=query&"[vaultToken] VarChar(50) NULL ,"
		query=query&"[amount] Currency NULL DEFAULT '0' ,"	
		query=query&"[paymentmethod] VarChar(250) NULL ,"	
		query=query&"[transtype] VarChar(250) NULL ,"
		query=query&"[authcode] VarChar(250) NULL ,"
		query=query&"[ccnum] VarChar(250) NULL ,"
		query=query&"[ccexp] VarChar(10) NULL ,"
		query=query&"[cctype] VarChar(25) NULL ,"	
		query=query&"[fname] VarChar(250) NULL ,"	
		query=query&"[lname] VarChar(250) NULL ,"	
		query=query&"[address] VarChar(250) NULL ,"
		query=query&"[zip] VarChar(25) NULL ,"	
		query=query&"[trans_id] VarChar(250) NULL "
		query=query&");"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp1.execute(query)
		if err.number <> 0 then
			TrapError("pcPay_EIG_Authorize")
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
<form action="upddbEIG.asp" method="get" name="form1" class="pcForms">
<%
if mode="complete" then 
	response.redirect "AddModRTPayment.asp?gwchoice=66"
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
				<span style="font-weight: bold">You must update your database before you can add NetSource Commerce Gateway.<u></u></span><br>
				In order to activate EIG real-time payment gateway you will need to update your ProductCart database to add the required table.</div>				<p><strong><br>
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
						  <input type="button" name="access" value="Update Your ProductCart MS Access Database" onClick="location='upddbEIG.asp?action=access';" class="submit2">
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
					<td>Databas  Name:	<input name="SSDB" type="text" id="SSDB" size="30"></td>
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
					  <input type="button" name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbEIG.asp?action=sql';" class="submit2">
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