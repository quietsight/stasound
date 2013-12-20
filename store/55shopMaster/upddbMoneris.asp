<%@LANGUAGE="VBSCRIPT"%>
<% On Error Resume Next %>
<% '--Updater File-- %>
<% pageTitle = "Moneris Gateway -   Database Update" %>
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

tmp_mode=request("mode")
tmp_id=request("id")

Response.Cookies("AgreeLicense") = ""
IF request("action")="sql" then
	if request.querystring("hmode")="2" then
		SSIP=request("SSIP")
		UID=request("UID")
		PWD=request("PWD")
		SSDB=request("SSDB")
		if SSIP="" or UID="" or PWD="" then
			response.redirect "upddbMoneris.asp?mode=3"
			response.End
		end if
		set connTemp=server.createobject("adodb.connection")
		connTemp.Open "Driver={SQL Server};Server="&SSIP&";Address="&SSIP&",1433;Network=DBMSSOCN;Database="&SSDB&";Uid="&UID&";Pwd="&PWD
		if err.number <> 0 then
			response.redirect "techErr.asp?error="&Server.Urlencode("Error while opening database")
		end if
	else
		if instr(ucase(scDSN),"DSN=") then
			response.redirect "upddbMoneris.asp?mode=1"
			response.End
		end if
		call openDb()
	end if
	
	iCnt=0
	ErrStr=""		
  
	'//  Create column pcPay_Moneris_Meth for table pcPay_Moneris
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Meth] [int] NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	else
		query="UPDATE [pcPay_Moneris] SET pcPay_Moneris_Meth=0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="ALTER TABLE [pcPay_Moneris] ADD CONSTRAINT [DF_pcPay_Moneris_pcPay_Moneris_Meth] DEFAULT (0) FOR [pcPay_Moneris_Meth];"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if
  
	'//  Create column pcPay_Moneris_Interac for table pcPay_Moneris 
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Interac] [int] NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	else
		query="UPDATE pcPay_Moneris SET pcPay_Moneris_Interac=0;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		query="ALTER TABLE [pcPay_Moneris] ADD CONSTRAINT [DF_pcPay_Moneris_pcPay_Moneris_Interac] DEFAULT (0) FOR [pcPay_Moneris_Interac];"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
	end if


	'//  Create column pcPay_Moneris_Interac_MerchantID for table pcPay_Moneris 
	query="ALTER TABLE [pcPay_Moneris] ADD [pcPay_Moneris_Interac_MerchantID] [nvarchar] (20) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_Moneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	
	query="ALTER TABLE [pcPay_OrdersMoneris] ADD [pcPay_MOrder_IDEBIT_ISSCONF] [nvarchar] (30) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_OrdersMoneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
	end if
	
	
	query="ALTER TABLE [pcPay_OrdersMoneris] ADD [pcPay_MOrder_IDEBIT_ISSNAME] [nvarchar] (30) NULL"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	if err.number <> 0 then
		if instr(ucase(Err.Description),"COLUMN NAMES IN EACH TABLE MUST BE UNIQUE") then
			Err.Description=""
			err.number=0
		else
			ErrStr=ErrStr&"Unable to update TABLE pcPay_OrdersMoneris - Error: "&Err.Description&"<BR>"
			err.number=0
			iCnt=iCnt+1
		end if
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
		strMode="errors"
	else
		strMode="complete"
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

		query="ALTER TABLE pcPay_Moneris ADD pcPay_Moneris_Meth LONG NULL;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp1.execute(query)
		if err.number <> 0 then
			call TrapError("pcPay_Moneris")
		else
			query="UPDATE pcPay_Moneris SET pcPay_Moneris_Meth=0;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp1.execute(query)
			query="ALTER TABLE pcPay_Moneris ALTER COLUMN pcPay_Moneris_Meth LONG NULL DEFAULT '0';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp1.execute(query)
		end if
		
		query="ALTER TABLE pcPay_Moneris ADD pcPay_Moneris_Interac LONG NULL;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp1.execute(query)
		if err.number <> 0 then
			call TrapError("pcPay_Moneris")
		else
			query="UPDATE pcPay_Moneris SET pcPay_Moneris_Interac=0;"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp1.execute(query)
			query="ALTER TABLE pcPay_Moneris ALTER COLUMN pcPay_Moneris_Interac LONG NULL DEFAULT '0';"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp1.execute(query)
		end if
		
		query="ALTER TABLE pcPay_Moneris ADD pcPay_Moneris_Interac_MerchantID varchar(20) NULL;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp1.execute(query)
		if err.number <> 0 then
			call TrapError("pcPay_Moneris")
		end if
		
		query="ALTER TABLE pcPay_OrdersMoneris ADD pcPay_MOrder_IDEBIT_ISSCONF varchar(30) NULL;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp1.execute(query)
		if err.number <> 0 then
			call TrapError("pcPay_Moneris")
		end if
		
		query="ALTER TABLE pcPay_OrdersMoneris ADD pcPay_MOrder_IDEBIT_ISSNAME varchar(30) NULL;"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp1.execute(query)
		if err.number <> 0 then
			call TrapError("pcPay_Moneris")
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
			strMode="errors"
		else
			strMode="complete"
		end if
	end if
End If

%>
<!--#include file="Adminheader.asp"-->
<form action="upddbMoneris.asp" method="get" name="form1" class="pcForms">
<% 
'mode=Edit&id=219&gwchoice=11
if strMode="complete" then
	response.redirect "AddModRTPayment.asp?mode="&tmp_mode&"&id="&tmp_id&"&gwchoice=11"
	response.end
else %>
	<input type="hidden" name="mode" value="<%=tmp_mode%>">
	<input type="hidden" name="id" value="<%=tmp_id%>">
	<table class="pcCPcontent" style="width:600px;" align="center">
		<tr>
			<td class="pcCPspacer" align="center"></td>
		</tr>

		<% if strMode="errors" then %>
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
				<span style="font-weight: bold">You must update your database before you can add/modify Moneris Gateway to support Verify By Visa And Interac&reg; online.<u></u></span><br>
				In order to activate Moneris real-time payment gateway you will need to update your ProductCart database to add the required table.</div>				<p><strong><br>
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
						  <input type="button" name="access" value="Update Your ProductCart MS Access Database" onClick="location='upddbMoneris.asp?action=access&mode=<%=tmp_mode%>&id=<%=tmp_id%>';" class="submit2">
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
					  <input type="button" name="access2" value="Update Your ProductCart MS SQL Database" onClick="location='upddbMoneris.asp?action=sql&mode=<%=tmp_mode%>&id=<%=tmp_id%>';" class="submit2">
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
