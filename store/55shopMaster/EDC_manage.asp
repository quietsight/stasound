<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Endicia Postage Label Services Settings" %>
<% response.Buffer=false %>
<% Section="shipOpt" %>
<%PmAdmin=4%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/EndiciaFunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%Dim connTemp,rs,query

pcPageName="EDC_manage.asp"

call opendb()

IF request("action")="reenable" THEN
	query="UPDATE pcEDCSettings SET pcES_Reg=0;"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	call closedb()
	response.redirect "EDC_manage.asp"
END IF
IF request("action")="checksignup" THEN
	if request("signedup")="" OR request("signedup")="0" then
		call closedb()
		response.redirect "EDC_signup.asp?reg=1"
	end if
	EDC_ErrMsg=""
	tmpUserID=request("edcaccountid")
	if tmpUserID="" then
		EDC_ErrMsg="Error: Please enter your Account ID<br>"
	end if
	if Not IsNumeric(tmpUserID) then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Please enter a number for your Account ID<br>"
	end if
	tmpPassP=request("edcpassp")
	tmpPassP1=request("edcpassp1")
	if tmpPassP="" OR tmpPassP1="" then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Please enter your Pass Phrase<br>"
	end if
	if tmpPassP<>tmpPassP1 then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Your Pass Phrase and Pass Phrase Confirmation are not the same<br>"
	end if
	
	if EDC_ErrMsg<>"" then
		call closedb()
		response.redirect pcPageName & "?msg=" & EDC_ErrMsg
	end if
	tmpPassP=enDeCrypt(tmpPassP, scCrypPass)
	query="SELECT pcES_UserID FROM pcEDCSettings;"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		query="UPDATE pcEDCSettings SET pcES_UserID=" & tmpUserID & ",pcES_PassP='" & tmpPassP & "',pcES_Reg=1;"
	else
		query="INSERT INTO pcEDCSettings (pcES_UserID,pcES_PassP,pcES_Reg,pcES_TestMode,pcES_LogTrans) VALUES (" & tmpUserID & ",'" & tmpPassP & "',1,1,1);"
	end if
	set rsQ=connTemp.execute(query)
	set rsQ=nothing

	msg="Your account information was added successfully!"
	msgType=1
END IF

call GetEDCSettings()

'Change Pass Phrase
IF request("action")="changepassp" THEN
	EDC_ErrMsg=""
	tmpUserID=request("edcaccountid")
	if tmpUserID="" then
		EDC_ErrMsg="Error: Please enter your Account ID<br>"
	end if
	if Not IsNumeric(tmpUserID) then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Please enter a number for your Account ID<br>"
	end if
	tmpCurrentPassP=request("edccurrentpassp")
	if tmpCurrentPassP="" then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Please enter your current Pass Phrase<br>"
	end if
	tmpPassP=request("edcpassp")
	tmpPassP1=request("edcpassp1")
	if tmpPassP="" OR tmpPassP1="" then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Please enter your new Pass Phrase<br>"
	end if
	if tmpPassP<>tmpPassP1 then
		EDC_ErrMsg=EDC_ErrMsg & "Error: Your new Pass Phrase and Pass Phrase Confirmation are not the same<br>"
	end if
	
	if EDC_ErrMsg<>"" then
		response.redirect pcPageName & "?msg=" & EDC_ErrMsg
	end if
	if (EDCUserID="0") OR (IsNull(EDCUserID)) then
		ComSign=1
	else
		ComSign=0
	end if
	tmpRun=ChangePassP(tmpUserID,tmpCurrentPassP,tmpPassP)
	if tmpRun="1" then
		if ComSign=1 then
			msg="Your account was activated successfully! You must deposit a minimum of $10 into your account before using the Endicia Postage label services."
			msgType=1
		else
			msg="Your Pass Phrase was updated successfully!"
			msgType=1
		end if
		call GetEDCSettings()
	else
		if ComSign=1 then
			msg="Cannot activate your Endicia account<br>Error: " & EDC_ErrMsg
			msgType=0
		else
			msg="Cannot update your Pass Phrase<br>Error: " & EDC_ErrMsg
			msgType=0
		end if
	end if
END IF

IF (request("edcaction")="refill") AND (request("amount")<>"") THEN
	msg=""
	tmpAmount=request("amount")
	if not IsNumeric(tmpAmount) then
		msg="Refill your Endicia Account error: The amount must be a numeric value"
		msgType=0
	else
		if (Cdbl(tmpAmount)<10) OR (Cdbl(tmpAmount)>99999.99) then
			msg="Refill your Endicia Account error: The Refill amount must be between: 10.00 - 99999.99"
			msgType=0
		end if
	end if
	if msg="" then
		call GetEDCSettings()
		FuncResult=BuyPostage(tmpAmount)
		if FuncResult<>"1" then
			msg="Refill your Endicia Account error: " & EDC_ErrMsg
			msgType=0
		else
			msg=EDC_SuccessMsg
			msgType=1
		end if
	end if
END IF

IF request("action")="updatesettings" THEN
	tmpEDCTest=request("edctest")
	if tmpEDCTest="" then
		tmpEDCTest="1"
	end if
	tmpEDCLog=request("edclog")
	if tmpEDCLog="" then
		tmpEDCLog="0"
	end if
	tmpEDCAutoFill=request("edcautofill")
	if tmpEDCAutoFill="" then
		tmpEDCAutoFill="0"
	end if	
	tmpEDCTrigger=request("edctrigger")
	if tmpEDCTrigger="" then
		tmpEDCTrigger="0"
	end if
	tmpEDCFillAmount=request("edcfillamount")
	if tmpEDCFillAmount="" then
		tmpEDCFillAmount="0"
	end if
	tmpEDCAutoRmv=request("edcautormv")
	if tmpEDCAutoRmv="" then
		tmpEDCAutoRmv="0"
	end if
	call opendb()
	query="UPDATE pcEDCSettings SET pcES_AutoRmvLogs=" & tmpEDCAutoRmv & ",pcES_AutoRefill=" & tmpEDCAutoFill & ",pcES_TriggerAmount=" & tmpEDCTrigger & ",pcES_FillAmount=" & tmpEDCFillAmount & ",pcES_LogTrans=" & tmpEDCLog & ",pcES_TestMode=" & tmpEDCTest & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	msg="Your settings were updated successfully!"
	msgType=1
	
	call GetEDCSettings()
END IF

IF msg="" then
	If (EDCReg="1") AND (EDCUserID>"0") then
		FuncResult=AutoRefill()
		if FuncResult="1" then
			msg=EDC_SuccessMsg
			msgType=1
		end if
	End If
END IF

%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<%IF EDCReg<>"1" then%>
	<script>
	function isDigit(s)
	{
		var test=""+s;
		if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
			return(true) ;
		}
		return(false);
	}
	
	function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k < test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}
	function checkFormA(tmpForm)
	{
		if (document.getElementById("signedup1").checked)
		{
		if (tmpForm.edcaccountid.value=="")
		{
			alert("Please enter a value for 'Account ID' field");
			tmpForm.edcaccountid.focus();
			return(false);
		}
		if (allDigit(tmpForm.edcaccountid.value) == false)
		{
			alert("Please enter a number for 'Account ID' field");
			tmpForm.edcaccountid.focus();
			return(false);
		}
		if (tmpForm.edcpassp.value=="")
		{
			alert("Please enter a value for 'Pass Phrase' field");
			tmpForm.edcpassp.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value=="")
		{
			alert("Please enter a value for 'Pass Phrase Confirmation' field");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value!=tmpForm.edcpassp.value)
		{
			alert("'Pass Phrase' and 'Pass Phrase Confirmation' values are not the same");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		}
		return(true);
	}
	</script>
	<form name="formA" method="post" action="<%=pcPageName%>?action=checksignup" onsubmit="javascript: return(checkFormA(this));">
	<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">Did you already register this or another ProductCart store with Endicia? If so, please note that you will be in TEST Mode when you re enable Endicia!</th>
	</tr>
	<tr>
		<td><input type="radio" name="signedup" id="signedup1" value="1" onclick="if (this.checked) {document.getElementById('enteracc').style.display=''} else {document.getElementById('enteracc').style.display='none'}" class="clearBorder"></td>
		<td width="95%">Yes</td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table class="pcCPcontent" id="enteracc" style="display:none;">
				<tr>
					<th colspan="2">Current account information</th>
				</tr>
				<tr valign="top">
					<td>Account ID:</td>
					<td>
						<input type="text" name="edcaccountid" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
					</td>
				</tr>
				</tr>
					<td>Pass Phrase:</td>
					<td>
						<input type="password" name="edcpassp" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
					</td>
					</tr>
					<tr valign="top">
						<td nowrap>Confirm Pass Phrase:</td>
						<td>
							<input type="password" name="edcpassp1" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
						</td>
					</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td><input type="radio" name="signedup" value="0" checked onclick="if (this.checked) {document.getElementById('enteracc').style.display='none'} else {document.getElementById('enteracc').style.display=''}" class="clearBorder"></td>
		<td width="95%">No</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">
			<input type="submit" name="submitFormA" value=" Submit " class="submit2">
		</td>
	</tr>
	</table>
<%ELSE%>
<table class="pcCPcontent">
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2"><%If (EDCReg="1") AND ((EDCUserID="0") OR (EDCUserID="")) then%>Initializing your account<%else%>Change your Pass Phrase<%end if%></th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%If (EDCReg="1") AND ((EDCUserID="0") OR (EDCUserID="")) then%>
<tr>
	<td colspan="2">
		After signing up for an account, you will receive an email confirmation with the Account Number (Account ID). Please use it to assign a <u>new Pass Phrase</u> to the account to complete the signing up process. Then you will be able to activate your account and begin printing postage for your USPS shipments.
	</td>
</tr>
<%End if%>
<tr>
	<td colspan="2" class="pcCPspacer">
	<script>
	function isDigit(s)
	{
		var test=""+s;
		if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
			return(true) ;
		}
		return(false);
	}
	
	function allDigit(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigit(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

	function checkForm(tmpForm)
	{
		if (tmpForm.edcaccountid.value=="")
		{
			alert("Please enter a value for 'Account ID' field");
			tmpForm.edcaccountid.focus();
			return(false);
		}
		if (allDigit(tmpForm.edcaccountid.value) == false)
		{
			alert("Please enter a number for 'Account ID' field");
			tmpForm.edcaccountid.focus();
			return(false);
		}
		if (tmpForm.edccurrentpassp.value=="")
		{
			alert("Please enter a value for 'Current Pass Phrase' field");
			tmpForm.edccurrentpassp.focus();
			return(false);
		}
		if (tmpForm.edcpassp.value=="")
		{
			alert("Please enter a value for 'New Pass Phrase' field");
			tmpForm.edcpassp.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value=="")
		{
			alert("Please enter a value for 'New Pass Phrase Confirmation' field");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		if (tmpForm.edcpassp1.value!=tmpForm.edcpassp.value)
		{
			alert("'New Pass Phrase' and 'New Pass Phrase Confirmation' values are not the same");
			tmpForm.edcpassp1.focus();
			return(false);
		}
		return(true);
	}
	</script>
	</td>
</tr>

<form name="form1" method="post" action="<%=pcPageName%>?action=changepassp" onsubmit="javascript:	if (checkForm(this)) {pcf_Open_EndiciaPop();return(true);} else {return(false);}">
<tr valign="top">
	<td>Account ID:</td>
	<td>
		<input type="text" name="edcaccountid" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
		<%If (EDCReg="1") AND ((EDCUserID="0") OR (EDCUserID="")) then%>
			<br>
			<i>You can find it in the email confirmation that you received after signing up for your account</i>
		<%End if%>
	</td>
</tr>
<tr valign="top">
	<td>Current Pass Phrase:</td>
	<td>
		<input type="password" name="edccurrentpassp" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
		<%If (EDCReg="1") AND ((EDCUserID="0") OR (EDCUserID="")) then%>
			<br>
			<i>The Pass Phrase that you used when signing up the account</i>
		<%End if%>
	</td>
</tr>
<tr valign="top">
	<td>New Pass Phrase:</td>
	<td>
		<input type="password" name="edcpassp" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr valign="top">
	<td nowrap>Confirm New Pass Phrase:</td>
	<td>
		<input type="password" name="edcpassp1" size="30" value=""> <img src="images/sample/pc_icon_required.gif" border="0">
	</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" name="submit1" value=" Update Pass Phrase" class="submit2">&nbsp;<input type="button" name="Another" value=" Use another Endicia Account " onclick="location='EDC_manage.asp?action=reenable';" class="ibtnGrey"></td>
</tr>
</form>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%
If (EDCReg="1") AND (EDCUserID>"0") then

	If EDCAutoRmv>"0" then
		dim dtTodaysDate
		dtTodaysDate=Date()-Clng(EDCAutoRmv)
		if SQL_Format="1" then
			dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
		else
			dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
		end if
		if scDB="SQL" then
			query="DELETE FROM pcEDCLogs WHERE pcET_ID IN (SELECT DISTINCT pcET_ID FROM pcEDCTrans WHERE pcET_TransDate<='" & dtTodaysDate & "');"
		else
			query="DELETE FROM pcEDCLogs WHERE pcET_ID IN (SELECT DISTINCT pcET_ID FROM pcEDCTrans WHERE pcET_TransDate<=#" & dtTodaysDate & "#);"
		end if
		call opendb()
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
		if scDB="SQL" then
			query="DELETE FROM pcEDCTrans WHERE pcET_TransDate<='" & dtTodaysDate & "';"
		else
			query="DELETE FROM pcEDCTrans WHERE pcET_TransDate<=#" & dtTodaysDate & "#;"
		end if
		set rsQ=connTemp.execute(query)
		set rsQ=nothing
	end if

    tmpEDC=GetAccountStatus()
    if tmpEDC="1" then
    %>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <th colspan="2">Refill your account</th>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td><font size="3"><b>Current Balance:</b></td>
            <td><font size="3"><b><%=scCurSign & money(EDCABalance)%></b></font></td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <form name="form2">
        <tr valign="top">
            <td>Refill Amount:</td>
            <td><%=scCurSign%><input name="edcrefill" id="edcrefill" type="text" size="10" value="0"> <img src="images/sample/pc_icon_required.gif" border="0"><br>
            <i>The Refill Amount must be between $10.00 and $99,999.99 US dollars, rounded to the nearest cent.</i></td>
        </tr>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td>&nbsp;</td>
            <td>
            <script>
            function isDigit1(s)
            {
                var test=""+s;
                if(test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
                {
                    return(true) ;
                }
                return(false);
            }
            
            function allDigit1(s)
            {
                var test=""+s ;
                for (var k=0; k <test.length; k++)
                {
                    var c=test.substring(k,k+1);
                    if (isDigit1(c)==false)
                    {
                        return (false);
                    }
                }
                return (true);
            }
            
            function checkNumber(tmpField)
            {
                if (tmpField.value=="")
                {
                    alert("Please enter a value for 'Amount' field");
                    tmpField.focus();
                    return(false);
                }
                if (allDigit1(tmpField.value) == false)
                {
                    alert("Please enter a numeric value for 'Amount' field");
                    tmpField.focus();
                    return(false);
                }
                if ((parseFloat(tmpField.value)<10) || (parseFloat(tmpField.value)>99999.99))
                {
                    alert("Please enter a numeric value in the 'Amount' field. The amount must be between: 10.00 - 99999.99");
                    tmpField.focus();
                    return(false);
                }
                return(true);
            }
            </script>
            <input type="button" name="refillbtn" class="submit2" value=" Refill Account " onclick="if (checkNumber(document.form2.edcrefill)) location='<%=pcPageName%>?edcaction=refill&amount='+document.form2.edcrefill.value;"></td>
        </tr>
        </form>
        <tr>
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
    <%
	end if
	%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Settings</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer">
	<script>	
	function checkForm3(tmpForm)
	{
		if (tmpForm.edcautofill.value=="1")
		{
			if (tmpForm.edctrigger.value=="")
			{
			alert("Please enter a number for this field");
			tmpForm.edctrigger.focus();
			return(false);
			}
			if (tmpForm.edcfillamount.value=="")
			{
			alert("Please enter a number for this field");
			tmpForm.edcfillamount.focus();
			return(false);
			}
		
			if (allDigit1(tmpForm.edctrigger.value) == false)
			{
			alert("Please enter a number for this field");
			tmpForm.edctrigger.focus();
			return(false);
			}
			if (allDigit1(tmpForm.edcfillamount.value) == false)
			{
			alert("Please enter a number for this field");
			tmpForm.edcfillamount.focus();
			return(false);
			}
		}
		return(true);
	}
	</script>
	</td>
</tr>
<form name="form3" action="<%=pcPageName%>?action=updatesettings" method="post" onsubmit="javascript: if (checkForm3(this)) {pcf_Open_EndiciaPop();return(true);} else {return(false);}">
<tr>
	<td colspan="2"><input type="radio" name="edctest" value="0" <%if EDCTestMode="0" then%>checked<%end if%> class="clearBorder"> Using "LIVE" Mode - This setting only affects the printing of USPS shipping labels</u> in the Shipping Wizard.</td>
</tr>
<tr>
	<td colspan="2"><input type="radio" name="edctest" value="1" <%if EDCTestMode="1" then%>checked<%end if%> class="clearBorder"> Using "TEST" Mode - "TEST" mode only affects the printing of USPS shipping labels</u> in the Shipping Wizard. While in "TEST" mode all labels will be printed as "SAMPLE" labels and cannot be used to ship packages. Use "TEST" mode to ensure that labels are being correctly generated.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><input type="checkbox" name="edclog" value="1" <%if EDCLogTrans="1" then%>checked<%end if%> class="clearBorder"> Log Transactions between your store and the Endicia Postage Label Server</td>
</tr>
<tr>
	<td colspan="2">Automatically clear logs older than <input type="text" name="edcautormv" size="5" value="<%=EDCAutoRmv%>"> days</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><input type="checkbox" name="edcautofill" value="1" <%if EDCAutoRefill="1" then%>checked<%end if%> class="clearBorder"> <u><b>Auto refill your account</b></u></td>
</tr>
<tr>
	<td>When balance is under:</td>
	<td><%=scCurSign%><input type="text" name="edctrigger" size="10" value="<%=EDCTriggerAmount%>"></td>
</tr>
<tr>
	<td>Amount to refill:</td>
	<td><%=scCurSign%><input type="text" name="edcfillamount" size="10" value="<%=EDCFillAmount%>"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" name="submit2" value=" Update Settings " class="submit2"></td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</form>
<%if EDCLogTrans="1" then
query="SELECT pcELog_ID FROM pcEDCLogs;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<th colspan="2">Transactions Logs</th>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2"><a href="EDC_Trans.asp">Click here</a> to view transaction logs between your store and the Endicia Postage Label Server</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%end if
set rsQ=nothing
end if%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="2">Please note that you can use your Account ID and Web Password to access your account at <a href="https://www.endicia.com/Account/LogIn/" target="_blank">www.endicia.com</a> and then change your payment option, view your transactions (purchase and print) and also issue refunds, etc.</td>
</tr>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
<%END IF%>
<tr>
	<td colspan="2" class="pcCPspacer"></td>
</tr>
</table>
<%END IF 'Not Signed Up%>
<%call closedb()%>
<%Response.write(pcf_ModalWindow("Connecting to Endicia Label Server... ","EndiciaPop", 300))%>
<!--#include file="AdminFooter.asp"-->