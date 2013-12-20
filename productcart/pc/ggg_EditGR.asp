<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/openDb.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

dim conntemp

call openDb()

pIdCustomer=session("idCustomer")
gIDEvent=getUserInput(request("IDEvent"),0)

IF request("delregistry")<>"" then
	query="delete from pcEvents where pcEv_IDEvent=" & gIDEvent & " and pcEv_IDCustomer=" & pIDCustomer
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	query="delete from pcEvProducts where pcEP_IDEvent=" & gIDEvent
	set rstemp=server.CreateObject("ADODB.RecordSet")
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	set rstemp=nothing
	call closedb()
	response.redirect "ggg_manageGRs.asp"
ELSE
	if (request("action")="update") and (request("rewrite")="0") then
		getype=getUserInput(request("etype"),0)
		gename=getUserInput(request("ename"),0)
		'getype=replace(getype,"'","''")
		'gename=replace(gename,"'","''")
		gedate=getUserInput(request("edate"),0)
		if gedate="" then
			gedate="01/01/1900"
		end if
		gedelivery=getUserInput(request("edelivery"),0)
		if gedelivery="" then
			gedelivery="0"
		end if
		gemyaddr=getUserInput(request("emyaddr"),0)
		if gemyaddr="" then
			gemyaddr="0"
		end if
		gehide=getUserInput(request("ehide"),0)
		if gehide="" then
			gehide="0"
		end if
		geHideAddress=getUserInput(request("eHideAddress"),0)
		if geHideAddress="" then
			geHideAddress="0"
		end if
		genotify=getUserInput(request("enotify"),0)
		if genotify="" then
			genotify="0"
		end if
		geincgc=getUserInput(request("eincgc"),0)
		if geincgc="" then
			geincgc="0"
		end if	
		geactive=getUserInput(request("eactive"),0)
		if geactive="" then
			geactive="0"
		end if
		
		Do while mytest=0
			myTest=0
			Tn1=""
			For w=1 to 16
				Randomize
				myC=Fix(3*Rnd)
				Select Case myC
				Case 0: 
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
				Case 1: 
					Randomize
					Tn1=Tn1 & Cstr(Fix(10*Rnd))
				Case 2: 
					Randomize
					Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
				End Select
			Next

			query="select pcEv_IDEvent from pcEvents where pcEv_Code='" & Tn1 & "'"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	
			if rstemp.eof then
				myTest=1
			end if
	
		Loop
	
		geCode=Tn1
		
		if SQL_Format="1" then
			geDate=(day(geDate)&"/"&month(geDate)&"/"&year(geDate))
		else
			gExpDate=(month(geDate)&"/"&day(geDate)&"/"&year(geDate))
		end if
		if scDB="SQL" then
			query="Update pcEvents set pcEv_Type='" & getype & "',pcEv_Name='" & geName & "',pcEv_Date='" & geDate & "',pcEv_Delivery=" & gedelivery & ",pcEv_MyAddr=" & gemyaddr & ",pcEv_Hide=" & gehide & ",pcEv_Notify=" & genotify & ",pcEv_IncGcs=" & geIncGc & ",pcEv_Active=" & geActive & ", pcEv_HideAddress=" & geHideAddress & " where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
		else
			query="Update pcEvents set pcEv_Type='" & getype & "',pcEv_Name='" & geName & "',pcEv_Date=#" & geDate & "#,pcEv_Delivery=" & gedelivery & ",pcEv_MyAddr=" & gemyaddr & ",pcEv_Hide=" & gehide & ",pcEv_Notify=" & genotify & ",pcEv_IncGcs=" & geIncGc & ",pcEv_Active=" & geActive & ", pcEv_HideAddress=" & geHideAddress &" where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
		end if
		set rstemp=server.CreateObject("ADODB.RecordSet")
		set rstemp=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if geincgc="1" then
	
			query="select IDProduct from Products where pcprod_GC=1"
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		
			do while not rstemp.eof
				IDProduct=rstemp("IDProduct")
				query="select pcEP_IDProduct from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_IDProduct=" & IDProduct & " and pcEP_GC=1"
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs1=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if rs1.eof then
					query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_GC) values (" & gIDEvent & "," & IDProduct & ",1)"
					set rs1=server.CreateObject("ADODB.RecordSet")
					set rs1=connTemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs1=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					set rs1=nothing
				end if
				set rs1=nothing
				rstemp.MoveNext
			loop
			set rstemp=nothing
		else
			query="delete from pcEvProducts where pcEP_Gc=1 and pcEP_HQty=0 and pcEP_IDEvent=" & gIDEvent
			set rs1=server.CreateObject("ADODB.RecordSet")
			set rs1=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs1=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rs1=nothing
		end if
		msg=dictLanguage.Item(Session("language")&"_instGR_15")
	end if

END IF 'not delete

IF request("rewrite")="1" then

	getype=getUserInput(request("etype"),0)
	gename=getUserInput(request("ename"),0)
	getype=replace(getype,"''","'")
	gename=replace(gename,"''","'")
	gedate=getUserInput(request("edate"),0)
	if gedate="" then
		gedate="01/01/1900"
	end if
	gedelivery=getUserInput(request("edelivery"),0)
	if gedelivery="" then
		gedelivery="0"
	end if
	gemyaddr=getUserInput(request("emyaddr"),0)
	if gemyaddr="" then
		gemyaddr="0"
	end if
	gehide=getUserInput(request("ehide"),0)
	if gehide="" then
		gehide="0"
	end if
	geHideAddress=getUserInput(request("eHideAddress"),0)
	if geHideAddress="" then
		geHideAddress="0"
	end if
	genotify=getUserInput(request("enotify"),0)
	if genotify="" then
		genotify="0"
	end if
	geincgc=getUserInput(request("eincgc"),0)
	if geincgc="" then
		geincgc="0"
	end if	
	geactive=getUserInput(request("eactive"),0)
	if geactive="" then
		geactive="0"
	end if

ELSE

	query="select pcEv_Type, pcEv_Name,pcEv_Date, pcEv_Delivery, pcEv_MyAddr, pcEv_Hide, pcEv_Notify, pcEv_IncGcs, pcEv_Active, pcEv_HideAddress from pcEvents where pcEv_IDEvent=" & gIDEvent & " and pcEv_IDCustomer=" & pIDCustomer
	set rstemp=connTemp.execute(query)
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "ggg_manageGRs.asp"
	end if

	getype=rstemp("pcEv_Type")
	gename=rstemp("pcEv_Name")
	gedate=rstemp("pcEv_Date")
	if year(gedate)="1900" then
		gedate=""
	end if
	gedelivery=rstemp("pcEv_Delivery")
	if gedelivery<>"" then
	else
		gedelivery="0"
	end if
	gemyaddr=rstemp("pcEv_MyAddr")
	if gemyaddr<>"" then
	else
		gemyaddr="0"
	end if
	gehide=rstemp("pcEv_Hide")
	if gehide<>"" then
	else
		gehide="0"
	end if
	geHideAddress=rstemp("pcEv_HideAddress")
	if geHideAddress<>"" then
	else
		geHideAddress="0"
	end if
	genotify=rstemp("pcEv_Notify")
	if genotify<>"" then
	else
		genotify="0"
	end if
	geincgc=rstemp("pcEv_IncGcs")
	if geincgc<>"" then
	else
		geincgc="0"
	end if	
	geactive=rstemp("pcEv_Active")
	if geactive<>"" then
	else
		geactive="0"
	end if
	
	if gedate<>"" then
		if scDateFrmt="DD/MM/YY" then
			gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
		else
			gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
		end if
	end if

	set rstemp=nothing
END IF
	
	gShowDel=1
	
	query="select sum(pcEP_HQty) as gHQty from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " group by pcEP_IDEvent"
    set rs1=connTemp.execute(query)
    if err.number<>0 then
		call LogErrorToDatabase()
		set rs1=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
    
    if not rs1.eof then
    	gHQty=rs1("gHQty")
    	if (gHQty<>"") then
	    	if Clng(gHQty)>0 then
		    	gShowDel=0
	    	end if
    	end if
    end if
    
    GCDel=1
    
	query="select sum(pcEP_HQty) as gHQty from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_GC=1 group by pcEP_IDEvent"
	set rs1=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs1=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
    
    if not rs1.eof then
    	gHQty=rs1("gHQty")
    	if (gHQty<>"") then
	    	if Clng(gHQty)>0 then
		    	GCDel=0
	    	end if
    	end if
    end if
%>
<!--#include file="header.asp"-->
<script>
function win(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=600,height=550')
	myFloater.location.href=fileName;
	checkwin();
	}
function checkwin()
{

if (myFloater.closed)
{
document.Form1.submit();
}
else
{
setTimeout('checkwin()',500);
}
}

</script>
<script language="JavaScript">
<!--


function check_date(field){
var checkstr = "0123456789";
var DateField = field;
var Datevalue = "";
var DateTemp = "";
var seperator = "/";
var day;
var month;
var year;
var leap = 0;
var err = 0;
var i;
   err = 0;
   DateValue = DateField.value;
   /* Delete all chars except 0..9 */
   for (i = 0; i < DateValue.length; i++) {
	  if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
	     DateTemp = DateTemp + DateValue.substr(i,1);
	  }
	  else
	  {
	  if (DateTemp.length == 1)
		{
    	  DateTemp = "0" + DateTemp
		}
	  else
	  {
	  	if (DateTemp.length == 3)
	  	{
	  	DateTemp = DateTemp.substr(0,2) + '0' + DateTemp.substr(2,1);
	  	}
	  }
	 }
   }
   DateValue = DateTemp;
   /* Always change date to 8 digits - string*/
   /* if year is entered as 2-digit / always assume 20xx */
   if (DateValue.length == 6) {
      DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
   if (DateValue.length != 8) {
      return(false);}
   /* year is wrong if year = 0000 */
   year = DateValue.substr(4,4);
   if (year == 0) {
      err = 20;
   }
   /* Validation of month*/
   <%if scDateFrmt="DD/MM/YY" then%>
   month = DateValue.substr(2,2);
   <%else%>
   month = DateValue.substr(0,2);
   <%end if%>
   if ((month < 1) || (month > 12)) {
      err = 21;
   }
   /* Validation of day*/
   <%if scDateFrmt="DD/MM/YY" then%>
   day = DateValue.substr(0,2);
   <%else%>
   day = DateValue.substr(2,2);
   <%end if%>
   if (day < 1) {
     err = 22;
   }
   /* Validation leap-year / february / day */
   if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
      leap = 1;
   }
   if ((month == 2) && (leap == 1) && (day > 29)) {
      err = 23;
   }
   if ((month == 2) && (leap != 1) && (day > 28)) {
      err = 24;
   }
   /* Validation of other months */
   if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
      err = 25;
   }
   if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
      err = 26;
   }
   /* if 00 ist entered, no error, deleting the entry */
   if ((day == 0) && (month == 0) && (year == 00)) {
      err = 0; day = ""; month = ""; year = ""; seperator = "";
   }
   if ((err == 0) && (day != "") && (month != "") && (year != "") && (seperator != ""))
   {
		var EDate=new Date(year, month-1, day);
		var NDate=new Date();
		if (EDate<NDate) err=1;
   }
   /* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
   if (err == 0) {
	<%if scDateFrmt="DD/MM/YY" then%>
	DateField.value = day + seperator + month + seperator + year;
    <%else%>
	DateField.value = month + seperator + day + seperator + year;   
    <%end if%>
	return(true);
   }
   /* Error-message if err != 0 */
   else {
	return(false);   
   }
}
	
function Form1_Validator(theForm)
{
	if (theForm.ename.value == "")
  	{
			alert("Please enter a value for this field.");
		    theForm.ename.focus();
		    return (false);
	}

	if (theForm.edate.value == "")
  	{
			alert("Please enter a valid date for this field.");
		    theForm.edate.focus();
		    return (false);
	}
	
	if (check_date(theForm.edate) == false)
  	{
		alert("Please enter a valid date for this field.");
	    theForm.edate.focus();
	    return (false);
	}
	
	if (theForm.subdel.value == "1")
  	{
    return (confirm('<%response.write dictLanguage.Item(Session("language")&"_GRDetails_14")%>'));
  	}
	
return (true);
}
//-->
</script>
<div id="pcMain">
<form method="post" name="Form1" action="ggg_EditGR.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcMainTable">
	<tr> 
		<td colspan="2"><h1><%response.write dictLanguage.Item(Session("language")&"_instGR_1a")%></h1></td>
	</tr>
	<% If msg<>"" then %>
	<tr>
		<td colspan="2"><div class="pcErrorMessage"><%=msg%></div></td>
	</tr>
	<% end if %>
	<tr>
		<td width="25%"> 
			<%response.write dictLanguage.Item(Session("language")&"_instGR_2")%>
		</td>
		<td width="75%"> 
			<input type=text name="etype" size="30" value="<%=getype%>">
		</td>
	</tr>
	<tr> 
		<td width="25%"> 
			<%response.write dictLanguage.Item(Session("language")&"_instGR_3")%>
		</td>
		<td width="75%"> 
			<input type=text name="ename" size="30" value="<%=gename%>">
		</td>
	</tr>
	<tr> 
		<td width="25%"> 
			<%response.write dictLanguage.Item(Session("language")&"_instGR_4")%>
		</td>
		<td width="75%"> 
			<input type=text name="edate" size="30" value="<%=gedate%>"> (<i><%response.write dictLanguage.Item(Session("language")&"_instGR_4a")%>
			<%if scDateFrmt="DD/MM/YY" then%>DD/MM/YY<%else%>MM/DD/YY<%end if%></i>)
		</td>
	</tr>
	<tr> 
		<td colspan="2"> 
			<%response.write dictLanguage.Item(Session("language")&"_instGR_5")%>
		</td>
	</tr>
	<tr> 
		<td width="25%" align="right" valign="top"><input type=radio name="edelivery" value="1" class="clearBorder" <%if gedelivery="1" then%>checked<%end if%>></td>
		<td width="75%"> 
			<%response.write dictLanguage.Item(Session("language")&"_instGR_6")%>
			&nbsp;<select name="emyaddr">
			<%
	        myTest=0

			query="SELECT address,city,state,statecode,zip,countrycode,shippingAddress, shippingCity, shippingState, shippingStateCode, shippingZip, shippingCountryCode, shippingCompany, shippingAddress2 FROM customers WHERE idCustomer=" &session("idCustomer")
			set rstemp=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		
			pshippingAddress=rstemp("shippingAddress")
			pshippingZip=rstemp("shippingZip")
			pshippingState=rstemp("shippingState")
			pshippingStateCode=rstemp("shippingStateCode") 
			pShippingCity=rstemp("shippingCity")
			pshippingCountryCode=rstemp("shippingCountryCode")
			pshippingCompany=rstemp("shippingCompany")
			pshippingAddress2=rstemp("shippingAddress2")
			
			myTest=1
			session("paddress")=ucase(rstemp("address"))
			session("pcity")=ucase(rstemp("city"))
			session("pstate")=ucase(rstemp("state") & rstemp("statecode"))
			
			session("pshipadd")=pshippingAddress
			session("pshipZip")=pshippingZip
			session("pshipState")=pshippingState
			session("pshipStateCode")=pshippingStateCode 
			session("pshipCity")=pShippingCity
			session("pshipCountryCode")=pshippingCountryCode
			session("pshipCom")=pshippingCompany
			session("pshipadd2")=pshippingAddress2
			%>
			<option value="0" <%if gemyaddr="0" then%>selected<%end if%>><%=dictLanguage.Item(Session("language")&"_CustSAmanage_10")%></option>
			<%
			query="SELECT idRecipient, recipient_NickName,recipient_Address,recipient_City,recipient_State,recipient_StateCode FROM recipients WHERE idCustomer=" &session("idCustomer")
			set rstemp=conntemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

	        do while not rstemp.eof
		        myTest=1
		        IDre=rstemp("idRecipient")
		        reFullName=trim(rstemp("recipient_NickName"))
		        reShipAddr=ucase(rstemp("recipient_Address"))
		        reShipCity=ucase(rstemp("recipient_City"))
		        reShipState=ucase(rstemp("recipient_State") & rstemp("recipient_StateCode"))
		        myTest1=0
		        if (reShipAddr=session("pAddress")) and (reShipState=session("pState")) and (reShipCity=session("pCity")) and (reFullName="") then 
					myTest1=1
				end if
		        if (reShipAddr=ucase(session("pShipAdd"))) and (reShipState=Ucase(session("pShipState")&session("pShipStateCode")) ) and (reShipCity=ucase(session("pShipCity"))) and (reFullName="") then 
					myTest1=1
				end if			
		        if MyTest1=0 then
			        if trim(reFullName)<>"" then
			        else
				        reFullName="No shipping name specified"
			        end if%>
					<option value="<%=IDre%>" <%if clng(gemyaddr)=clng(IDre) then%>selected<%end if%>><%=reFullName%></option>
			        <%
		        end if
		        rstemp.movenext
			loop%>
			</select>
			<br>
			<a href="javascript:win('CustAddShipPop.asp');"><%response.write dictLanguage.Item(Session("language")&"_instGR_8")%></a>
		</td>
	</tr>
	<tr> 
		<td width="25%" align="right" valign="top"><input type=radio name="edelivery" value="0" class="clearBorder" <%if gedelivery<>"1" then%>checked<%end if%>></td>
		<td width="75%"><%response.write dictLanguage.Item(Session("language")&"_instGR_9")%></td>
	</tr>
	<tr> 
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr> 
		<td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_instGR_5b")%></td>
	</tr>
	<tr> 
		<td width="25%" align="right" valign="top">
			<input type=checkbox name="ehide" value="1" class="clearBorder" <%if gehide="1" then%>checked<%end if%>>
		</td>
		<td width="75%">
			<p>
			<%response.write dictLanguage.Item(Session("language")&"_instGR_10")%>
			<br>
			<i><%response.write dictLanguage.Item(Session("language")&"_instGR_10a")%></i>
			</p>
		</td>
	</tr>
	<tr> 
		<td width="25%" align="right" valign="top">
			<input type=checkbox name="eHideAddress" value="1" class="clearBorder" <%if geHideAddress="1" then%>checked<%end if%>>
		</td>
		<td width="75%"><%response.write dictLanguage.Item(Session("language")&"_instGR_17")%></td>
	</tr>
	<tr> 
		<td width="25%" align="right" valign="top">
			<input type=checkbox name="enotify" value="1" class="clearBorder" <%if genotify="1" then%>checked<%end if%>>
		</td>
		<td width="75%"><%response.write dictLanguage.Item(Session("language")&"_instGR_11")%></td>
	</tr>
	<%if GCDel=0 then%>
		<input type=hidden name="eincgc" value="<%=geincgc%>">
	<%else
	if GCDel=1 then%>
	<tr> 
		<td width="25%" align="right" valign="top">
			<input type=checkbox name="eincgc" value="1" class="clearBorder" <%if geincgc="1" then%>checked<%end if%>>
		</td>
		<td width="75%"><%response.write dictLanguage.Item(Session("language")&"_instGR_12")%></td>
	</tr>
	<%end if
	end if%>
	<tr>
		<td width="25%" valign="top"> 
			<input type=checkbox name="eactive" value="1" class="clearBorder" <%if geactive="1" then%>checked<%end if%> style="float: right">
		</td> 
		<td width="75%"><%response.write dictLanguage.Item(Session("language")&"_instGR_13")%></td>
	</tr>
	<tr> 
		<td colspan="2"> 
			<p>
			<br><br>
			<a href="javascript:location='ggg_manageGRs.asp';"><img src="<%=rslayout("back")%>" border=0></a>&nbsp;
			<input type="image" id="submit" name="submit" value="<%response.write dictLanguage.Item(Session("language")&"_instGR_12")%>" src="<%=RSlayout("UpdRegistry")%>" border="0" onclick="document.Form1.subdel.value='0';document.Form1.rewrite.value='0';">
			<%if gShowDel=1 then%>
				&nbsp;<input type="image" id="submit" name="delreg" value="<%response.write dictLanguage.Item(Session("language")&"_instGR_16")%>" onclick="document.Form1.subdel.value='1'; document.Form1.delregistry.value='ok';" src="<%=RSlayout("DelRegistry")%>" border="0">
			<%end if%>
			<input type=hidden name="IDEvent" value="<%=gIDEvent%>">
			<input type=hidden name="subdel" value="0">
			<input type=hidden name="delregistry" value="">
			<input type=hidden name="rewrite" value="1">
			<br>
			</p>
		</td>
	</tr>
</table>
</form>
</div>
<%call closedb()%><!--#include file="footer.asp"-->