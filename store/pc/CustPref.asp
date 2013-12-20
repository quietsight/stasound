<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/bto_language.asp"-->  
<!--#include file="../includes/rewards_language.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"-->
<% dim conntemp, query, rs

call opendb()

'MAILUP-S

	tmp_setup=0
	pcMailUpSett_APIUser=""
	pcMailUpSett_APIPassword=""
	pcMailUpSett_URL=""

	query="SELECT pcMailUpSett_APIUser,pcMailUpSett_APIPassword,pcMailUpSett_URL,pcMailUpSett_AutoReg,pcMailUpSett_RegSuccess FROM pcMailUpSettings;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		pcMailUpSett_APIUser=rs("pcMailUpSett_APIUser")
		session("SF_MU_APIUser")=pcMailUpSett_APIUser
		pcMailUpSett_APIPassword=enDeCrypt(rs("pcMailUpSett_APIPassword"), scCrypPass)
		session("SF_MU_APIPassword")=pcMailUpSett_APIPassword
		pcMailUpSett_URL=rs("pcMailUpSett_URL")
		session("SF_MU_URL")=pcMailUpSett_URL
		tmp_Auto=rs("pcMailUpSett_AutoReg")
		if IsNull(tmp_Auto) or tmp_Auto="" then
			tmp_Auto=0
		end if
		session("SF_MU_Auto")=tmp_Auto
		tmp_setup=rs("pcMailUpSett_RegSuccess")
		if IsNull(tmp_setup) or tmp_setup="" then
			tmp_setup=0
		end if
		session("SF_MU_Setup")=tmp_setup
	end if
	set rs=nothing
'MAILUP-E
	
if session("customerCategory")<>0 then
	query="SELECT pcCC_Name, pcCC_Description FROM pcCustomerCategories WHERE idCustomerCategory="&session("customerCategory")&";"
	SET rs=Server.CreateObject("ADODB.RecordSet")
	SET rs=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	strpcCC_Name=rs("pcCC_Name")
	strpcCC_Description=rs("pcCC_Description")
	SET rs=nothing
end if

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

' START - Retrieve customer name
if session("pcStrCustName") = "" OR session("pcStrCustEmail") = "" then
	pcIntCustomerId = session("idCustomer")
	if not validNum(pcIntCustomerId) then
		session("idCustomer") = Cdbl(0)
		response.Redirect("default.asp")
	end if	
	query = "SELECT name, lastName, email FROM customers WHERE idCustomer = " & pcIntCustomerId
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	pcStrCustName = rs("name") & " " & rs("lastName")
	session("pcStrCustName") = pcStrCustName
	pcStrCustEmail = rs("email")
	session("pcStrCustEmail") = pcStrCustEmail
	set rs = nothing	
	pEmail = pcStrCustEmail
else
	pcIntCustomerId = session("idCustomer")
	if not validNum(pcIntCustomerId) then
		session("idCustomer") = Cdbl(0)
		response.Redirect("default.asp")
	end if	
	query = "SELECT email FROM customers WHERE idCustomer = " & pcIntCustomerId
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs = conntemp.execute(query)
	pEmail = rs("email")
	set rs = nothing
end if
' END - Retrieve customer name

call closedb()
%>
<!--#include file="header.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<div id="pcMain">
	<table class="pcMainTable">
		<tr> 
			<td>
			<%If request.querystring("msg")<>"" then %>
				<div class="pcErrorMessage"><%response.write server.HTMLEncode(request.querystring("msg"))%></div>
			<%end if%>
			<%If request.querystring("mode")="new" then %>
				<div class="pcSuccessMessage"><%response.write dictLanguage.Item(Session("language")&"_RegThankyou_1")%></div>
			<%end if%>
			<h1><%response.write(session("pcStrCustName") & " - " & dictLanguage.Item(Session("language")&"_CustPref_1"))%></h1>
			<%if session("customerType")="1" then%>
				<p><%response.write dictLanguage.Item(Session("language")&"_CustPref_6")%></p>
			<%else%>
				<p><%response.write dictLanguage.Item(Session("language")&"_CustPref_7")%></p>
			<%end if%>
			<% if session("customerCategory")<>0 then%>
				<p><%response.write dictLanguage.Item(Session("language")&"_CustPref_15") & strpcCC_Name %></p>
			<%end if%>
			<p>&nbsp;</p>
			<p><%response.write(dictLanguage.Item(Session("language")&"_CustPref_10") & session("pcStrCustName") & "!")%></p>
			<ul>
			<li><a href="default.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_9")%></a></li>
			<%
			'// GGG Add-on start
			pcv_strGoogleOnly=pcf_PaymentTypes("GoogleCheckout")
			pcv_strPayPalExpOnly=pcf_PaymentTypes("PayPalExp")			
			if scDisableGiftRegistry <> "1" AND (pcv_strGoogleOnly=1 OR pcv_strPayPalExpOnly=0) then
			%>  
				<li><a href="ggg_manageGRs.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_13")%></a></li>
			<%
			end if
			'//GGG Add-on end 
			%>
			<li><a href="CustviewPast.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_8")%></a></li>
			<% 
			'// Start Reward Points
			If RewardsActive <> 0 AND session("customerType")<>"1" then %>
				<li><a href="CustRewards.asp"><%response.write dictRewardsLanguage.Item(Session("language")&"_CustPref_11")%><%=RewardsLabel%></a></li>
			<% End If %> 
			<% If RewardsActive <> 0 AND session("customerType")="1" AND RewardsIncludeWholesale=1 then %>
				<li><a href="CustRewards.asp"><%response.write dictRewardsLanguage.Item(Session("language")&"_CustPref_11")%><%=RewardsLabel%></a></li>
			<% End If
			'// End Reward Points %> 
			<li><a href="login.asp?lmode=1" <%if session("SF_MU_Setup")="1" then%>onclick="javascript:pcf_Open_MailUp();"<%end if%>><%response.write dictLanguage.Item(Session("language")&"_CustPref_3")%></a></li>
			<li><a href="CustSAmanage.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_11")%></a></li>
			<% If (scWL=-1) or ((scBTO=1) and (iBTOQuote=1)) then %>
				<li><a href="Custquotesview.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_5")%></a></li>
			<% End If %>
			<li><a href="CustSavedCarts.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_16")%></a></li>
			<%
			'SB S 
			if scSBStatus="1" then
			%>
			<li><a href="sb_CustViewSubs.asp"><%response.write dictLanguage.Item(Session("language")&"_SB_3")%></a></li>
			<%
			end if
			'SB E 
			%>
            <%
			call openDB()
			query="SELECT pcPay_EIG_Vault_ID, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE idCustomer="& Session("idCustomer") &""
			set rs=Server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)		
			if NOT rs.eof then
				%>
                <li><a href="CustviewPayment.asp"><%response.write dictLanguage.Item(Session("language")&"_EIG_10")%></a></li>
                <%
			end if
			set rs=nothing
			call closeDB()
			%>
			<li><a href="contact.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_12")%></a></li>
			<li><a href="CustLO.asp"><%response.write dictLanguage.Item(Session("language")&"_CustPref_4")%></a></li>
		</ul>

		<% '// Account Consolidation %>
		<% call openDB() %>
            <!--#include file="opc_inc_CustConsolidate.asp"-->            
        <%
		'// START - Check Gift Certificate Balance
		'// Check to see if there are active Gift Certificates		
		Dim pcvIntGCExist
		pcvIntGCExist=0
		query="SELECT pcGO_GcCode FROM pcGCOrdered WHERE pcGO_Status = 1;"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			if not rs.eof then
				pcvIntGCExist=1
			end if
		set rs=nothing
		call closeDB()
		
		IF pcvIntGCExist<>0 THEN '// START - There are gift certificates
			pGiftCode = getUserInput(request.Form("pcGCcode"),100)
			IF pGiftCode<>"" THEN
				call openDB()
				query="SELECT products.IDProduct, products.Description, pcGO_GcCode, pcGO_ExpDate, pcGO_Amount, pcGO_Status FROM Products,pcGCOrdered WHERE products.idproduct=pcGCOrdered.pcGO_idproduct AND pcGCOrdered.pcGO_GcCode='"& pGiftCode & "'"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
	
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
	
					IF NOT rs.eof THEN
					
						pcvGiftCertName=rs("Description")
						pcvGiftCertExp=rs("pcGO_ExpDate")
						pcIntGiftCertStatus=rs("pcGO_Status")
						
							if year(pcvGiftCertExp)="1900" then
								pcvGiftCertExp = dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_15")
							else
								if scDateFrmt="DD/MM/YY" then
									pcvGiftCertExp=day(pcvGiftCertExp) & "/" & month(pcvGiftCertExp) & "/" & year(pcvGiftCertExp)
								else
									pcvGiftCertExp=month(pcvGiftCertExp) & "/" & day(pcvGiftCertExp) & "/" & year(pcvGiftCertExp)
								end if
							If datediff("d", Now(), pcvGiftCertExp) <= 0 Then pcIntGiftCertExpired = 1
							end if
							
						pcvGiftCertAmount=rs("pcGO_Amount")
							if pcvGiftCertAmount<0 then pcvGiftCertAmount=0
	
						set rs = nothing
						%>
						<br /><br />
                        <form name="checkGC2" action="" method="" class="pcForms">
						<fieldset>
						<legend><%=pcvGiftCertName%></legend>
                        <%
						if pcIntGiftCertStatus<>0 and pcIntGiftCertExpired<>1 then
							'// Gift Certificate is active
							%><img src="images/pc_checkmark_sm_green.gif" hspace="5" alt="<%=dictLanguage.Item(Session("language")&"_CustPref_21")%>"><%
							response.write dictLanguage.Item(Session("language")&"_CustPref_21")
							else
							'// Gift Certificate is inactive
							%><img src="images/pc_icon_error.gif" hspace="5" alt="<%=dictLanguage.Item(Session("language")&"_CustPref_22")%>"><%
							response.write dictLanguage.Item(Session("language")&"_CustPref_22")
						end if
						%>
                        <br />
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_11")%><strong><%=pGiftCode%></strong><br />
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_16")%><%=pcvGiftCertExp%> <br />
						<%response.write dictLanguage.Item(Session("language")&"_ggg_OrdInvoice_14")%><strong><%=scCurSign & money(pcvGiftCertAmount)%></strong>
						</fieldset>
                        </form>
						<%
					
					ELSE
					
						%>
						<br /><br />
						
						<div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_CustPref_18")%></div>                 	
						
			<%
					END IF '// Retrieving information from the database
				call closeDB()
				END IF '// GC Check form has been submitted
			'// END
			%>
	
			
			<br /><br />
			<form name="checkGC" action="custPref.asp" method="post" class="pcForms">
			<fieldset>
				<legend><%=dictLanguage.Item(Session("language")&"_CustPref_17")%></legend>
					<p><%=dictLanguage.Item(Session("language")&"_CustPref_19")%><input type="text" size="20" name="pcGCcode"></p>
					<p>&nbsp;</p>
					<p><input type="image" id="submit" src="<%=rslayout("submit")%>" name="submitGCcheck" value="<%=dictLanguage.Item(Session("language")&"_CustPref_20")%>"></p>
			</fieldset>
			</form>
            
        <%
		END IF '// END - There are gift certificates
		%>
		<%Response.write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_MailUp_SynNote3"),"MailUp", 300))%>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->