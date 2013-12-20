<% 
'---Start Offline Credit Card---
Function OffLineCCEdit()
	call opendb()
	
	cvv=0
	pcv_idPayment=request.Form("id")
	Cbtob=request.Form("Cbtob")
	if Cbtob = "" then
		Cbtob = "0"
	end if
	if Cbtob="1" then
		Cbtob="-1"
	end if
	if Cbtob="2" then
		Cbtob2 = request("Cbtob2")
		if Cbtob2 = "" then
			Cbtob = "0"
		end if
	end if

	query="UPDATE payTypes SET pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus & ",priceToAdd="& priceToAdd &" , percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"', Cbtob="&Cbtob&" WHERE gwCode= 6"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	query = "DELETE FROM CustCategoryPayTypes WHERE idPayment = "&pcv_idPayment&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	if Cbtob="2" then
		CbtobArray=split(Cbtob2,",")
		for t=lbound(CbtobArray) to ubound(CbtobArray)
			query = "INSERT INTO CustCategoryPayTypes (idCustomerCategory,idPayment) VALUES ("&CbtobArray(t)&" ,"&pcv_idPayment&");"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
		next
	end if
	
	CCTypeM=request.form("CCTypeM")
	if CCTypeM="" then
		CCTypeM=0
	end if
	CCTypeV=request.form("CCTypeV")
	if CCTypeV="" then
		CCTypeV=0
	end if
	CCTypeA=request.form("CCTypeA")
	if CCTypeA="" then
		CCTypeA=0
	end if
	CCTypeD=request.form("CCTypeD")
	if CCTypeD="" then
		CCTypeD=0
	end if
	CCTypeDC=request.form("CCTypeDC")
	if CCTypeDC="" then
		CCTypeDC=0
	end if
	query="UPDATE CCTypes SET active="&CCTypeM&" WHERE CCcode='M';"
	set rstemp=conntemp.execute(query)
	query="UPDATE CCTypes SET active="&CCTypeV&" WHERE CCcode='V';"
	set rstemp=conntemp.execute(query)
	query="UPDATE CCTypes SET active="&CCTypeA&" WHERE CCcode='A';"
	set rstemp=conntemp.execute(query)
	query="UPDATE CCTypes SET active="&CCTypeD&" WHERE CCcode='D';"
	set rstemp=conntemp.execute(query)
	query="UPDATE CCTypes SET active="&CCTypeDC&" WHERE CCcode='DC';"
	set rstemp=conntemp.execute(query)
	call closedb()
end function

Function OffLineCustomEdit()
	call opendb()
	'request gateway variables
	idPayment=request.Form("id")
	terms=replace(request.Form("Cterms"),"'","''")
	terms=replace(terms,vbcrlf,"<br>")
	paymentDesc=replace(request.Form("CDesc"),"'","''")
	Creq=request.Form("Creq")
	if Creq="1" then
		Creq="-1"
	Else
		Creq="0"
	End if
	Cprompt=replace(request.Form("Cprompt"),"'","''")
	Cbtob=request.Form("Cbtob")
	if Cbtob = "" then
		Cbtob = "0"
	end if
	if Cbtob="1" then
		Cbtob="-1"
	end if
	if Cbtob="2" then
		Cbtob2 = request("Cbtob2")
		if Cbtob2 = "" then
			Cbtob = "0"
		end if
	end if

	query="UPDATE PayTypes SET paymentDesc='"&paymentDesc&"', terms='"& terms &"', Creq="& Creq &", Cbtob="& Cbtob &", Cprompt='"& Cprompt &"',priceToAdd="& priceToAdd &",percentageToAdd="& percentageToAdd &", paymentNickName='"&paymentNickName&"', pcPayTypes_processOrder=" & pcv_processOrder & ",pcPayTypes_setPayStatus=" & pcv_setPayStatus &"  WHERE idPayment="&idPayment
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	if err.number <> 0 then
		call closedb()
	  response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddPaymentOpt: "&Err.Description) 
	end If

	query = "DELETE FROM CustCategoryPayTypes WHERE idPayment = "&idPayment&";"
	set rstemp=Server.CreateObject("ADODB.Recordset")     
	set rstemp=conntemp.execute(query)
	
	if Cbtob="2" then
		CbtobArray=split(Cbtob2,",")
		for t=lbound(CbtobArray) to ubound(CbtobArray)
			query = "INSERT INTO CustCategoryPayTypes (idCustomerCategory,idPayment) VALUES ("&CbtobArray(t)&" ,"&idPayment&");"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
		next
	end if

	call closedb()
end function
%>                


<% 
Dim pcv_NoCPC, pcv_CPCString

if request("gwchoice")="6" then 
	M="0"
	V="0"
	A="0"
	D="0"
	DC="0"
	
	if request("mode")="Edit" then
		call opendb()
		query= "SELECT CCType,CCcode FROM CCTypes WHERE active=-1"
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if err.number <> 0 then
			strErrorDescription=err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error: "&strErrorDescription) 
		end If
		do until rs.eof
			CCType=rs("CCType")
			CCcode=rs("CCcode")
			select case CCcode
				case "V"
					V="-1"
				case "M"
					M="-1"
				case "A"
					A="-1"
				case "D"
					D="-1"
				case "DC"
					DC="-1"
			end select
			rs.moveNext
		loop 
		
		'Get Customer Categories
		query="SELECT idCustomerCategory, pcCC_Name FROM pcCustomerCategories Order by pcCC_Name;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		rs.Open query, conntemp
		if err.number <> 0 then
			strErrDescription = err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 345: "&strErrDescription) 
		end If

		pcv_NoCPC=1
		pcv_CPCString = ""
		pcv_IdPayment = request("id")

		do until rs.eof
			pcv_NoCPC=0
			pcv_IdCustomerCategory = rs("idCustomerCategory")
			pcv_CCName = rs("pcCC_Name")
			'check if this is selected for the CC already.
			query = "SELECT * FROM CustCategoryPayTypes WHERE idPayment="&pcv_IdPayment&" AND idCustomerCategory="&pcv_IdCustomerCategory&";"
			set rsCPCObj=Server.CreateObject("ADODB.Recordset")     
			set rsCPCObj=connTemp.execute(query)
			if rsCPCObj.eof then
				pcv_Checked=""
			else
				pcv_Checked=" checked"
			end if
			set rsCPCObj=nothing
			pcv_CPCString = pcv_CPCString&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
			rs.movenext
		loop
		set rs=nothing
		
		call closedb()
		%>
		<input type="hidden" name="mode" value="Edit">
        <% 
	end if
	%>
	<input type="hidden" name="addGw" value="6">
    <input type="hidden" name="id" value="<%=request("id")%>" />
    <tr> 
        <td height="21" colspan="2"><b>Modify offline credit card processing settings</b></td>
    </tr>
    <tr> 
        <td>
        	<table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr bgcolor="#FFFFFF"> 
                    <td colspan="2">Accepted Cards: 
                        <% if M="-1" then %> <input name="CCTypeM" type="checkbox" class="clearBorder" value="-1" checked> 
                        <% else %> <input name="CCTypeM" type="checkbox" class="clearBorder" value="-1"> 
                        <% end if %>
                        MasterCard 
                        <% if V="-1" then %> <input name="CCTypeV" type="checkbox" class="clearBorder" value="-1" checked> 
                        <% else %> <input name="CCTypeV" type="checkbox" class="clearBorder" value="-1"> 
                        <% end if %>
                        Visa 
                        <% if A="-1" then %> <input name="CCTypeA" type="checkbox" class="clearBorder" value="-1" checked> 
                        <% else %> <input name="CCTypeA" type="checkbox" class="clearBorder" value="-1"> 
                        <% end if %>
                        American Express 
                        <% if D="-1" then %> <input name="CCTypeD" type="checkbox" class="clearBorder" value="-1" checked> 
                        <% else %> <input name="CCTypeD" type="checkbox" class="clearBorder" value="-1"> 
                        <% end if %>
                        Discover 
                        <% if DC="-1" then %> <input name="CCTypeDC" type="checkbox" class="clearBorder" value="-1" checked> 
                        <% else %> <input name="CCTypeDC" type="checkbox" class="clearBorder" value="-1"> 
                        <% end if %>
                        Diners Club</td>
                </tr>
                <tr> 
                    <th colspan="2">You have the option to charge a processing fee for this payment option.</th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
                </tr>
                <tr> 
                    <th colspan="2">You can change the display name that is shown for this payment type. </th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td>&nbsp;</td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="0" <% if Cbtob="0" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to all customers</td>
                    </tr>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="1" <% if Cbtob="-1" then%>checked<% end if %>  /></td>
                      <td width="98%">Apply to Wholesale Customers only</td>
                    </tr>
                    <% if pcv_CPCString&""<>"" then %>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="2" <% if Cbtob="2" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to the following customer pricing categories</td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td><table width="95%" border="0" cellspacing="0" cellpadding="2">
                        <%=pcv_CPCString%>
                      </table></td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td colspan="2" height="10"></td>
                    </tr>
                  </table></td>
                </tr>
            </table>
        </td>
    </tr>
<% end if %>




<% if request("gwchoice")="7" then 
	if request("mode")="Edit" then
		call opendb()
		'Get Customer Categories
		query="SELECT idCustomerCategory, pcCC_Name FROM pcCustomerCategories Order by pcCC_Name;"
		set rs=Server.CreateObject("ADODB.Recordset")     
		rs.Open query, conntemp
		if err.number <> 0 then
			strErrDescription = err.description
			set rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in AddCCPaymentOpt 345: "&strErrDescription) 
		end If

		pcv_NoCPC=1
		pcv_CPCString = ""
		pcv_IdPayment = request("id")

		do until rs.eof
			pcv_NoCPC=0
			pcv_IdCustomerCategory = rs("idCustomerCategory")
			pcv_CCName = rs("pcCC_Name")
			'check if this is selected for the CC already.
			query = "SELECT * FROM CustCategoryPayTypes WHERE idPayment="&pcv_IdPayment&" AND idCustomerCategory="&pcv_IdCustomerCategory&";"
			set rsCPCObj=Server.CreateObject("ADODB.Recordset")     
			set rsCPCObj=connTemp.execute(query)
			if rsCPCObj.eof then
				pcv_Checked=""
			else
				pcv_Checked=" checked"
			end if
			set rsCPCObj=nothing
			pcv_CPCString = pcv_CPCString&"<tr><td width=""6%""><input type=""checkbox"" name=""Cbtob2"" value='"&pcv_IdCustomerCategory&"' class='clearBorder'"&pcv_Checked&"></td><td width=""94%"">Apply to "&pcv_CCName&"</td></tr>"
			rs.movenext
		loop
		set rs=nothing
		
		call closedb()
		%>
		<input type="hidden" name="mode" value="Edit">
        <% 
	end if
	%>
	<input type="hidden" name="addGw" value="7">
    <input type="hidden" name="id" value="<%=request("id")%>" />
    <tr> 
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <th colspan="2">Modify non real-time payment option</th>
    </tr>
    <tr> 
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr> 
        <td colspan="2">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr> 
                    <td width="20%" valign="top">Description:</td>
                    <td height="80%"><input type="text" name="CDesc" value="<%=paymentDesc%>"></td>
                </tr>
                <tr> 
                    <td width="20%" valign="top">Terms:<br />(Displayed to customers during checkout)</td>
                    <td width="80%"><textarea name="Cterms" cols="30" rows="5"><%if terms<>"" then%><%=replace(terms,"<br>",vbcrlf)%><%end if%></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td>&nbsp;</td>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="0" <% if Cbtob="0" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to all customers</td>
                    </tr>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="1" <% if Cbtob="-1" then%>checked<% end if %>  /></td>
                      <td width="98%">Apply to Wholesale Customers only</td>
                    </tr>
                    <% if pcv_CPCString&""<>"" then %>
                    <tr>
                      <td width="2%" align="right"><input type="radio" name="Cbtob" id="Cbtob" value="2" <% if Cbtob="2" then%>checked<% end if %> /></td>
                      <td width="98%">Apply to the following customer pricing categories</td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td><table width="95%" border="0" cellspacing="0" cellpadding="2">
                        <%=pcv_CPCString%>
                      </table></td>
                    </tr>
                    <% end if %>
                    <tr>
                      <td colspan="2" height="10"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr> 
                    <td colspan="2">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <% if pcv_CPCString&""<>"" then %>
                            <% end if %>
                            <tr> 
                                <td colspan="2" height="10"><b>Optional:</b><br />Checking the box below will prompt the customer to input more information, such as an account number or purchase order number, before completing their order.</td>
                            </tr>
                            <tr> 
                                <td width="5%"> <% if CReq="-1" then %> <input type="checkbox" class="clearBorder" name="Creq" value="1" checked> 
                                    <% else %> <input type="checkbox" class="clearBorder" name="Creq" value="1"> 
                                    <% end if %> </td>
                                <td width="95%">Require additional information for this payment option</td>
                            </tr>
                        </table></td>
                </tr>
                <tr>
                    <td width="20%" valign="top">Terms:<br />(Displayed to customers during checkout)</td>
                    <td width="80%"><input type="text" name="Cprompt" value="<%=Cprompt%>"></td>
                </tr>
                <tr> 
                    <th colspan="2">You have the option to charge a processing fee for this payment option.</th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td width="111"> <div align="right">Processing fee: </div></td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="price" <% if priceToAddType="price" then%>checked<%end if%>>
                        Flat fee&nbsp;&nbsp; &nbsp;<%=scCurSign%> <input name="priceToAdd" size="6" value="<%=money(priceToAdd)%>">
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td>&nbsp;</td>
                    <td> <input type="radio" class="clearBorder" name="priceToAddType" value="percentage" <% if priceToAddType="percentage" then%>checked<%end if%>>
                        Percentage of Order Total&nbsp;&nbsp; &nbsp; % 
                        <input name="percentageToAdd" size="6" value="<%=percentageToAdd%>">
                    </td>
                </tr>
                <tr> 
                    <th colspan="2">You can change the display name that is shown for this payment type. </th>
                </tr>
                <tr bgcolor="#FFFFFF"> 
                    <td><div align="right">Payment Name:&nbsp;</div></td>
                    <td><input name="paymentNickName" value="<%=paymentNickName%>" size="35" maxlength="255"></td>
                </tr>
            </table></td>
    </tr>
<% end if %>
