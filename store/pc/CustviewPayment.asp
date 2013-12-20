<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/rc4.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

msg=getUserInput(request.querystring("message"),0)

'// DELETE VAULT
Dim pcv_intVaultID
pcv_intVaultID=getUserInput(request.querystring("VaultID"),8)
If len(pcv_intVaultID)>0 Then

	call openDb()
	
	'// Load Settings
	query="SELECT pcPay_EIG_Type, pcPay_EIG_Username, pcPay_EIG_Password, pcPay_EIG_Key, pcPay_EIG_Curcode, pcPay_EIG_CVV, pcPay_EIG_TestMode, pcPay_EIG_SaveCards, pcPay_EIG_UseVault FROM pcPay_EIG WHERE pcPay_EIG_ID=1"
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)		
	if NOT rs.eof then
		x_Username=rs("pcPay_EIG_Username")
		x_Username=enDeCrypt(x_Username, pcs_GetSecureKey)
		x_Password=rs("pcPay_EIG_Password")
		x_Password=enDeCrypt(x_Password, pcs_GetSecureKey)
		x_Key=rs("pcPay_EIG_Key")
		x_Key=enDeCrypt(x_Key, pcs_GetSecureKey)
		x_CVV=rs("pcPay_EIG_CVV")
		x_Type=rs("pcPay_EIG_Type")
		x_TypeArray=Split(x_Type,"||")
		x_TransType=x_TypeArray(0)
		x_Curcode=rs("pcPay_EIG_Curcode")
		x_TestMode=rs("pcPay_EIG_TestMode")
		x_SaveCards=rs("pcPay_EIG_SaveCards")
		x_UseVault=rs("pcPay_EIG_UseVault")
	end if
	set rs=nothing
	
	'// Get the Vault ID
	query="SELECT pcPay_EIG_Vault_Token FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_ID="& pcv_intVaultID &""
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)		
	if NOT rs.eof then
		pcv_strCustomerVaultID = rs("pcPay_EIG_Vault_Token")
		pcv_strCustomerVaultID=enDeCrypt(pcv_strCustomerVaultID, pcs_GetSecureKey)
	end if
	set rs=nothing
	
	'// Contact Vault
	strTest = ""
	strTest = strTest & "username=" & x_Username
	strTest = strTest & "&password=" & x_Password	
	strTest = strTest & "&customer_vault=delete_customer"
	strTest = strTest & "&customer_vault_id=" & pcv_strCustomerVaultID

	set xml = Server.CreateObject("Msxml2.serverXmlHttp"&scXML)
	xml.open "POST", "https://secure.nmi.com/api/transact.php", false
	xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xml.send strTest
	strStatus = xml.Status
	strRetVal = xml.responseText
	Set xml = Nothing

	query="DELETE FROM pcPay_EIG_Vault WHERE pcPay_EIG_Vault_ID="& pcv_intVaultID &""
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)	
	set rs=nothing
	
	call closeDb()

	response.redirect "CustviewPayment.asp?message="& Server.Urlencode("Credit Card Successfully Removed") 
	
End If

iPageSize=25
iPageCurrent=getUserInput(request("iPageCurrent"),0)
if iPageCurrent="" then
	iPageCurrent=1
end if
if not IsNumeric(iPageCurrent) then
	response.redirect "CustPref.asp"
end if

dim query, conntemp, rstemp

call openDb()

query="SELECT pcPay_EIG_Vault_ID, pcPay_EIG_Vault_CardNum, pcPay_EIG_Vault_CardExp FROM pcPay_EIG_Vault WHERE idCustomer="& Session("idCustomer") & "  AND IsSaved=1"
set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in CustviewPayment.asp: "&err.description) 
end If

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=307"     
end if

iPageCount=rstemp.PageCount

	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	rstemp.AbsolutePage=iPageCurrent
	pCnt=0         

%> 

<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">   
		<tr>
			<td>
				<h1>
				<%
				if session("pcStrCustName") <> "" then
					response.write(session("pcStrCustName") & " - " & dictLanguage.Item(Session("language")&"_EIG_11"))
					else
					response.write(dictLanguage.Item(Session("language")&"_EIG_11"))
				end if
				%>
				</h1>
			</td>
		</tr>
		<tr>
			<td>
            
			<% if msg<>"" then %>
                <div class="pcInfoMessage"><%=Msg%></div>
            <% end if %>
                 
			<table class="pcShowContent">
				<tr>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_EIG_12")%></th>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_EIG_13")%></th>
					<th>&nbsp;</th>
				</tr>
				<tr class="pcSpacer">
					<td colspan="6"></td>
				</tr>
				<%
				do while not rstemp.eof and pCnt<iPageSize
						pCnt=pCnt+1
						pcv_strCardNum = rstemp("pcPay_EIG_Vault_CardNum")
						pcv_strCardExp = rstemp("pcPay_EIG_Vault_CardExp")
						pcv_intVaultID = rstemp("pcPay_EIG_Vault_ID")
				%>
				<tr>
					<td>
						<%=pcv_strCardNum%>
					</td>
					<td>
						<%=pcv_strCardExp%>
					</td>
					<td nowrap>
						<div align="right" class="pcSmallText">
							<!--<a href="CustmodPayment.asp?VaultID=<%=pcv_intVaultID%>"><%response.write dictLanguage.Item(Session("language")&"_EIG_14")%></a>
							&nbsp; -->
							<a href="CustviewPayment.asp?VaultID=<%=pcv_intVaultID%>" class="delete" title="<%=pcv_strCardNum%>"><%response.write dictLanguage.Item(Session("language")&"_EIG_15")%></a>

                    	</div>
					</td>
				</tr>
				<%
				rstemp.movenext
			  loop
				%>
			</table>
  			<% 
			set rstemp = nothing
			call closeDb()
			%>
				<script type="text/javascript">
                    $(document).ready(function(){
                    	$('.delete').click(function(){
							var answer = confirm('<%response.write dictLanguage.Item(Session("language")&"_EIG_16")%> '+jQuery(this).attr('title')+'?');
                            return answer;
                		}); 
                    });  
                </script>
    
			</td>
		</tr>
		<tr>
			<td>
						<%
						iRecSize=10

						'*******************************
						' START Page Navigation
						'*******************************

						If iPageCount>1 then %>
							<div class="pcPageNav">
							<%response.write(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
							&nbsp;-&nbsp;
						    <% if iPageCount>iRecSize then %>
								<% if cint(iPageCurrent)>iRecSize then %>
									<a href="CustviewPayment.asp?iPageCurrent=1"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>&nbsp;
					        	<% end if %>
								<% if cint(iPageCurrent)>1 then
	            					if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
	                					iPagePrev=cint(iPageCurrent)-1
	            					else
	                					iPagePrev=iRecSize
	            					end if %>
	            					<a href="CustviewPayment.asp?iPageCurrent=<%=cint(iPageCurrent)-iPagePrev%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2")%>&nbsp;<%=iPagePrev%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
								<% end if
								if cint(iPageCurrent)+1>1 then
									intPageNumber=cint(iPageCurrent)
								else
									intPageNumber=1
								end if
							else
								intPageNumber=1
							end if

							if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
								iPageNext=cint(iPageCount)-cint(iPageCurrent)
							else
								iPageNext=iRecSize
							end if

							For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
								If Cint(pageNumber)=Cint(iPageCurrent) Then %>
									<strong><%=pageNumber%></strong> 
								<% Else %>
		      						<a href="CustviewPayment.asp?iPageCurrent=<%=pageNumber%>"><%=pageNumber%></a>
								<% End If 
							Next
	
							if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
							else
								if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
									<a href="CustviewPayment.asp?iPageCurrent=<%=cint(intPageNumber)+iPageNext%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4")%>&nbsp;<%=iPageNext%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
								<% end if
    
								if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
						    		&nbsp;<a href="CustviewPayment.asp?iPageCurrent=<%=cint(iPageCount)%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
						    	<% end if 
							end if 

						end if

						'*******************************
						' END Page Navigation
						'*******************************
						%>
			</td>
		</tr>

		<tr>
			<td><hr></td>
		</tr>
		<tr> 
			<td><a href="custPref.asp"><img src="<%=rslayout("back")%>"></a></td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->