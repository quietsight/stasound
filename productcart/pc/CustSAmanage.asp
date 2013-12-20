<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="header.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp"-->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp"-->
<% 
'// Check if store is turned off and return message to customer
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

dim query, conntemp, rs
call openDb()
%> 

<div id="pcMain">		
	<table class="pcMainTable">   
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_CustSAmanage_1")%></h1>
			</td>
		</tr>
		<tr>
			<td class="pcSectionTitle">
			<p><a href="CustAddShip.asp"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_5")%></a></p>
			</td>
		</tr>
		<tr>
			<td>

				<%if request("msg")<>"" then%>
					<div class="pcSuccessMessage">
					<%if request("msg")="1" then
							response.write dictLanguage.Item(Session("language")&"_CustSAmanage_7")
						elseif request("msg")="2" then
							response.write dictLanguage.Item(Session("language")&"_CustSAmanage_8")
						elseif request("msg")="3" then
							response.write dictLanguage.Item(Session("language")&"_CustSAmanage_9")
						end if %>
					</div>
				<%end if%>
				
				<table class="pcShowContent">  
				<% 
				query="SELECT address, city, state, stateCode, shippingaddress, shippingcity, shippingState, shippingStateCode FROM customers WHERE (((idcustomer)="&session("idCustomer")&"));"

				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				pcDefaultAddress=rs("address")
				pcDefaultCity=rs("city")
				pcDefaultState=rs("state")
				pcDefaultStateCode=rs("stateCode")
				pcStrDefaultShipAddress=rs("shippingAddress")
				If len(pcStrDefaultShipAddress)<1 then
					pcStrDefaultShipAddress=pcDefaultAddress
					pcStrDefaultShipCity=pcDefaultCity
					pcStrDefaultShipState=pcDefaultState
					pcStrDefaultShipStateCode=pcDefaultStateCode
				Else
					pcStrDefaultShipCity=rs("shippingCity")
					pcStrDefaultShipState=rs("shippingState")
					pcStrDefaultShipStateCode=rs("shippingStateCode") 
				End if
				pcStrDefaultShipState=pcStrDefaultShipState & pcStrDefaultShipStateCode
				set rs=nothing
				%>
				<tr> 
          <td width="80%">
					<p><%=dictLanguage.Item(Session("language")&"_CustSAmanage_10")%></p>
          </td>
          <td width="20%" nowrap> 
          <p><a href="CustModShip.asp?reID=0"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_3")%></a></p>
          </td>
        </tr>
				<% 
				query="SELECT idRecipient, recipient_NickName, recipient_FullName, recipient_Address, recipient_City, recipient_State, recipient_StateCode FROM recipients WHERE (((idCustomer)="&session("idCustomer")&"));"
				set rs = Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				If rs.eof then
					intShipAddressExist=0
				end if
        
				do while not rs.eof
					intShipAddressExist=1
					IDre=rs("idRecipient")
					reNickName=trim(rs("recipient_NickName"))
					reFullName=trim(rs("recipient_FullName"))
					reShipAddr=ucase(rs("recipient_Address"))
					reShipCity=ucase(rs("recipient_City"))
					reShipState=ucase(rs("recipient_State") & rs("recipient_StateCode"))        	

					if len(reNickName)<1 then
						reNickName=dictLanguage.Item(Session("language")&"_CustSAmanage_12")
					end if %>
					<tr> 
						<td>
							<p><%=reNickName%></p>
						</td>
						<td nowrap> 
						 <p>
						 <a href="CustModShip.asp?reID=<%=IDre%>"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_3")%></a> | <a href="javascript:if (confirm('<%=dictLanguage.Item(Session("language")&"_CustSAmanage_11")%>')) location='CustDelShip.asp?reID=<%=IDre%>'"><%=dictLanguage.Item(Session("language")&"_CustSAmanage_4")%></a>
						 </p>
						</td>
					</tr>
					<% 					
					rs.movenext
				loop
				set rs = nothing
				call closeDb()
				%>	
      	</table>
	 </td>
	</tr>
	<tr>
		<td><hr></td>
	</tr>
	<tr>
		<td>
		<a href="CustPref.asp"><img src="<%=rslayout("back")%>"></a>
		</td>
	</tr>
</table>
</div>
<!--#include file="footer.asp"-->