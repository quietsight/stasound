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
<!--#include FILE="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartinc.asp"-->
<% 
dim conntemp, query, rs
call opendb()
IF request("action")="del" THEN
	tmpID=getUserInput(request("id"),0)
	if tmpID="" or IsNull(tmpID) then
		tmpID=0
	end if
	if not IsNumeric(tmpID) then
		tmpID=0
	end if
	IF tmpID>0 THEN
		query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartID=" & tmpID & " AND IDCustomer=" & session("IDCustomer") & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			query="DELETE FROM pcSavedCartArray WHERE SavedCartID=" & tmpID & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & tmpID & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
		end if
		set rsQ=nothing
	END IF
ELSE
	IF request("action")="res" THEN
		tmpID=getUserInput(request("id"),0)
		if tmpID="" or IsNull(tmpID) then
			tmpID=0
		end if
		if not IsNumeric(tmpID) then
			tmpID=0
		end if
		IF tmpID>0 THEN
			query="SELECT SavedCartGUID FROM pcSavedCarts WHERE SavedCartID=" & tmpID & " AND IDCustomer=" & session("IDCustomer") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				Response.Cookies("SavedCartGUID")=rsQ("SavedCartGUID")
				set rsQ=nothing
				call closedb()
				Response.Cookies("SavedCartGUID").Expires=Date()+365
				dim pcCartArray(100,45)
				session("pcCartSession")=pcCartArray
				Session("pcCartIndex")=0
				HaveToRestore="yes"
				%>
				<!--#include file="inc_RestoreShoppingCart.asp"-->
				<%
				response.redirect "viewcart.asp"
			end if
			set rsQ=nothing
		END IF
	END IF
END IF
call closedb()
%>
<!--#include file="header.asp"-->
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1>
					<%response.write dictLanguage.Item(Session("language")&"_CustPref_50")%>
				</h1>
			</td>
		</tr>
		<tr> 
			<td>
				<%
				call opendb()
				query="SELECT SavedCartID,SavedCartDate,SavedCartName FROM pcSavedCarts WHERE IDCustomer=" & session("IDCustomer") & " ORDER BY SavedCartID DESC;"
				set rsQ=Server.CreateObject("ADODB.Recordset")
				set rsQ=connTemp.execute(query)
				If rsQ.eof then 
					%>
					<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_1")%></div>
				<% else %>
					<table class="pcShowContent">
						<tr>
							<th><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_2")%></th>
                            <th><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_8")%></th>
							<th></th>
						</tr>
                    	<tr>
                        	<td class="pcSpacer"></td>
                        </tr>
						<%
                        pcArr=rsQ.getRows()
                        intCount=ubound(pcArr,2)
                        For i=0 to intCount
							Rev_Date = pcArr(1,i)
							If scDateFrmt="DD/MM/YY" then 
                            	Rev_Date = day(Rev_Date) & "/" & month(Rev_Date) & "/" & year(Rev_Date)
                        	Else
                            	Rev_Date = month(Rev_Date) & "/" & day(Rev_Date) & "/" & year(Rev_Date)
                        	End If
							%>
                            <tr>
                                <td><%=Rev_Date%></td>
                                <td><%=pcArr(2,i)%></td>
                                <td align="right">
                                    <a href="CustSavedCarts.asp?action=res&id=<%=pcArr(0,i)%>"><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_3")%></a>&nbsp;|&nbsp;<a href="CustSavedCartsRename.asp?id=<%=pcArr(0,i)%>"><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_9")%></a>&nbsp;|&nbsp;<a href="javascript:if (confirm('<%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_10")%>')) location='CustSavedCarts.asp?action=del&id=<%=pcArr(0,i)%>';"><%response.write dictLanguage.Item(Session("language")&"_CustSavedCarts_4")%></a>
                                </td>
                            </tr>
                        <%
                        Next
						%>
					</table>
                <% 
				end if
				set rsQ=nothing
				call closedb()
				%>
			</td>
		</tr>
       	<tr>
        	<td><hr></td>
        </tr>
        <tr> 
            <td>
                <a href="CustPref.asp"><img src="<%=rslayout("back")%>"></a>
                <a href="viewCart.asp"><img src="<%=rslayout("viewcartbtn")%>"></a>
            </td>
        </tr>
	</table>
</div>
<!--#include file="footer.asp"-->