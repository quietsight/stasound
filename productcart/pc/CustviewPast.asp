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

<%
on error resume next

'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
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

query="SELECT idOrder,orderstatus,orderDate,total,ord_OrderName FROM orders WHERE idCustomer=" &Session("idcustomer") &" AND OrderStatus>1 ORDER BY idOrder DESC"
set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
	set rstemp=nothing
	call closeDb()
 	response.redirect "techErr.asp?error="& Server.Urlencode("Error in CustviewPast: "&err.description) 
end If

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
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
					response.write(session("pcStrCustName") & " - " & dictLanguage.Item(Session("language")&"_CustviewPast_4"))
					else
					response.write(dictLanguage.Item(Session("language")&"_CustviewPast_4"))
				end if
				%>
				</h1>
			</td>
		</tr>
		<tr>
			<td>     
			<table class="pcShowContent">
				<tr>
			    	<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_5")%></th>
                	<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewOrd_11")%></th>
					<%if scOrderName="1" then 'Show order name %>
						<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_9")%></th>
					<% end if %>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_7")%></th>
					<th nowrap><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_6")%></th>
					<th>&nbsp;</th>
				</tr>
				<tr class="pcSpacer">
					<td colspan="6"></td>
				</tr>
				<%do while not rstemp.eof and pCnt<iPageSize
						pCnt=pCnt+1
						pIdOrder = rstemp("idOrder")
						porderstatus = rstemp("orderstatus")
						pOrderName = rstemp("ord_OrderName")
						pOrderTotal = rstemp("total")
						pOrderDate = rstemp("orderDate")
				
				%>
				<tr>
					<td>
						<a href="CustviewPastD.asp?idOrder=<%response.write (scpre+int(pIdOrder))%>"><%response.write (scpre+int(pIdOrder))%></a>
					</td>
                    <td>
                    	<!--#include file="inc_orderStatus.asp"-->
                    </td>
					<%if scOrderName="1" then 'Show order name %>
					<td>
						<%=pOrderName%>
					</td>
					<% end if %>
					<td>
						<%=scCurSign&money(pOrderTotal)%>
					</td>
					<td>
						<%=showdateFrmt(pOrderDate)%>
					</td>
					<td nowrap>
						<div align="right" class="pcSmallText">
							<a href="CustviewPastD.asp?idOrder=<%response.write (scpre+int(pIdOrder))%>"><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_3")%></a>
							&nbsp;
							<a href="RepeatOrder.asp?idOrder=<%=pIdOrder%>"><%response.write dictLanguage.Item(Session("language")&"_CustviewPast_8")%></a>
						<% 'Hide/show link to Help Desk
						If scShowHD <> 0 then %>
							&nbsp;
							<a href="userviewallposts.asp?idOrder=<%=clng(scpre)+clng(pIdOrder)%>"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_3")%></a>
						</div>
						<% end if %>
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
									<a href="CustviewPast.asp?iPageCurrent=1"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>&nbsp;
					        	<% end if %>
								<% if cint(iPageCurrent)>1 then
	            					if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
	                					iPagePrev=cint(iPageCurrent)-1
	            					else
	                					iPagePrev=iRecSize
	            					end if %>
	            					<a href="CustviewPast.asp?iPageCurrent=<%=cint(iPageCurrent)-iPagePrev%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2")%>&nbsp;<%=iPagePrev%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
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
		      						<a href="CustviewPast.asp?iPageCurrent=<%=pageNumber%>"><%=pageNumber%></a>
								<% End If 
							Next
	
							if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
							else
								if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
									<a href="CustviewPast.asp?iPageCurrent=<%=cint(intPageNumber)+iPageNext%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4")%>&nbsp;<%=iPageNext%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
								<% end if
    
								if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
						    		&nbsp;<a href="CustviewPast.asp?iPageCurrent=<%=cint(iPageCount)%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
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