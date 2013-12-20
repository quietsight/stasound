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
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<% 
on error resume next
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If 

dim query, conntemp, rstemp

call openDb()

query="select pcEv_IDEvent,pcEv_Name,pcEv_Type,pcEv_Date,pcEv_Hide,pcEv_Active from pcEvents where pcEv_IDCustomer=" & Session("idcustomer")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

%> 
<!--#include file="header.asp"-->
<div id="pcMain">
<table class="pcMainTable">
<tr>
	<td width="100%"><h1><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_1")%></h1></td>
</tr>
<% if rstemp.eof then%>
<tr>
	<td width="100%">
		<div class="pcErrorMessage">
			<%response.write dictLanguage.Item(Session("language")&"_ManageGRs_3")%>
		</div>
	</td>
</tr>
<%else%>
<tr>
	<td width="100%">
	<table class="pcMainTable">
        <tr>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_4")%></th>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_5")%></th>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_6")%></th>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_7")%></th>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_8")%></th>
            <th nowrap><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_9")%></th>
        </tr>
	<%do while not rstemp.eof
		gIDEvent=rstemp("pcEv_IDEvent")
		gType=rstemp("pcEv_Type")
		if gType<>"" then
		else
			gType="N/A"
		end if
		gName=rstemp("pcEv_Name")
		gDate=rstemp("pcEv_Date")
		if year(gDate)="1900" then
			gDate=""
		end if
		if gDate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gDate=(day(gDate)&"/"&month(gDate)&"/"&year(gDate))
			else
				gDate=(month(gDate)&"/"&day(gDate)&"/"&year(gDate))
			end if
        end if
        gHide=rstemp("pcEv_Hide")
	        if gHide<>"" then
	        else
		        gHide="0"
	        end if
	        gActive=rstemp("pcEv_Active")
	        if gActive<>"" then
	        else
		        gActive="0"
	        end if
	        query="select sum(pcEP_Qty) as gQty,sum(pcEP_HQty) as gHQty from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_GC=0 group by pcEP_IDEvent"
	        set rs1=connTemp.execute(query)
	        if err.number<>0 then
				call LogErrorToDatabase()
				set rs1=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
	        if not rs1.eof then
		        gQty=rs1("gQty")
		        gHQty=rs1("gHQty")
	        else
		        gQty="0"
		        gHQty="0"
	        end if
	        set rs1=nothing
	        if gQty<>"" then
	        else
		        gQty="0"
	        end if
			        
	        if gHQty<>"" then
	        else
		        gHQty="0"
	        end if%>
	        <tr>
				<td nowrap><strong><%=gName%></strong></td>
				<td nowrap><%=gType%></td>
				<td nowrap><%=gDate%></td>
				<td nowrap>
					<%if gHide="1" then%>
						<font color="#FF0000"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_7b")%></font>
					<%else%>
						<%response.write dictLanguage.Item(Session("language")&"_ManageGRs_7a")%>
					<%end if%>
				</td>
				<td nowrap>
					<%if gActive="1" then%>
						<%response.write dictLanguage.Item(Session("language")&"_ManageGRs_8a")%>
					<%else%>
						<font color="#FF0000"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_8b")%></font>
					<%end if%>
				</td>
				<td nowrap align="center"><%=gQty%> (<%=clng(gQty)-clng(gHQty)%>)</td>
	        </tr>
			<tr>
				<td colspan="6" align="center"><a href="JavaScript:;" onClick="document.getElementById('addProductsInfo<%=gIDEvent%>').style.display='';"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_14")%></a> - <a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_12")%></a> - <a href="ggg_NotifyGR.asp?IDEvent=<%=gIDEvent%>"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_11")%></a> - <a href="ggg_EditGR.asp?IDEvent=<%=gIDEvent%>"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_13")%></a>
                
                <div class="pcInfoMessage" id="addProductsInfo<%=gIDEvent%>" style="display: none;"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_20")%> <a href="default.asp"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_21")%></a> | <a href="JavaScript:;" onClick="document.getElementById('addProductsInfo<%=gIDEvent%>').style.display='none';"><%response.write dictLanguage.Item(Session("language")&"_ManageGRs_22")%></a></div>
                
                </td>
			</tr>
			<tr>
				<td colspan="6"><hr></td>
			</tr>
		<%rstemp.movenext
	loop
	set rstemp=nothing%>
	</table>
<% end if 'not rstemp.eof
call closeDb() %>
	</td>
</tr>
<tr> 
	<td>
	<a href="ggg_instGR.asp"><img src="<%=rslayout("CreRegistry")%>" border=0></a><br><br><a href="javascript:history.go(-1)"><img src="<%=rslayout("back")%>" border=0></a></td>
</tr>
</table>
</div>
<!--#include file="footer.asp"-->