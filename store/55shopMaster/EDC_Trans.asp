<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Endicia Postage Label Services - View Transactions" %>
<% response.Buffer=true %>
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
<%Dim connTemp,rs,query

private const MaxRecords=25 'Max Records to display per page

if request("iPageCurrent")="" then
    iPageCurrent=1 
else
    iPageCurrent=Request("iPageCurrent")
end If

pcPageName="EDC_Trans.asp"

call opendb()

IF request("action")="removeall" THEN
	query="DELETE FROM pcEDCTrans;"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	
	query="DELETE FROM pcEDCLogs;"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	
	call closedb()
	
	response.redirect "EDC_manage.asp?msg=" & Server.URLEncode("All transaction log(s) were removed successfully!") & "&s=1"
END IF

HaveRecords=0

query="SELECT pcET_ID,IDOrder,pcPackageInfo_ID,pcET_TransDate,pcET_Method,pcET_Success,pcET_LabelFile,pcET_PicNum,pcET_CustomsNum,pcET_TransID,pcET_Postage,pcET_subPostage,pcET_Fees,pcET_RefundID,pcET_ErrMsg,pcET_FeesDetails FROM pcEDCTrans ORDER BY pcET_ID DESC;"
Set rsInv=Server.CreateObject("ADODB.Recordset")
rsInv.CacheSize=MaxRecords
rsInv.PageSize=MaxRecords
rsInv.Open query, connTemp, adOpenStatic, adLockReadOnly
If rsInv.eof Then
	msg="Transactions not found!"
	msgType=0
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>
<%Else
	HaveRecords=1%>
	<table class="pcCPcontent">
	<tr>
	<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>ID#</th>
		<th>Method</th>
		<th>Date</th>
		<th>Status</th>
		<th>Error Msg.</th>
		<th>Information</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
	<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<%
	rsInv.MoveFirst
	' get the max number of pages
	Dim iPageCount
	iPageCount=rsInv.PageCount
	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
	If iPageCurrent < 1 Then iPageCurrent=1
				
	' set the absolute page
	rsInv.AbsolutePage=iPageCurrent  
			
	Count=0
	Do While NOT rsInv.EOF And Count < rsInv.PageSize
	count=count + 1
	tmpETID=rsInv("pcET_ID")
	tmpIDOrder=rsInv("IDOrder")
	tmpPackID=rsInv("pcPackageInfo_ID")
	tmpTransDate=rsInv("pcET_TransDate")
	tmpMethod=rsInv("pcET_Method")
	tmpSuccess=rsInv("pcET_Success")
	tmpPIC=rsInv("pcET_PicNum")
	tmpCustomsNum=rsInv("pcET_CustomsNum")
	tmpPostagePrice=rsInv("pcET_Postage")
	tmpsubPostage=rsInv("pcET_subPostage")
	tmpFees=rsInv("pcET_Fees")
	tmpRefundID=rsInv("pcET_RefundID")
	tmpErrMsg=rsInv("pcET_ErrMsg")
	tmpFeesDetails=rsInv("pcET_FeesDetails")%>
	<tr valign="top" onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist">
		<td><%=tmpETID%>
		<td nowrap><b>
			<%Select Case tmpMethod
			Case "1": response.write "Get Postage Label Request"
			Case "2": response.write "Refill account"
			Case "3": response.write "Change Pass Phrase"
			Case "5": response.write "Get Account Status"
			Case "7": response.write "Refund Request"
			End Select%>
			</b>
		</td>
		<td><%=tmpTransDate%></td>
		<td>
			<%if tmpSuccess="1" then
				response.write "Success"
			else
				response.write "Failure"
			end if%>
		</td>
		<td>
			<%=tmpErrMsg%>
		</td>
		<td>
			<%Select Case tmpMethod
			Case "1":if tmpSuccess="1" then%>
				Order # <%=scpre+int(tmpIDOrder)%><br>
				Package # <%=tmpPackID%><br>
				Tracking Number: <%=tmpPIC & tmpCustomsNum%><br>
				Postage Price: <%=scCurSign & money(tmpPostagePrice)%><br>
				<%if tmpFees>"0" then%>
				-------------------<br>
				Sub Postage: <%=scCurSign & money(tmpsubPostage)%><br>
				<%if tmpFeesDetails<>"" then
					tmpFeesDetails=replace(tmpFeesDetails,"|~|","<br>")
					tmpFeesDetails=replace(tmpFeesDetails,"|!|",": ")%>
					Additional Fees:<br>
					<%=tmpFeesDetails%>
				<%end if
				end if%>
				<%if tmpRefundID>"0" then%>
				Related Refund # <%=tmpRefundID%>
				<%end if
				end if
			Case "2": if tmpSuccess="1" then%>Refilled Amount: <%=scCurSign & money(tmpPostagePrice)%><%end if
			Case "5": if tmpSuccess="1" then%>Account Balance: <%=scCurSign & money(tmpPostagePrice)%><%end if
			Case "7": if tmpSuccess="1" then%>
				Order # <%=scpre+int(tmpIDOrder)%><br>
				Package # <%=tmpPackID%><br>
				Tracking Number: <%=tmpPIC & tmpCustomsNum%>
				<%end if
			End Select%>
		</td>
		<td>
			<%query="SELECT pcET_ID FROM pcEDCLogs WHERE pcET_ID=" & tmpETID & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then%>
				<a href="EDC_Log.asp?id=<%=tmpETID%>&iPageCurrent=<%=iPageCurrent%>">More Details</a>
			<%end if
			set rsQ=nothing%>
		</td>
	</td>
	</tr>
	<% 
	rsInv.MoveNext
	Loop
	%>
	<tr>
	<td colspan="7" class="pcCPspacer"></td>
	</tr>
	<%If iPageCount>1 Then%>
  	<tr>
		<td colspan="7" class="pcCPspacer"></td>
	</tr>                            
	<tr> 
		<td colspan="7"><%response.Write("Currently viewing page "& iPageCurrent & " of "& iPageCount)%></td>
	</tr>
	<tr>                   
	<td colspan="7"> 
		<%' display Next / Prev buttons
		if iPageCurrent > 1 then %>
		<a href="EDC_Trans.asp?iPageCurrent=<%=iPageCurrent-1%>"><img src="../pc/images/prev.gif" width="10" height="10" border="0"></a> 
		<%
		end If
		For I=1 To iPageCount
		If Cint(I)=Cint(iPageCurrent) Then %>
			<b><%=I%></b> 
		<%
		Else
		%>
			<a href="EDC_Trans.asp?iPageCurrent=<%=I%>"><%=I%></a> 
		<%
		End If
		Next
		if CInt(iPageCurrent) < CInt(iPageCount) then %>
				<a href="EDC_Trans.asp?iPageCurrent=<%=iPageCurrent+1%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
		<%
		end If
		%>
	</td>
	</tr>
	<tr>
	<td colspan="7" class="pcCPspacer"></td>
	</tr>
<% End If %>
	</table>
<%End If%>
<table class="pcCPcontent">
<tr>
	<td>&nbsp;</td>
	<td><%if HaveRecords=1 then%><input type="button" name="del" value=" Remove All Transaction Log(s) " onclick="javascript:location='EDC_Trans.asp?action=removeall';" class="submit2">&nbsp;&nbsp;<%end if%><input type="button" name="back" value=" Back to Manage Endicia Settings " onclick="javascript:location='EDC_manage.asp';" class="submit2"></td>
</tr>
</table>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->