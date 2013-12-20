<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<% pageTitle="View Saved Shopping Carts Statistics" %>
<!--#include file="AdminHeader.asp"-->
<%
 
dim query, conntemp, rstemp
call openDb()

PreviousMonthID=0
query="SELECT pcSCStatID FROM pcSavedCartStatistics ORDER BY pcSCStatID DESC;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	PreviousMonthID=rsQ("pcSCStatID")
end if
set rsQ=nothing
if PreviousMonthID>"0" then
	query="SELECT TOP 10 idProduct,pcSPS_SavedTimes FROM pcSavedPrdStats WHERE pcSPS_SavedTimes>0 ORDER BY pcSPS_SavedTimes DESC;"
	set rsQ=connTemp.execute(query)
	SCPrds=""
	if not rsQ.eof then
		iCount=0
		Do while (not rsQ.eof) AND (iCount<10)
			iCount=iCount+1
			SCPrds=SCPrds & rsQ("idproduct") & "|*|" & rsQ("pcSPS_SavedTimes") & "|$|"
			rsQ.MoveNext
		Loop
	end if
	set rsQ=nothing
				
	query="UPDATE pcSavedCartStatistics SET pcSCTopPrds='" & SCPrds & "' WHERE pcSCStatID=" & PreviousMonthID & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
end if
%>
		<table class="pcCPcontent">
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
			<%query="SELECT pcSCMonth,pcSCYear,pcSCTotals,pcSCTopPrds,pcSCAnonymous FROM pcSavedCartStatistics ORDER BY pcSCYear DESC,pcSCMonth DESC;"
			set rsQ=connTemp.execute(query)
			IF rsQ.eof THEN
				set rsQ=nothing%>
			<tr> 
              	<td colspan="4">
					<div class="pcCPmessage">The store does not contain enough data to build this report.</div>
				</td>
	          	</tr>
			<%ELSE
				tmpArr=rsQ.getRows()
				set rsQ=nothing
				intCount=ubound(tmpArr,2)
				%>
			<tr> 
              	<td colspan="4">
                This page provides aggregate, monthly statistics on shopping carts saved in your storefront. <a href="http://wiki.earlyimpact.com/productcart/reports-saved-carts" target="_blank">Learn more about this report</a> <a href="http://wiki.earlyimpact.com/productcart/reports-saved-carts" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Learn more" border="0"></a>. </td>
	        </tr>
            <tr>
            	<td colspan="4" class="pcCPspacer"></td>
            </tr>
          	<tr> 
              	<th nowrap>Month/Year</th>
				<th nowrap>Saved Cart Total</th>
				<th>Anonymous</th>
				<th nowrap>Top 10 Most Frequently Saved Products</th>
          	</tr>
            <tr>
            	<td colspan="4" class="pcCPspacer"></td>
            </tr>
			<%For i=0 to intCount%>
            <tr valign="top"> 
				<td><h3 style="margin: 0; padding: 0;"><%=tmpArr(0,i)%>/<%=tmpArr(1,i)%></h3></td>
				<td><%=tmpArr(2,i)%></td>
				<td><%=tmpArr(4,i)%></td>
				<%tmpList=tmpArr(3,i)
				if tmpList<>"" then
					tmpList1=split(tmpList,"|$|")
					tmpList=""
					for j=0 to ubound(tmpList1)
						if tmpList1(j)<>"" then
						tmpList2=split(tmpList1(j),"|*|")
						query="SELECT Description FROM Products WHERE idproduct=" & tmpList2(0) & ";"
						set rsQ=connTemp.execute(query)
						if not rsQ.eof then
							if tmpList<>"" then
								tmpList=tmpList & "<br>"
							end if
							tmpList=tmpList & rsQ("Description") & ": " & tmpList2(1) & " times"
						end if
						set rsQ=nothing
						end if
					next
				end if%>
				<td><%=tmpList%></td>
		  </tr>
            <tr>
            	<td colspan="4"><hr /></td>
            </tr>
			<%Next
			END IF%>
          	<tr> 
            	<td colspan="4" class="pcCPspacer"></td>
			</tr>
			<tr> 
            	<td colspan="4"><input type="button" name="Back" value=" Back to Main menu " onclick="location='menu.asp';" class="ibtnGrey"></td>
			</tr>
		</table>
<%
call closeDb()
%>
<!--#include file="AdminFooter.asp"-->