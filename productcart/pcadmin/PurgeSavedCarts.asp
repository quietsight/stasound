<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<% pageTitle="Remove Saved Shopping Carts" %>
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next

dim query, conntemp, rstemp
call openDb()

If request("submit1")<>"" OR request("submit2")<>"" OR request("submit3")<>"" then	

	If request("submit1")<>"" then
		RmvCarts=0
		'// 1.) Optimize Performance/ Purge Customer Sessions
		pcCustSession_Date=Date()
		if SQL_Format="1" then
			pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
		else
			pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
		end if
		if scDB="SQL" then
			strDtDelim="'"
		else
			strDtDelim="#"
		end if
		if request("submit1")<>"" then
			ndays=request("ndays")
			if ndays="" then
				ndays=7
			end if
			if not IsNumeric(ndays) then
				ndays=7
			end if
		end if
		ndays=-1*Clng(ndays)
		query="SELECT Count(*) As TotalCarts FROM pcSavedCarts WHERE SavedCartDate<" &strDtDelim&dateadd("d",ndays,pcCustSession_Date)&strDtDelim& ";"	
		set rstemp=server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)

		if not rstemp.eof then
			RmvCarts=rstemp("TotalCarts")
		end if
	
		set rstemp=nothing
	
		query="DELETE FROM pcSavedCartArray WHERE pcSavedCartArray.SavedCartID IN (SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartDate<" &strDtDelim&dateadd("d",ndays,pcCustSession_Date)&strDtDelim& ");"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="DELETE FROM pcSavedCarts WHERE SavedCartDate<" &strDtDelim&dateadd("d",ndays,pcCustSession_Date)&strDtDelim& ";"	
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
	end if
	
	If request("submit2")<>"" then
		R1=request("r1")
		if R1="" then
			R1=0
		end if
		if Not IsNumeric(R1) then
			R1=0
		end if
		if R1=0 then
			query1=" WHERE IDCustomer=0 "
		else
			if R1="1" then
				query1=" WHERE IDCustomer>0 "
			else
				query1=""
			end if
		end if
		query="SELECT Count(*) As TotalCarts FROM pcSavedCarts" & query1 & ";"	
		set rstemp=server.CreateObject("ADODB.Recordset")
		set rstemp=conntemp.execute(query)

		if not rstemp.eof then
			RmvCarts=rstemp("TotalCarts")
		end if
	
		set rstemp=nothing
	
		query="DELETE FROM pcSavedCartArray WHERE pcSavedCartArray.SavedCartID IN (SELECT SavedCartID FROM pcSavedCarts" & query1 & ");"
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		query="DELETE FROM pcSavedCarts" & query1 & ";"
		set rstemp=conntemp.execute(query)
		set rstemp=nothing
	end if
	If request("submit3")="" then%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<p>
                  <div class="pcCPmessageSuccess">
                        <b><%=RmvCarts%></b> saved shopping cart(s) were removed successfully.
                        <br />
                        <a href="PurgeSavedCarts.asp">Remove other data</a> or return to the <a href="menu.asp">Main Menu</a>.
                  </div>
            	</p>
			</td>
		</tr>
	</table>
	<%else
		query="DELETE FROM pcSavedPrdStats;"
		set rsQ=connTemp.execute(query)
		set rsQ=nothing%>
		<table class="pcCPcontent">
		<tr> 
			<td>
				<p>
                  <div class="pcCPmessageSuccess">
                        Product-level statistics were removed successfully.
                        <br />
                        <a href="PurgeSavedCarts.asp">Remove other data</a> or return to the <a href="menu.asp">Main Menu</a>.
                  </div>
            	</p>
			</td>
		</tr>
	</table>
	<%end if%>
<% else %>
	<form action="PurgeSavedCarts.asp" method="post" name="form" id="form" class="pcForms">
		<table class="pcCPcontent">
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
          	<tr> 
              	<th colspan="2">Based on the number of days</th>
          	</tr>
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
            <tr> 
				<td colspan="2">
                	When you remove this information from the database, registered customers will no longer view and be able to restore saved shopping carts. In addition, anonymous and registered customers will not be prompted to restore their shopping carts when they revisit the store after a previous visit in which they &quot;abandoned&quot; the cart.<br />
                	<br />
                    Remove all saved shopping carts older than <input type="text" name="ndays" value="7" size="10"> days
        		</td>
			</tr>
			<tr> 
            	<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="2">
					<input name="submit1" type="submit" id="submit1" value=" Clean Up Database " class="submit2">
        		</td>
       		</tr>				
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
			<tr> 
              	<th colspan="2">Based on who saved the cart</th>
          	</tr>
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
          	<tr> 
				<td colspan="2">
                	Same as above, but here you can discriminate between registered and anonymous &quot;saved carts&quot;, and there is no date filter.<br /><br />
					<input type="radio" name="r1" value="0" class="clearBorder" checked> 
					Delete anonymously saved carts ONLY.</td>
			</tr>
			<tr>
				<td colspan="2">
					<input type="radio" name="r1" value="1" class="clearBorder"> Delete registered customers' saved carts ONLY.
        		</td>
			</tr>
			<tr>
				<td colspan="2">
					<input type="radio" name="r1" value="2" class="clearBorder"> Delete all saved carts
        		</td>
			</tr>
			<tr> 
            	<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="2">
					<input name="submit2" type="submit" id="submit2" value=" Clean Up Database " class="submit2">
        		</td>
       		</tr>
          	<tr> 
            	<td colspan="2" class="pcCPspacer"></td>
			</tr>
          	<tr> 
              	<th colspan="2">Remove product-level statistics</th>
          	</tr>
			<tr> 
            	<td colspan="2" class="pcCPspacer"></td>
			</tr>
			<tr>
				<td colspan="2">
                	<%
					'// Count records in the table
					query="SELECT idProduct FROM pcSavedPrdStats "
					set rs3=Server.CreateObject("ADODB.Recordset")     
					rs3.Open query, conntemp
					if err.number <> 0 then
						set rs3=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in PurgeSavedCarts: "&Err.Description) 
					end if
					itemCount = 0
					if NOT rs3.eof then					
						do while NOT rs3.eof
							itemCount = itemCount + 1
						rs3.movenext
						loop
						set rs3=nothing
					end if
					%>
                
                	ProductCart saves data at the item level to give you <a href="viewSCLogs.asp">statistics on the products that are most commonly added to the car</a>t. The data is automatically cleared on a monhtly basis (when either "<em>pc/inc_SaveShoppingCart.asp</em>" or "<em>pcadmin/viewSCLogs.asp</em>" are run). So you normally don't need to use this tool, unless the table has grown particularly large and you need to reduce the database size. The table currently <strong>contains <%=itemCount%> records</strong>.
                    <br /><br />
					<input name="submit3" type="submit" id="submit3" value=" Clean Up Database " class="submit2">
        		</td>
       		</tr>				
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
		</table>
	</form>
<% end if
call closeDb()
%>
<!--#include file="AdminFooter.asp"-->