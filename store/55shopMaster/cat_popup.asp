<% response.Buffer=true 
Server.ScriptTimeout = 120 %>
<% PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<%
dim f, query, conntemp, rstemp, upCnt

Const iPageSize=250

Dim iPageCurrent
if request.querystring("iPageCurrent")="" then
	iPageCurrent=1 
else
	iPageCurrent=Request.QueryString("iPageCurrent")
end if

Dim tmpParent
Dim intCount
Dim pcArray
tmpParent=""

Function FindParent(idCat)
	Dim k
	if clng(idCat)<>1 then
		For k=0 to intCount
			if (clng(pcArray(0,k))=clng(idCat)) and (clng(pcArray(0,k))<>1)	then
				if tmpParent<>"" then
					tmpParent="/" & tmpParent
				end if
				tmpParent=pcArray(1,k) & tmpParent
				FindParent(pcArray(2,k))
				exit for
			end if
		Next
	end if
End function
%>

<html>
<head>
<title>Edit Category Assignment</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<% IF request("action")="update" THEN
	iCnt=request("iCnt")
	idproduct=request("idproduct")
	intTotCat=request("intTotCat")
	
	call opendb()

	For i=1 to iCnt
		if request("C" & i)="1" then
			query="select * from categories_products where idproduct=" & idproduct & " AND idcategory=" & request("ID" & i)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			if rs.eof then
				query="insert into categories_products (idproduct,idcategory) values (" & idproduct & "," & request("ID" & i) & ")"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=conntemp.execute(query)
				set rs=nothing
			end if
		else
			query="delete from categories_products where idproduct=" & idproduct & " AND idcategory=" & request("ID" & i)
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			set rs=nothing
		end if
	Next
	
	call closedb()
	%>
	<table class="pcCPcontent">
		<tr>
			<th>Edit Category Assignment</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>
			<p>This product has been assigned to categories listed in on the &quot;Modify Product&quot; page.</p>
				<% if request("show")="NotAssigned" then
					if request("iPageCurrent")<>request("iPageCount") then %>
					<p>If you would like to add this product to additional categories, <a href="cat_popup.asp?show=NotAssigned&iPageCurrent=<%=request("iPageCurrent")+1%>&idProduct=<%=idProduct%>">go to the next page of results</a>.</p>
					<% end if
				end if %>
			</td>
		</tr>
		<tr>
			<td>
				<p align="center"><a href="JavaScript:;" onClick="opener.location.reload(); self.close();">Close Window</a></td>
		</tr>
	</table>
<% ELSE 
	
	dim parent,iPageCount

	call opendb()
	
	function GetParent() 
		query="SELECT idparentCategory, categoryDesc FROM categories WHERE idCategory=" & idparentCategory
		set rsParentObj=server.CreateObject("ADODB.RecordSet")
		set rsParentObj=conntemp.execute(query)
		idparentCategory=rsParentObj("idparentCategory")
		parent=parent & "/" & rsParentObj("categoryDesc")
		set rsParentObj=nothing
		If idparentCategory<>1 then
			call GetParent() 
		end if
	End function
	
	
	
	idproduct=request("idproduct")
	intTotCat=request("intTotCat")
	query="select idCategory from categories_products where idproduct=" & idproduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	count1=0
	dim A(999)
	do while not rs.eof
		count1=count1+1
		A(count1-1)=rs("idcategory")
		rs.movenext
	loop 
	set rs=nothing
	
	if request("show")="Assigned" then %>
		<form name="form1" method="post" action="cat_popup.asp?action=update&show=Assigned&iPageCurrent=<%=iPageCurrent%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<th>Edit Category Assignment</th>
		</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
			<tr>
				<td>This product is currently assigned to the following categories:</td>
			</tr>
			<tr>
				<td>
					<table class="pcCPcontent" style="width:auto;">
					<tr>
						<th width="50%" colspan="2">Category</th>
						<th width="50%">Parent</th>
					</tr>
					<%
					query="SELECT categories.idCategory, categories.categoryDesc, categories.idParentCategory, categories_products.idProduct FROM categories INNER JOIN categories_products ON categories.idCategory = categories_products.idCategory WHERE categories_products.idProduct="&idproduct&" ORDER BY categories.categoryDesc, categories.idParentCategory;"
					' set the cache size=to the # of records/page
					Set rs=Server.CreateObject("ADODB.Recordset")
					
					rs.CacheSize=250
					rs.PageSize=250
					
					rs.Open query, connTemp, adOpenStatic, adLockReadOnly
					rs.MoveFirst

					' get the max number of pages
					iPageCount=rs.PageCount
					If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
					If iPageCurrent < 1 Then iPageCurrent=1
					
					' set the absolute page
					rs.AbsolutePage=iPageCurrent
					
					if err.number <> 0 then
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in cat_popup.asp: "&Err.Description) 
					end if
					
					Dim iCnt
					iCnt=0
					upCnt=0
	
					responseStr=""	
					Count=0
					Do While NOT rs.EOF And Count < rs.PageSize
						upCnt=upCnt+1
						iCnt=iCnt+1
						idCategory=rs("idCategory")
						categoryDesc=rs("categoryDesc")
						idparentCategory=rs("idparentCategory")
	
						if idparentCategory=1 then
							responseStr=responseStr&"<tr><td width='5%'><input type='checkbox' name='C"&iCnt&"' value='1'"
							for i=0 to count1-1
								if A(i)=idCategory then
									responseStr=responseStr&" checked"
								end if
							next
							responseStr=responseStr&" class='clearBorder'><input type=hidden name='ID"&iCnt&"' value='"&idCategory&"'></td><td>"&categoryDesc&"</font></td><td></td></tr>"
						else 
							query="SELECT idCategory, categoryDesc, idparentCategory FROM categories WHERE idCategory="&idparentCategory&";"
							set rstemp=server.CreateObject("ADODB.RecordSet")
							set rstemp=conntemp.execute(query)
							idparentCategory=rstemp("idparentCategory")
							parent=rstemp("categoryDesc")
							set rstemp=nothing
							if idparentCategory<>1 then
								call GetParent()	
							end if
							responseStr=responseStr&"<tr><td width='5%'><input type='checkbox' name='C"&iCnt&"' value='1'" 
							for i=0 to count1-1
								if A(i)=idCategory then
									responseStr=responseStr&" checked"
								end if
							next
							responseStr=responseStr&" class='clearBorder'><input type=hidden name='ID"&iCnt&"' value='"&idCategory&"'></td><td>"&categoryDesc&"</td><td>"&parent&"</td></tr>"
						end if
						count=count + 1
					  rs.MoveNext
					Loop
					set rs=nothing
					response.write responseStr %>
				<tr>
					<td>
					<Input type=hidden name=iCnt value="<%=iCnt%>">
					<Input type=hidden name=iPageCount value="<%=iPageCount%>">
					<Input type=hidden name=idproduct value="<%=idproduct%>">&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				</tr>
				<tr>
				  <td colspan="3">
	          <% Response.Write "Displaying Page <b>" & iPageCurrent & "</b> of <b>" & iPageCount & "</b>"%>
					</tr>
				<tr>
				  <td colspan="3">
				  <!-- Navigtion through pages -->
				  <%

					'Display Next / Prev buttons
					if iPageCurrent > 1 then
						'We are not at the beginning, show the prev button %>
								<a href="cat_popup.asp?show=<%=request("show")%>&iPageCurrent=<%=iPageCurrent-1%>&idProduct=<%=idProduct%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
					  
		      <% end If
		If iPageCount <> 1 then
			For I=1 To iPageCount
				If I=iPageCurrent Then
					response.write I
				Else %>
          <a href="cat_popup.asp?show=<%=request("show") %>&iPageCurrent=<%=I%>&idProduct=<%=idProduct%>"><%=I%></a> 
			    <% End If
			Next
		end if
		if CInt(iPageCurrent) <> CInt(iPageCount) then
			'We are not at the end, show a next link %>
		      <a href="cat_popup.asp?show=<%=request("show") %>&iPageCurrent=<%=iPageCurrent+1%>&idProduct=<%=idProduct%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
	        <% end If %>
				  <!--end Navigation through pages -->
			    </td>
			  </tr>
				<tr>
					<td colspan="3" align="center">
					<input type="submit" name="Submit" value="Update" class="submit2">
					&nbsp;<input type="button" name="Back" value="Back" onClick="JavaScript:history.back(-1)">
					&nbsp;<input type="button" name="Close" value="Close" onClick="self.close();">
					</td>
				</tr>
				<tr>
					<td colspan="3" align="center">&nbsp;</td>
				</tr>
			</table>      
			</td>
		</tr>
	</table>
	</form>
	<% else 
		if request("show")="NotAssigned" OR intTotCat=0 then %>
			<form name="form1" method="post" action="cat_popup.asp?action=update&show=NotAssigned&iPageCurrent=<%=iPageCurrent%>" class="pcForms">
		<table class="pcCPcontent">
			<tr>
				<td><h2>Edit Category Assignment</h2></td>
			</tr>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
				<% if intTotCat=0 then %>
				<tr>
					<td><p>Assign this product to one or more of the following categories:</p></td>
				</tr>
				<% else %>
				<tr>
					<td><p>This product is currently assigned to the following categories:</p></td>
				</tr>
				<% end if %>
				<tr>
					<td>
						<table class="pcCPcontent" style="width:auto;">
							<tr>
								<th width="50%" colspan="2">Category</th>
								<th width="50%">Parent</th>
							</tr>
						<%
						query="SELECT categories.idcategory,categories.categorydesc,categories.idParentCategory FROM categories ORDER BY categories.idcategory asc;"
						set rstemp=conntemp.execute(query)
						if err.number <> 0 then
						    call closeDb()
						    response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving categories from database: "&Err.Description) 
						end If
						if not rstemp.eof then
							pcArray=rstemp.getRows()
							intCount=ubound(pcArray,2)
						end if
						set rstemp=nothing

						query="SELECT idCategory, categoryDesc, idparentCategory FROM categories ORDER BY categoryDesc, idparentCategory;"
						' set the cache size=to the # of records/page
						Set rs=Server.CreateObject("ADODB.Recordset")
						
						rs.CacheSize=250
						rs.PageSize=250
	
						rs.Open query, connTemp, adOpenStatic, adLockReadOnly
						rs.MoveFirst
	
						' get the max number of pages
						iPageCount=rs.PageCount
						If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
						If iPageCurrent < 1 Then iPageCurrent=1
						
						' set the absolute page
						rs.AbsolutePage=iPageCurrent
						
						if err.number <> 0 then
							set rs=nothing
							call closedb()
							response.redirect "techErr.asp?error="& Server.Urlencode("Error in cat_popup.asp: "&Err.Description) 
						end if
						
						iCnt=0
		
						responseStr=""	
						Count=0
						
						if NOT rs.eof then
							CategoryArray=rs.getRows()
							intCatCount=ubound(CategoryArray,2)
						end if
						set rs=nothing
						
						if intCatCount>250 then
							intCatCount=250
						end if
						
						For m=0 to intCatCount
							iCnt=iCnt+1
							idCategory=CategoryArray(0,m)
							categoryDesc=CategoryArray(1,m)
							idparentCategory=CategoryArray(2,m)
		
							if idparentCategory=1 then
								responseStr=responseStr&"<tr><td width='5%'><input type='checkbox' name='C"&iCnt&"' value='1'"
								for i=0 to count1-1
									if A(i)=idCategory then
										responseStr=responseStr&" checked"
									end if
								next
								responseStr=responseStr&" class='clearBorder'><input type=hidden name='ID"&iCnt&"' value='"&idCategory&"'></td><td>"&categoryDesc&"</td><td>&nbsp;</td></tr>"
							else
								tmpParent=""
								FindParent(idparentCategory)
								responseStr=responseStr&"<tr><td width='5%'><input type='checkbox' name='C"&iCnt&"' value='1'" 
								for i=0 to count1-1
									if A(i)=idCategory then
										responseStr=responseStr&" checked"
									end if
								next
								responseStr=responseStr&" class='clearBorder'><input type=hidden name='ID"&iCnt&"' value='"&idCategory&"'></td><td>"&categoryDesc&"</td><td>"
								if tmpParent<>"" then
								responseStr=responseStr&tmpParent
								end if
								responseStr=responseStr&"</td></tr>"
							end if 
						count=count + 1
					Next
					response.write responseStr%>
					<tr>
						<td>
						<Input type=hidden name=iCnt value="<%=iCnt%>">
						<Input type=hidden name=iPageCount value="<%=iPageCount%>">
						<Input type=hidden name=idproduct value="<%=idproduct%>">&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
            <td colspan="3">
            <% Response.Write "Displaying Page <b>" & iPageCurrent & "</b> of <b>" & iPageCount & "</b></td>" %>
          </tr>
					<tr>
            <td colspan="3">
							<!-- Navigtion through pages -->
							<%	'Display Next / Prev buttons
							if iPageCurrent > 1 then
								'We are not at the beginning, show the prev button %>
								<a href="cat_popup.asp?show=<%=request("show")%>&iPageCurrent=<%=iPageCurrent-1%>&idProduct=<%=idProduct%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a>
							<% end If
							If iPageCount <> 1 then
								For I=1 To iPageCount
									If I=iPageCurrent Then
										response.write I
									Else %>
										<a href="cat_popup.asp?show=<%=request("show") %>&iPageCurrent=<%=I%>&idProduct=<%=idProduct%>"><%=I%></a>
									<% End If
								Next
							end if
							if CInt(iPageCurrent) <> CInt(iPageCount) then
								'We are not at the end, show a next link %>
								<a href="cat_popup.asp?show=<%=request("show") %>&iPageCurrent=<%=iPageCurrent+1%>&idProduct=<%=idProduct%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a>
							<% end If %>
							<!--end Navigation through pages -->
            </td>
				  </tr>
					<tr>
						<td colspan="3" align="center">
						<input type="submit" name="Submit" value="Update" class="submit2">
						&nbsp;<input type="button" name="Back" value="Back" onClick="JavaScript:history.back(-1)">
						&nbsp;<input type="button" name="Close" value="Close" onClick="self.close();">
						</td>
					</tr>
					<tr>
						<td colspan="3">&nbsp;</td>
					</tr>
					</table>      
					</td>
				</tr>
			</table>
			</form>
		<% else
			'show choices %>
			<table class="pcCPcontent">
				<tr>
					<th>Edit Category Assignment</th>
				</tr>
				<tr>
					<td class="pcCPspacer"></td>
				</tr>
				<tr>
					<td><p>Please select one of the following options. <strong>NOTE:</strong> If your store includes a high number of categories and/or category levels, the second option may require up to a few minutes to execute.</p></td>
				</tr>
				<tr>
				  <td>
					<form action="cat_popup.asp" method="get" name="form" id="form" class="pcForms">
				    <p><input type="radio" name="show" value="Assigned" class="clearBorder"> <strong>Review</strong> current category assignment (faster)</p>
						<p><input type="radio" name="show" value="NotAssigned" class="clearBorder"> <strong>Add/remove</strong> product from any category (slower)</p>
						<p>&nbsp;</p>
						<p>
					  <input type='hidden' name="idproduct" value="<%=idproduct%>">
					  <input name="submit" type="submit" value="Continue" class="submit2">
		        </p>  
				  </form>
				</td>
		  </tr>
			</table>
		<% end if %>
	<% end if
	call closedb() %>
<%END IF%>
</div>
</body>
</html>
