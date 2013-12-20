<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin="11*12*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../pc/pcSeoFunctions.asp"-->

<%

dim query, conntemp, rs, pcvParentPageName, pcArray, pcInt_Parent, pcInt_Published, pcInt_Inactive, pcInt_Active, pcv_IntNumrowsCount, pcv_IntNumrows, m, pcInt_AlertStoreManager, pcInt_DraftPresent

pcInt_Parent=request("parent")
if not validNum(pcInt_Parent) then pcInt_Parent = 0
if pcInt_Parent>0 then
	' Load Parent Page Name
	call openDb()
	query="SELECT pcCont_PageName FROM pcContents WHERE pcCont_IDPage="&pcInt_Parent
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	pcvParentPageName=rs("pcCont_PageName")
	pageTitle="Manage Content Pages under " & pcvParentPageName
	query1=" WHERE pcCont_Parent="&pcInt_Parent
	set rs=nothing
	call closeDb()
else
	pcvParentPageName=""
	pageTitle="Manage Content Pages"
end if

%>

<!--#include file="AdminHeader.asp"-->

<%
call openDb()

'// START - Determine the type of user
pcInt_LimitedUser=0
if session("PmAdmin") <> "19" and (not isNull(findUser(pcUserArr,12,pcUserArrCount))) then
	pcInt_LimitedUser=1
end if
'// END - Determine the type of user

IF request("submit1")<>"" OR request("submit2")<>"" THEN

	pcv_IntNumrowsCount=request("IntNumrowsCount")
	pcInt_Parent=request.form("parent")
	if not validNum(pcInt_Parent) then pcInt_Parent = 0
	
	IF validNum(pcv_IntNumrowsCount) then
	
	For k=1 to clng(pcv_IntNumrowsCount)
	
		if request("CT"&k)="1" then
			pcv_id=request("CT"&k&"_id")
			pcInt_Active=request("active"&k)
			pcInt_priority=request.form("priority"&k)
			pcInt_Published=request.form("published"&k)
			
			if not validNum(pcInt_Active) then pcInt_Active=0
			if pcInt_Active="0" then
				pcInt_Inactive="1"
			else
				pcInt_Inactive="0"
			end if
			
			if not validNum(pcInt_Published) then pcInt_Published=0
			
			'// UPDATE Selected Pages
			if request("submit1")<>"" then
				query="UPDATE pcContents SET pcCont_InActive=" & pcInt_Inactive & ", pcCont_Order=" & pcInt_priority & ", pcCont_Published=" & pcInt_Published & " WHERE pcCont_IDPage=" & pcv_id  
				set rstemp=Server.CreateObject("ADODB.Recordset")
				set rstemp=conntemp.execute(query)
				if err.number <> 0 then
					strErrDescription = Err.Description
					set rstemp = nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error: "& strErrDescription) 
				else
					msg="Content Pages updated successfully"
				end if
			end if
			
			'// DELETE Selected Pages
			if request("submit2")<>"" then
				query="DELETE FROM pcContents WHERE pcCont_IDPage=" & pcv_id  
				set rstemp=conntemp.execute(query)
				if err.number <> 0 then
					strErrDescription = Err.Description
					set rstemp = nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error: "& strErrDescription) 
				else
					msg="Content Pages deleted successfully"
				end if
			end if
		end if	
	
	Next
	
	END IF
	
	set rstemp=nothing
	call closeDb()
	response.Redirect "cmsManage.asp?s=1&msg=" & server.URLEncode(msg)
	
END IF

' Load Pages
query="SELECT pcCont_IDPage, pcCont_PageName, pcCont_InActive, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_DraftStatus FROM pcContents" & query1 & " ORDER BY pcCont_Order, pcCont_Parent, pcCont_PageName ASC;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error loading content pages") 
end If

'// START - TO DO items
	if pcInt_LimitedUser=0 then
		' //Check to see if any pages need to be reviewed
		pcInt_AlertStoreManager=0
		query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_Published=0 or pcCont_Published=2;"
		set rstempCheck=server.CreateObject("ADODB.RecordSet")
		set rstempCheck=conntemp.execute(query)
		if not rstempCheck.eof then pcInt_AlertStoreManager=1
		if pcInt_AlertStoreManager=1 then
		%>
			<div class="pcCPmessage">One or more Content Pages need to be <strong>reviewed</strong>.<br><em>It's the pages for which the checkbox in the 'Pub' column below is not checked.</em></div>
		<%
		end if
	
		' Check to see if any pages need to be reviewed
		pcInt_AlertStoreManager=0
		query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_DraftStatus=1;"
		set rstempCheck=conntemp.execute(query)
		if not rstempCheck.eof then pcInt_AlertStoreManager=1
		set rstempCheck=nothing
		if pcInt_AlertStoreManager=1 then
		%>
			<div class="pcCPmessage">One or more Content Pages have a <strong>draft</strong> saved to the database, which might need to be completed, reviewed, and published.<br><em>It's the pages for which the column 'Draft' says 'Yes' below.</em></div>
		<%
		end if
	end if
'// END - TO DO items
%>

<script language="JavaScript">
<!--
    function chgWin(file,window) {
    msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
    if (msgWindow.opener == null) msgWindow.opener = self;
}
//-->
</script>

<%
'Check URL

	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
	mycount1=mycount1+1
	end if
	if mycount1<2 then
	SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
	loop
	SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

%>

<form name="form1" action="cmsManage.asp" method="post" class="pcForms">
	<table class="pcCPcontent">
        <tr>
            <td colspan="8" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
        	<th></th>
        	<th>Order</th>
			<th>Active</th>
 			<th>Pub</th>
            <th>Draft</th>
			<th>Name</th>
            <th colspan="2">
            <div style="float: right;">
				<%
                ' Load parent pages - Start
				if pcInt_Parent>0 then
				%>
                <span class="cpLinksList"><a href="cmsManage.asp">Show All</a></span>
                <%
				else
                    Dim pcPageParentExist, intPageCount
                    query="SELECT pcCont_idPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 ORDER BY pcCont_PageName ASC"
                    set rs=Server.CreateObject("ADODB.Recordset")
                    set rs=connTemp.execute(query)
                    if rs.EOF then
                        pcPageParentExist=0
                    else
                        pcPageParentExist=1
                        pcPageArr=rs.getRows()
                    end if
                    set rs=nothing
                        if pcPageParentExist=1 then
                        %>
                            <select name="Parent" tabindex="104" onChange="this.form.submit()">
                                <option value="0">Only show pages under...</option>
                        <%
                            intPageCount=ubound(pcPageArr,2)
                            For m=0 to intPageCount 
								query="SELECT pcCont_idPage FROM pcContents WHERE pcCont_Parent="&pcPageArr(0,m)
								set rsParent=Server.CreateObject("ADODB.Recordset")
								set rsParent=connTemp.execute(query)
								if not rsParent.eof then
							%>
                                <option value="<%=pcPageArr(0,m)%>"><%=pcPageArr(1,m)%></option>
                        <%
								end if
								set rsParent=nothing
                            Next
                        %>
                            </select>
                    <%
                    end if
				end if
                ' Load parent pages - End
                %>
            	</div>
                Parent
            </th>
		</tr>
		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<%
		'-------------------------
		' NO Content Pages Found
		'-------------------------
		If rstemp.eof Then
			set rstemp=nothing
			call closeDb()
		%>
			<tr> 
				<td colspan="8" align="center">
					<div class="pcCPmessage">No Content Pages Found. <a href="cmsAddEdit.asp">Add New</a>&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=436')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></div>
				</td>
			</tr>                  
		<% 
		'-------------------------
		' LIST Content Pages
		'-------------------------
		Else 
		
			pcArray = rstemp.getRows()
			set rstemp=nothing
			call closeDb()
			
			pcv_IntNumrows = UBound(pcArray, 2)

			pcv_IntNumrowsCount=0
			FOR m = 0 to pcv_IntNumrows

				pcv_IntNumrowsCount=pcv_IntNumrowsCount+1
				pcv_lngIDPage= pcArray(0,m)
				pcv_strPageName = pcArray(1,m)
				pcInt_Inactive = pcArray(2,m)
				pcInt_priority = pcArray(3,m)
				pcInt_Parent = pcArray(4,m)
				pcInt_Published = pcArray(5,m)
				pcInt_DraftPresent = pcArray(6,m)
				
				'// SEO Links
				'// Build Navigation Page Link
				if scSeoURLs=1 then
					pcStrCntPageLink=pcv_strPageName & "-d" & pcv_lngIDPage & ".htm"
					pcStrCntPageLink=removeChars(pcStrCntPageLink)
					pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink &"?adminPreview=1"
				else
					pcStrCntPageLink="../pc/viewContent.asp?idpage="&pcv_lngIDPage&"&adminPreview=1"
				end if
				
				'// Change links if this is a parent page
				Dim intAlreadyParent
				call openDb()
				query="SELECT pcCont_idPage FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage
				set rsParent=Server.CreateObject("ADODB.Recordset")
				set rsParent=connTemp.execute(query)
				if rsParent.EOF then
					intAlreadyParent=0
				else
					intAlreadyParent=1
				end if
				if intAlreadyParent=1 then 
					pcStrCntPageLink="../pc/viewPages.asp?idpage=" & pcv_lngIDPage
				end if
				set rsParent=nothing
				call closeDb()
				'//
				
				if not validNum(pcInt_Inactive) then
					pcInt_Inactive=0
				end if
				
				if not validNum(pcInt_Parent) then
					pcInt_Parent=0
				end if
				
				if not validNum(pcInt_Published) then
					pcInt_Published=0
				end if
				
				if not validNum(pcInt_DraftPresent) then
					pcInt_DraftPresent=0
				end if
				
			%>           
				<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
					<td>
						<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
                            <input name="ct<%=pcv_IntNumrowsCount%>" type="checkbox" value="1" class="clearBorder">
                        <% end if %>
						<input type="hidden" name="ct<%=pcv_IntNumrowsCount%>_id" value="<%=pcv_lngIDPage%>">
					</td>
                    <td>
						<input type="text" name="priority<%=pcv_IntNumrowsCount%>" size="2" maxlength="4" value="<%=pcInt_priority%>"<%if pcInt_LimitedUser=1 then%> disabled<%end if%>>
                    </td>
					<td>
						<input type="checkbox" name="active<%=pcv_IntNumrowsCount%>" value="1" <%if pcInt_Inactive="0" then%>checked<%end if%><%if pcInt_LimitedUser=1 then%> disabled <%end if%> class="clearBorder">
					</td>
					<td>
						<input type="checkbox" name="published<%=pcv_IntNumrowsCount%>" value="1" <%if pcInt_Published="1" then%>checked<%end if%><%if pcInt_LimitedUser=1 then%> disabled <%end if%> class="clearBorder">
					</td>
					<td>
						<% if pcInt_DraftPresent<>0 then %>Yes<% else %>No<% end if %>
					</td>
					<td width="40%">
						<a href="cmsAddEdit.asp?idpage=<%=pcv_lngIDPage%>"><%=pcv_strPageName%></a>
					</td>
                    <td width="40%">
                    <% 
					if pcInt_Parent > 0 then
						call openDb()
						query="SELECT pcCont_PageName FROM pcContents WHERE pcCont_IDPage=" & pcInt_Parent
						set rstemp = Server.CreateObject("ADODB.Recordset")
						set rstemp = conntemp.execute(query)
						if not rstemp.eof then
						pcv_ParentPageName = rstemp("pcCont_PageName")
						else
						pcv_ParentPageName = "N/A"
						end if
						set rstemp = nothing
						call closeDb()
					%>
                    	<a href="cmsAddEdit.asp?idpage=<%=pcInt_Parent%>"><%=pcv_ParentPageName%></a>
                    <%
					end if
					%>
                    </td>
					<td nowrap align="right">
                    	<a href="cmsAddEdit.asp?idpage=<%=pcv_lngIDPage%>"><img src="images/pcIconGo.jpg" border="0" alt="Edit" title="Edit this Content Page"></a>&nbsp;
                        <a href="<%=pcStrCntPageLink%>" target="_blank"><img src="images/pcIconPreview.jpg" border="0" alt="Preview" title="Preview this Content Page"></a>
					</td>
				</tr>
				<%
				NEXT
				%>
				<tr>
					<td colspan="8" class="pcCPspacer"></td>
				</tr>
                <% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<tr>
					<td colspan="8">
                    	<% if pcv_IntNumrowsCount<>"" then %>
						<script language="JavaScript">
                            <!--
                            function checkAll() {
                            for (var j = 1; j <= <%=pcv_IntNumrowsCount%>; j++) {
                            box = eval("document.form1.ct" + j); 
                            if (box.checked == false) box.checked = true;
                                 }
                            }
                            
                            function uncheckAll() {
                            for (var j = 1; j <= <%=pcv_IntNumrowsCount%>; j++) {
                            box = eval("document.form1.ct" + j); 
                            if (box.checked == true) box.checked = false;
                                 }
                            }
                            //-->
                        </script>
						<%end if%>
						<span class="cpLinksList"><a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a></span>
					</td>
				</tr>	
		<%
				end if
			END IF
		'-------------------------
		' END listing content pages
		'-------------------------
		%>					

		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="center" colspan="8">
            	<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<input name="submit1" type="submit" value="Update Selected" class="submit2">&nbsp;
				<input name="submit2" type="submit" value="Delete Selected" onclick="return(confirm('You are about to remove selected content pages from your database. Are you sure you want to complete this action?'));">&nbsp;
               	<input type="hidden" name="IntNumrowsCount" value="<%=pcv_IntNumrowsCount%>">
                <% end if %>
				<input type="button" value="Add New" onclick="location='cmsAddEdit.asp';">&nbsp;
				<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<input type="button" value="Generate Navigation" onclick="location='cmsNavigation.asp';">&nbsp;
                <% end if %>
                <input type="button" value="Browse Pages" onclick="window.open('../pc/viewPages.asp');">&nbsp;
                <input type="button" value="Help" onclick="window.open('http://wiki.earlyimpact.com/productcart/settings-content-pages');">&nbsp;
				<input type="button" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->