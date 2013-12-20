<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle = "Welcome to Your Store's Control Panel" 
pageIcon = ""
pcStrPageName = "menu.asp"
%>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/UpdateVersionCheck.asp"-->
<!--#include file="../includes/PPDStatus.inc"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"--> 
<!--#include file="../includes/SQLFormat.txt"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/GoogleCheckoutConstants.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="AdminHeader.asp"-->
<script language="JavaScript">
<!--
	function chgWin(file,window) {
	msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=650,height=380');  
	if (msgWindow.opener == null) msgWindow.opener = self;
	}
//-->
</script>

<%dim query, conntemp, rs

Dim pcvShowCharts

pcvShowCharts=1

call opendb()%>
<!--#include file="pcCharts.asp"-->
<%

'// START Check Saved Carts table to see if contains too many records
query="SELECT Count(*) As SCTotal FROM pcSavedCarts;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)
if not rs.eof then
	tmpTotal=cLng(rs("SCTotal"))
	if tmpTotal>10000 then
		msg="The Saved Shopping Carts table exceeds 10,000 records<br><a href='PurgeSavedCarts.asp'>Click here</a> if you want to clear it."
	end if
end if
set rs=nothing
'// END

'// START Check Product Reviews to see if there are Pending reviews that need attention	
if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("PmAdmin")="19") then
	query="SELECT pcRev_IDReview,pcRev_Date FROM pcReviews WHERE pcRev_Active=0 AND pcRev_IDProduct<>0"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	if not rs.eof then
		msg="There are pending <strong>product reviews</strong> awaiting approval. <a href=prv_ManageRevPrds.asp?nav=1>Manage reviews &gt;&gt;</a>"
	end if
	set rs=nothing
end if
'// END
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<% 


'****************************
'* Version Update Tasks
'****************************
	if trim(updDBScript)<>"" then
		if updtrigger=1 then
				response.redirect updDBScript
		end if
	end if
	
'****************************
'* First Install
'****************************
	if scCompanyName="" then
			response.redirect "AdminSettings.asp?tab=1"
	end if

'****************************
'* Admin login: show info
'****************************
'Check admin rights
call openDb()
tmpRights=0
query="SELECT IDPm FROM Permissions ORDER BY IDPm ASC;"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
tmpRList=""
if not rstemp.eof then
	tmpArr=rstemp.getRows()
	intCount=ubound(tmpArr,2)
	set rstemp=nothing
	For i=0 to intCount
		if tmpArr(0,i)<>"12" then
			tmpRList=tmpRList & tmpArr(0,i) & "*"
		end if
	Next
end if
set rstemp=nothing

query="SELECT Id FROM Admins WHERE ((adminLevel LIKE '" & tmpRList & "') OR (adminLevel LIKE '" & replace(tmpRList,"*11*","*12*") & "')) AND ID=" & session("IDAdmin") & ";"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if not rstemp.eof then
	tmpRights=1
end if
set rstemp=nothing

IF (session("PmAdmin")="19") OR (tmpRights=1) THEN %> 
	<table class="pcCPcontent">
		<tr>
			<td width="55%" valign="top">
			<% 'FIRST CELL: last 10 orders %>
            <div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image:url(images/pcv4_icon_sales.gif); background-position: 10px -10px; background-repeat:no-repeat; min-height: 180px;">
				<h2 style="padding-left: 60px;">Most Recent Orders <span class="pcSmallText">&nbsp;|&nbsp;<a href="resultsAdvancedAll.asp?B1=View+All&dd=1">All Orders</a>&nbsp;|&nbsp;<a href="invoicing.asp">Find an Order</a></span></h2>
				<%
                call openDb()
                query="SELECT TOP 10 idorder, orderDate, total, idcustomer, rmaCredit, orderstatus, pcOrd_PaymentStatus FROM orders WHERE orderStatus>1 AND orderStatus<>5 ORDER BY idorder DESC"
                set rs=Server.CreateObject("ADODB.Recordset")
                set rs=conntemp.execute(query)
                if rs.EOF then %>
                <p><%=dictLanguageCP.Item(Session("language")&"_cpSearch_17")%></p>
                <% else %>
                <%
                    While Not rs.EOF and count<10
                    pIdOrder=rs("idorder")
                    pOrderDate=rs("orderDate")
                    pOrderDate=ShowDateFrmt(pOrderDate)
                    pTotal=rs("total")
                        pc_rmaCredit=rs("rmaCredit")
                        if trim(pc_rmaCredit)="" or IsNull(pc_rmaCredit) then
                            pc_rmaCredit=0
                        end if
                        pTotal=pTotal-pc_rmaCredit
                    porderstatus=rs("orderstatus")
                    pcv_PaymentStatus=rs("pcOrd_PaymentStatus")
                        if IsNull(pcv_PaymentStatus) or pcv_PaymentStatus="" then
                            pcv_PaymentStatus=0
                        end if
                    pIdCustomer=rs("idcustomer")
                    query="SELECT name,lastname FROM customers WHERE idcustomer="& pIdCustomer
                    Set rsCust=CreateObject("ADODB.Recordset")
                    set rsCust=conntemp.execute(query)
                    CustName=rsCust("name")& " "&rsCust("lastname")
                    set rsCust=nothing
                    %>
                        <div onMouseOver="this.className='activeRow'" onMouseOut="this.className='pcLinkRow'" style="padding: 3px;" class="pcLinkRow">
                        <a href="Orddetails.asp?id=<%=pIdOrder%>">
                        <!--#include file="inc_orderStatusIcons.asp"--> 
                        <%=(scpre + int(pIdOrder))%> <%=dictLanguageCP.Item(Session("language")&"_cpCommon_300")%> <%=CustName%>&nbsp;|&nbsp;<%=pOrderDate%>&nbsp;|&nbsp;<%=scCurSign & money(pTotal)%></a></div>
                    <%
                        rs.MoveNext
						count=count + 1
                    Wend
                end if
                Set rs=Nothing 
                call closeDb()
				count=0
                %>
            </div>
			</td>
			<% 
			' END FIRST CELL: last 10 orders
			' SECOND CELL: most recent customers
			%>
			<td width="45%" valign="top">
            	<div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image: url(images/pcv4_icon_people.png); background-position: 10px -10px; background-repeat:no-repeat; min-height: 180px;">
					<h2 style="padding-left: 60px;">Most Recent Customers <span class="pcSmallText">&nbsp;|&nbsp;<a href="viewcusta.asp">Find a customer</a></span></h2>
						<% 
						call openDb()
						
						query="SELECT TOP 10 idcustomer,LastName,[name],customerCompany,email FROM customers WHERE email<>'REMOVED' ORDER BY pcCust_DateCreated DESC;"
						set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=conntemp.execute(query)
						IF rstemp.EOF THEN
						%>
                        There are no customers in the database.
                        <%
						ELSE
						
							Do While NOT rstemp.EOF AND count<10
								pLastName=rstemp("LastName")
								pname=rstemp("name")
								pcustomerCompany=trim(rstemp("customerCompany"))
								pidcustomer=rstemp("idcustomer")
								pemail=rstemp("email")
						%>
	                        <div onMouseOver="this.className='activeRow'" onMouseOut="this.className='pcLinkRow'" style="padding: 3px;" class="pcLinkRow">
								<a href="modCusta.asp?idcustomer=<%=pidcustomer%>"><%=pname & " " & pLastName%>
                                <% if pcustomerCompany<>"" then%>&nbsp;|&nbsp;<%=pcustomerCompany%><% end if %></a>
							</div>
                            
                        <%
							count=count + 1
							rstemp.MoveNext
							loop

						END IF
						set rstemp=nothing
						call closeDb()
						count=0
						%>	
                 </div>
			</td>
			<% 
			' END SECOND CELL: most recent customers
			%>
		</tr>
	</table>
	<table class="pcCPcontent">    
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<%call opendb()
		tmpYear=Year(Date())
		query="SELECT TOP 1 idorder FROM Orders WHERE OrderStatus>=2 AND Year(OrderDate)=" & tmpYear & ";"
		set rs=connTemp.execute(query)
		if (not rs.eof) AND (pcvShowCharts=1) then
		set rs=nothing%>
		<tr>
			<td width="100%" valign="top" colspan="2">     
            <div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image:url(images/pcv4_icon_sales.gif); background-position: 10px -10px; background-repeat:no-repeat; min-height: 142px; overflow: auto;">
				<h2 style="padding-left: 60px;"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_232")%> <span class="pcSmallText">&nbsp;|&nbsp;<a href="dashboard.asp">Sales Charts</a>&nbsp;|&nbsp;<a href="srcOrdByDate.asp">Other Sales Reports</a></span></h2>
				<table class="pcCPcontent">
					<tr>
						<td colspan="2">
							<div id="chartMonthlySales" style="height:250px;"></div>
							<%Dim pcv_YearTotal
							pcv_YearTotal=0
							call pcs_MonthlySalesChart("chartMonthlySales",Year(Date()),0,1)%>
						</td>
					</tr>
					<%if pcv_YearTotal>0 then%>
					<tr> 
						<td colspan="2" class="pcCPspacer"></td>
					</tr>
					<tr> 
						<td colspan="2"> 
							<%=dictLanguageCP.Item(Session("language")&"_cpCommon_231")%>: <b><%=scCurSign & money(pcv_YearTotal)%></b>
						</td>
					</tr>
					<%end if%>
										
						<% 
						call openDb()
						
						totalyear=0
						
						query="SELECT year(orderdate) AS yearsql FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) GROUP BY year(orderdate) ORDER BY year(orderdate) DESC;"
						set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=conntemp.execute(query)
						
						stryear=""
						do until rstemp.eof 
							yearvalue=rstemp("yearsql")
							if clng(yearvalue)<>clng(year(now())) then
								stryear=stryear & yearvalue & "***"
								totalyear=totalyear+1
							end if   
							rstemp.movenext
						loop
						set rstemp=nothing
						
						if totalyear>0 then
						%>
							<tr>
								<td colspan="2">
								<%=dictLanguageCP.Item(Session("language")&"_cpCommon_230")%>: &nbsp;
								<%
								Ayear=split(stryear,"***")
								For dd=1 to totalyear
								%>
								<a href="#" onClick="chgWin('salescharts.asp?year=<%=Ayear(dd-1)%>','window2')"><%=Ayear(dd-1)%></a>
								<%
								If dd <> totalyear Then Response.Write " - " End if
								Next
								%>
								</td>
							</tr>
						<%
						end if
						%>	
				</table>
                </div>
			</td>
		</tr>
		<%
		end if
		set rs=nothing
		call closedb()
		%>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<% 
			' START FIRST CELL: most recent products
			%>
			<td valign="top" width="50%">
            
				<div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image:url(images/pcv4_icon_inventoryAdded.gif); background-position: 10px -10px; background-repeat:no-repeat; min-height: 200px;">
					<h2 style="padding-left: 60px;">Recently Added Products <span class="pcSmallText">&nbsp;|&nbsp;<a href="locateproducts.asp">Find a product</a></span></h2>
						
						<% 
						call openDb()
						query="SELECT TOP 5 idproduct,description,smallImageUrl,pcprod_EnteredOn FROM products WHERE removed<>-1 AND NOT (pcprod_EnteredOn IS NULL) ORDER BY pcprod_EnteredOn DESC;"
						set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=conntemp.execute(query)
						IF rstemp.EOF THEN
						%>
                        There are no products in the database.
                        <%
						ELSE
							count=0
							Do Until count=5 OR rstemp.eof
								pcStrProductName=rstemp("description")
								pcIntProductID=rstemp("idproduct")
								pcvStrSmallImage = rstemp("smallImageUrl")
								if pcvStrSmallImage = "" or pcvStrSmallImage = "no_image.gif" then
									pcvStrSmallImage = "hide"
								end if
								
								' Image size
								pcIntSmImgWidth = 20
								pcIntSmImgHeight = 20
						%>
                        
	                        <div onMouseOver="this.className='activeRow'" onMouseOut="this.className='pcLinkRow'" style="padding: 3px;" class="pcLinkRow">
                            	<% if pcvStrSmallImage <> "hide" then %><a href="findProductType.asp?idproduct=<%=pcIntProductID%>"><img src="../pc/catalog/<%=pcvStrSmallImage%>" width="<%=pcIntSmImgWidth%>" height="<%=pcIntSmImgHeight%>" align="middle" style="border:none; padding-top: 2px; padding-bottom: 5px; padding-right: 4px;"></a><% end if %>
								<a href="findproducttype.asp?idproduct=<%=pcIntProductID%>"><%=pcStrProductName%></a>
							</div>
                            
                        <%
							count=count + 1
							rstemp.MoveNext
							loop

						END IF
						set rstemp=nothing
						call closeDb()

						%>	
                 </div>
                 
			</td>
			<% 
			' END FIRST CELL: most recent products
			' SECOND CELL: most recent product updates
			%>
			<td valign="top" width="50%">
				<div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image:url(images/pcv4_icon_inventoryAdded.gif); background-position: 10px -10px; background-repeat:no-repeat; min-height: 200px;">
					<h2 style="padding-left: 60px;">Recently Edited Products <span class="pcSmallText">&nbsp;|&nbsp;<a href="locateproducts.asp">Find a product</a></span></h2>
						
						<% 
						call openDb()
						query="SELECT TOP 5 idproduct,description,smallImageUrl FROM products WHERE removed<>-1 AND NOT (pcprod_EditedDate IS NULL) ORDER BY pcprod_EditedDate DESC;"
						set rstemp=Server.CreateObject("ADODB.Recordset")
						set rstemp=conntemp.execute(query)
						IF rstemp.EOF THEN
						%>
                        The database does not yet contain information on recently edited products. As you edit products, a list of the most recently edited products will be displayed here.
                        <%
						ELSE
							count=0
							Do Until count=5 OR rstemp.eof
								pcStrProductName=rstemp("description")
								pcIntProductID=rstemp("idproduct")
								pcvStrSmallImage = rstemp("smallImageUrl")
								if pcvStrSmallImage = "" or pcvStrSmallImage = "no_image.gif" then
									pcvStrSmallImage = "hide"
								end if
								
								' Image size
								pcIntSmImgWidth = 20
								pcIntSmImgHeight = 20
						%>
                        
	                        <div onMouseOver="this.className='activeRow'" onMouseOut="this.className='pcLinkRow'" style="padding: 3px;" class="pcLinkRow">
                            	<% if pcvStrSmallImage <> "hide" then %><a href="findProductType.asp?idproduct=<%=pcIntProductID%>"><img src="../pc/catalog/<%=pcvStrSmallImage%>" width="<%=pcIntSmImgWidth%>" height="<%=pcIntSmImgHeight%>" align="middle" style="border:none; padding-top: 2px; padding-bottom: 5px; padding-right: 4px;"></a><% end if %>
								<a href="findproducttype.asp?idproduct=<%=pcIntProductID%>"><%=pcStrProductName%></a>
							</div>
                            
                        <%
							count=count + 1
							rstemp.MoveNext
							loop

						END IF
						set rstemp=nothing
						call closeDb()

						%>	
                 </div>
			</td>
			<% 
			' END SECOND CELL: most recent product updates
			%>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>

		<%
        ' START - Recent Help Desk Messages
        IF scShowHD<>0 THEN ' Check if help desk is turned on/off
        %>
        <tr>
        	<td colspan="2">
            
            <div class="pcCProundBox" style="border: 1px solid #e1e1e1; padding: 5px; background-image:url(images/pcv4_icon_helpDesk.png); background-position: 10px -10px; background-repeat:no-repeat; min-height: 200px;">
            <h2 style="padding-left: 60px;">Recent Help Desk Messages <span class="pcSmallText">&nbsp;|&nbsp;<a href="adminFBsettings.asp">View Messages in a Date Range</a></span></h2>
            <%

                call openDb() 
                
                dim A(30,2),Count,FCount,k
                
                query="Select pcFStat_IDStatus,pcFStat_Name from pcFStatus"
                set rstemp=Server.CreateObject("ADODB.Recordset")
                set rstemp=connTemp.execute(query)
                
                Count=0
                do while not rstemp.eof
                    Count=Count+1
                    A(Count-1,0)=rstemp("pcFStat_IDStatus")
                    A(Count-1,1)=rstemp("pcFStat_Name")
                    rstemp.movenext
                loop
                
                redim B(Count-1)
                
                query="SELECT pcComm_FStatus FROM pcComments WHERE pcComm_IDParent=0" & query1
                set rstemp=connTemp.execute(query)
                
                FCount=0
                do while not rstemp.eof
                    FCount=FCount+1
                    For k=0 to Count-1
                        if cint(rstemp("pcComm_FStatus"))=cint(A(k,0)) then
                            B(k)=B(k)+1
                        end if
                    Next
                    rstemp.Movenext
                loop
                
                set rstemp = nothing
                
                %>
            <table class="pcCPcontent">
                <tr style="border-bottom: 1px dashed #CCC;">
                    <td nowrap valign="middle"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_n")%></td>
                    <td nowrap valign="middle">Order</td>
                    <td nowrap>Priority</td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_p")%></td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_q")%></td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_r")%></td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_s")%></td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_t")%></td>
                    <td nowrap><%response.write dictLanguage.Item(Session("language")&"_viewPostings_u")%></td>
                </tr>
                <tr class="main">
                    <td colspan="7" class="pcCPspacer"></td>
                </tr>
            <%
            
            Dim SOrder,SSort,APageCount,strsortOrder
            
            query="SELECT TOP 5 * FROM pcComments WHERE pcComments.pcComm_IDParent=0 ORDER BY pcComm_editedDate DESC"
            Set rstemp=Server.CreateObject("ADODB.Recordset")
            
            Dim iPageCount,lngIDfeedback,lngIDUser,dtcreatedDate,dteditedDate,intFType,intFStatus,intPriority,strFDesc
            
            rstemp.Open query, connTemp, 3, 1
            
            if rstemp.eof then
            %>
                <tr>
                    <td colspan="7">
                        <div class="pcCPmessage"><%response.write dictLanguage.Item(Session("language")&"_viewPostings_v")%></div>
                    </td>
                </tr>
            <%else
                rstemp.MoveFirst
                Count=0
                DO While not rstemp.eof and Count < 5
                lngIDOrder=rstemp("pcComm_IDOrder")
                lngIDfeedback=rstemp("pcComm_idfeedback")
                lngIDUser=rstemp("pcComm_iduser")
                dtcreatedDate=rstemp("pcComm_createdDate")
                dteditedDate=rstemp("pcComm_editedDate")
                intFType=rstemp("pcComm_FType")
                intFStatus=rstemp("pcComm_FStatus")
                intPriority=rstemp("pcComm_Priority")
                strFDesc=rstemp("pcComm_Description")
                
                Dim rstemp1,strFBgColor,intshowbgcolor
            
                query="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
                set retemp1=Server.CreateObject("ADODB.Recordset")
                set rstemp1=connTemp.execute(query)
                FBgColor=""
                if not rstemp1.eof then
                strFBgColor=rstemp1("pcFStat_BgColor")
                intshowbgcolor=1
                end if
                %>    
                <tr class="main" <%if intshowbgcolor="1" then
                if strFBgColor<>"" then%>bgcolor="<%=strFBgColor%>"<%end if
                end if%>>
                <td style="border-bottom: 1px solid #FFF;" nowrap><a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=lngIDfeedback%></a></td>
                <td style="border-bottom: 1px solid #FFF;" nowrap><a href="orddetails.asp?id=<%=lngIDOrder%>"><%=clng(scpre)+clng(lngIDOrder)%></a></td>    
                <td style="border-bottom: 1px solid #FFF;">
                <%
                Dim strPName,strPImg,intPriorityImage
                query="Select * from pcPriority where pcPri_IDPri=" & intPriority
                set rstemp1=connTemp.execute(query)
                if not rstemp1.eof then
                    strPName=rstemp1("pcPri_Name")
                    strPImg=rstemp1("pcPri_Img")
                    intPriorityImage=rstemp1("pcPri_ShowImg")
                    if intPriorityImage="1" then
                        if strPImg<>"" then%>
                            <img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0">
                        <%end if
                    else%>
                        <%=strPName%>
                    <%end if
                end if
                set rstemp1=nothing
                %>
                </td>
                <td style="border-bottom: 1px solid #FFF;"><a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=strFDesc%></a></td>
                <td style="border-bottom: 1px solid #FFF;">
                <%
                Dim intTypeImage
            
                query="Select * from pcFTypes where pcFType_IDType=" & intFType
                set rstemp1=connTemp.execute(query)
                if not rstemp1.eof then
                    strPName=rstemp1("pcFType_Name")
                    strPImg=rstemp1("pcFType_Img")
                    intTypeImage=rstemp1("pcFType_ShowImg")
                    if intTypeImage="1" then
                        if strPImg<>"" then%>
                            <img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0">
                        <%end if
                    else%>
                        <%=strPName%>
                    <%end if
                end if%>
                </td>
                <td style="border-bottom: 1px solid #FFF;" nowrap><%=ShowDateFrmt(dtcreatedDate)%></td>
                <td style="border-bottom: 1px solid #FFF;" nowrap><%=ShowDateFrmt(dteditedDate)%></td>
                <td style="border-bottom: 1px solid #FFF;">
                <%
                if validNum(lngIDUser) and lngIDUser<>0 then
                    query="Select email,name,lastname from Customers where IDCustomer=" & lngIDUser
                    set rstemp1=connTemp.execute(query)
                    if not rstemp1.eof then%>
                        <a href="modcusta.asp?idcustomer=<%=lngIDUser%>" target="_blank"><%=rstemp1("Name") & " " & rstemp1("LastName")%></a>
                    <%else%>
                    Customer has been deleted
                    <%		
                    end if
                    else
                    %>
                        <%response.write dictLanguage.Item(Session("language")&"_viewPostings_2")%>
                    <%end if%>
                </td>
                <td style="border-bottom: 1px solid #FFF;">
            <%
            
                Dim intStatusImage
            
                query="Select * from pcFStatus where pcFStat_IDStatus=" & intFStatus
                set rstemp1=connTemp.execute(query)
                if not rstemp1.eof then
                    strPName=rstemp1("pcFStat_Name")
                    strPImg=rstemp1("pcFStat_Img")
                    intStatusImage=rstemp1("pcFStat_ShowImg")
                    if intStatusImage="1" then
                        if strPImg<>"" then%>
                            <a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><img src="../pc/images/<%=strPImg%>" alt="<%=strPName%>" border="0"></a>
                        <%end if
                    else%>
                        <a href="adminviewfeedback.asp?IDOrder=<%=lngIDOrder%>&IDFeedback=<%=lngIDFeedback%>"><%=strPName%></a>
                    <%end if
                end if
                set retemp1=nothing
                %>
                </td>
                </tr>
                <%
                Count=Count+1
                rstemp.MoveNext
                Loop
                set rstemp=nothing
            end if
            call closeDb()
            %>
            </table>
			</div>

            </td>
        </tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<%
        END IF ' Help Desk turned on/off
        ' END - Recent Help Desk Messages
        %>
	</table>
<%
'****************************
'* Sub-admin login
'****************************
else
%>
	<table class="pcCPcontent">
		<tr> 
			<td>
				<%=dictLanguageCP.Item(Session("language")&"_cpMenu_100")%>
			</td>
		</tr>
	</table>
<% 
end if
%>
<% 'PRV41 start %>
<script>
	$(document).ready(function()
	{	
		function pcf_AutoSendEmails() {
				$.ajax({
					type: "POST",
					url: "prv_AutoSendEmails.asp",
					timeout: 6000,
					global: false,
					success: function(data, textStatus){
						if (data=="0" || data=="")
						{

						} else {

							window.location=data;
	
						}
					}
				});
		}
		pcf_AutoSendEmails();
	});	
</script>
<% 'PRV41 end %>
<!--#include file="AdminFooter.asp"-->