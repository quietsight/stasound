<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Specials" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<%
Dim query, rs, rsTemp, connTemp, pcDontshow, pIdProduct, pSku, pDescription, pcMessage, pAction, sMode, strORD, iPageCount, strCol, Count

'****************************
'* Action Items
'****************************

	' UPDATE list of current specials
		sMode=Request("Submit")
		If sMode <> "" Then
			prdlist=split(request("prdlist"),",")
			call openDb()
			For i=lbound(prdlist) to ubound(prdlist)
			id=prdlist(i)
			If (id<>"0") and (id<>"") then
				
				query="UPDATE products SET hotDeal=-1 WHERE idproduct="&id
				Set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			Next
			call closedb()
			response.redirect "manageSpecials.asp?msg="&msg
		End If

	' REMOVE selected product from list of specials
		pAction=request.QueryString("action")
			if pAction = "remove" then
				pIdproduct=request.Querystring("idproduct")
					if trim(pIdproduct)="" then
						 response.redirect "msg.asp?message=2"
					end if
				call openDb()
				query="UPDATE products SET hotDeal=0 WHERE idproduct=" &pIdproduct
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
		
					if err.number <> 0 then
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in delSpc: "&Err.Description) 
					end If
				
				set rs=nothing
				call closeDb()
				response.Redirect("manageSpecials.asp?s=1&msg=The product was successfully removed from your list of Specials")
			end if
		
'****************************
'* END Action Items
'****************************
			
	' Paging and sorting

	if request("iPageCurrent")="" then
    iPageCurrent=1 
	else
		iPageCurrent=Request("iPageCurrent")
	end If
	
	strORD=request("order")
	if strORD="" then
		strORD="description"
	End If
	
	strSort=request("sort")
	if strSort="" Then
		strSort="ASC"
	End If 
	
	call openDb()

	' Retrieve specials from database
	query="SELECT idproduct, sku, description FROM products WHERE active=-1 AND hotdeal=-1 ORDER BY "& strORD &" "& strSort
	Set rsS=Server.CreateObject("ADODB.Recordset")   
	rsS.CacheSize=100
	rsS.PageSize=100	
	rsS.Open query, connTemp, adOpenStatic, adLockReadOnly

		if err.number <> 0 then
			SET rs=nothing
			call closeDb()
			response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving specials from database on manageSpecials: "&Err.Description) 
		end If
	
	' Set a variable if no specials are found
	if rsS.EOF then
		pcDontshow = 1
	end if
	
	' Find out if all products have been set as Specials
		query="SELECT idProduct FROM products WHERE active=-1 AND configOnly=0 AND removed=0 AND hotdeal=0"
		Set rsTemp=Server.CreateObject("ADODB.Recordset")
		rsTemp.Open query, connTemp, adOpenStatic, adLockReadOnly
		if rsTemp.EOF then
			pcAllSpecials = 1
		end if
		set rsTemp = nothing
%>

	<table class="pcCPcontent">
		<tr>
			<td colspan="5">
			<p><strong>Specials</strong> are shown on the &quot;<a href="../pc/showspecials.asp" target="_blank">View Specials</a>&quot; in your storefront and can also be shown on the &quot;home page&quot; (<a href="../pc/home.asp">View</a> - <a href="manageHomePage.asp">Manage</a>). You can make a product a &quot;special product&quot; by using this page, when adding or modifying a product, or when changing the properties of multiple products at once using the <a href="globalchanges.asp?nav=0">Global Changes</a> or the Import features.</p>
				<table id="FindProducts" class="pcCPcontent" style="display:none;">
					<tr>
						<td>
						<%
							src_FormTitle1="Find Products"
							src_FormTitle2="Add New Specials"
							src_FormTips1="Use the following filters to look for products in your store."
							src_FormTips2="Select the products that you would like to add to your list of store specials."
							src_IncNormal=1
							src_IncBTO=1
							src_IncItem=0
							src_Special=2
							src_DisplayType=1
							src_ShowLinks=0
							src_FromPage="manageSpecials.asp"
							src_ToPage="manageSpecials.asp?submit=yes"
							src_Button1=" Search "
							src_Button2=" Add as Special "
							src_Button3=" Back "
							src_PageSize=15
						%>
							<!--#include file="inc_srcPrds.asp"-->
						</td>
					</tr>
				</table>
			</td>
		</tr>

        <tr>
            <td colspan="5" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>

		<tr>
			<th width="5%" nowrap>
				<a href="manageSpecials.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" alt="Sort Ascending"></a><a href="manageSpecials.asp?iPageCurrent=<%=iPageCurrent%>&order=sku&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" alt="Sort Descending"></a>
			</th>
			<th width="15%" nowrap>SKU</th>
			<th width="80%" nowrap colspan="3"> 
				<a href="manageSpecials.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=ASC"><img src="images/sortasc_blue.gif" width="14" height="14" alt="Sort Ascending"></a><a href="manageSpecials.asp?iPageCurrent=<%=iPageCurrent%>&order=description&sort=Desc"><img src="images/sortdesc_blue.gif" width="14" height="14" alt="Sort Descending"></a>&nbsp;Product&nbsp;&nbsp;<% if pcAllSpecials <> 1 then %>
			  <div class="pcSmallText" style="float:right; padding-top: 3px;"><a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''">Add New</a>
			  <% end if %>
		    </div></th>
		</tr>
       	<tr>
        	<td class="pcCPspacer" colspan="5"></td>
        </tr>
			<%
			If pcDontshow=1 then
			' No specials found
			%>
			<tr>
				<td colspan="5">
					<div class="pcCPmessage">No products have been setup as &quot;Specials&quot; on this store.</div>
				</td>
			</tr>
                      
		<% Else 
					' get the max number of pages
					iPageCount=rsS.PageCount
					If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
					If iPageCurrent < 1 Then iPageCurrent=1
												
					' set the absolute page
					rsS.AbsolutePage=iPageCurrent
					hCnt=0
					
					pcArr=rsS.getRows()
					set rsS=nothing
					
					intCount=ubound(pcArr,2)
					
					For i=0 to intCount
						' Assign values							
						pIdProduct=pcArr(0,i)
						pSku=pcArr(1,i)
						pDescription=pcArr(2,i)
						%>
                      
						<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
							<td colspan="2" bgcolor="<%=strCol%>"><a href="FindProductType.asp?id=<%=pIdProduct%>" target="_blank"><%=pSku%></a></td>
							<td colspan="2" bgcolor="<%=strCol%>"><a href="FindProductType.asp?id=<%=pIdProduct%>" target="_blank"><%=pDescription%></a></td>
							<td bgcolor="<%= strCol %>" align="center">
								<a href="javascript:if (confirm('You are about to remove this item from your list of store Specials. Are you sure you want to complete this action?')) location='manageSpecials.asp?action=remove&idproduct=<%=pIdProduct%>'"><img src="images/pcIconDelete.jpg"></a>
							</td>
						</tr>
					<% 
					Next
			End If
			set rsS=nothing
			set rsTemp=nothing
			call closedb()%>
	</table>
<!--#include file="AdminFooter.asp"-->