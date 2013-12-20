<%src_ForCats="1"%>
<% call openDb() %>
<!--#include file="inc_srcPrdQuery.asp"-->
<%
Set rsCATcount=Server.CreateObject("ADODB.Recordset")
rsCATcount.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
Dim pcIntCategoryCount
pcIntCategoryCount = 0
if not rsCATcount.eof then
	pcIntCategoryCount = rsCATcount.recordcount
else
	pcIntCategoryCount = 0
end if
if pcIntCategoryCount=1 then
	pcCatScrollStyle="display:none;"
	else
	pcCatScrollStyle="show"	
end if
set rsCATcount=nothing

Set rsCAT=Server.CreateObject("ADODB.Recordset")
Set rsCAT=connTemp.execute(query)

IF NOT rsCAT.eof THEN
	form_idcategory=getUserInput(request("idcategory"),0)
	form_customfield=getUserInput(request("customfield"),0)
	form_SearchValues=getUserInput(request("SearchValues"),0)
	form_priceFrom=getUserInput(request("priceFrom"),0)
	form_priceUntil=getUserInput(request("priceUntil"),0)
	form_withstock=getUserInput(request("withstock"),0)
	form_sku=getUserInput(request("sku"),0)
	form_IDBrand=getUserInput(request("IDBrand"),0)
	form_keyWord=getUserInput(request("keyWord"),0)
	form_exact=getUserInput(request("exact"),0)
	form_resultCnt=getUserInput(request("resultCnt"),0)
	form_order=getUserInput(request("order"),0)
	form_pageStyle=getUserInput(request("pageStyle"),0)
	form_incSale=getUserInput(request("incSale"),2)
	form_IDSale=getUserInput(request("IDSale"),6)%>
	
	<tr<% if pcCatScrollStyle<>"show" then%> style="<%=pcCatScrollStyle%>"<%end if%>>
		<td colspan="2" class="pcSectionTitle">
			<%response.write dictLanguage.Item(Session("language")&"_showSearchResults_2")%>
		</td></tr>
	<tr<% if pcCatScrollStyle<>"show" then%> style="<%=pcCatScrollStyle%>"<%end if%>>
		<td colspan="2">
	
		<form name="ajaxSearch">
			<input type="hidden" name="idcategory" value="<%=form_idcategory%>">
			<input type="hidden" name="customfield" value="<%=form_customfield%>">
			<input type="hidden" name="SearchValues" value="<%=form_SearchValues%>">
			<input type="hidden" name="priceFrom" value="<%=form_priceFrom%>">
			<input type="hidden" name="priceUntil" value="<%=form_priceUntil%>">
			<input type="hidden" name="withstock" value="<%=form_withstock%>">
			<input type="hidden" name="sku" value="<%=form_sku%>">
			<input type="hidden" name="IDBrand" value="<%=form_IDBrand%>">
			<input type="hidden" name="keyWord" value="<%=form_keyWord%>">
			<input type="hidden" name="exact" value="<%=form_exact%>">
			<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
			<input type="hidden" name="order" value="<%=form_order%>">
			<input type="hidden" name="iPageCurrent" value="1">
			<input type="hidden" name="pageStyle" value="<%=form_pageStyle%>">
			<input type="hidden" name="incSale" value="<%=form_incSale%>">
			<input type="hidden" name="IDSale" value="<%=form_IDSale%>">
		</form>
			
			<div id="pcCatScroll"<% if pcCatScrollStyle<>"show" then%> style="<%=pcCatScrollStyle%>"<%end if%>>
			
				<div id="pcCatScrollArrows">
					<div id="pcpcCatScrollUp"><img style="cursor: pointer; cursor: hand;" src="<%=rsIconObj("arrowUp")%>" id="UpArrow" alt="Scroll Up" onmouseover="javascript:moveup()" onmouseout="javascript:stopscroll()"></div>
					<div id="pcpcCatScrollDown"><img style="cursor: pointer; cursor: hand;" src="<%=rsIconObj("arrowDown")%>" id="DownArrow" alt="Scroll Down" onmouseover="javascript:movedown()" onmouseout="javascript:stopscroll()"></div>
				</div>
	
				<div id="pcCatScrollItems">
					<SCRIPT language="JavaScript1.2">
					
					//Manual Scroller- © Dynamic Drive 2001
					//For full source code, visit http://www.dynamicdrive.com
					
					//specify speed of scroll (greater=faster)
					var speed=2
					
					iens6=document.all||document.getElementById
					ns4=document.layers
					
					if (iens6){
					document.write('<div onmouseover="javascript:getcontent_height();" id="catcontainer" style="z-index : 90; position:relative;width:100%;height:100px; overflow:hidden">')
					document.write('<div onmouseover="javascript:getcontent_height();" id="catcontent" style="position:absolute;width:100%;left:0px;top:0px;">')
					}
					</script>
					
					<ilayer onmouseover="javascript:getcontent_height();" name="nscontainer" width=100% clip="0,0,155,100">
					<layer onmouseover="javascript:getcontent_height();" name="nscontent" width=100% visibility=hidden>
					<div class="pcCatSearchResults">
						<ul>
						<%
						tmp_CatID=0
						tmpCatName=""
						tmpCount=0
						
						Do while not rsCAT.eof				
							tmp_CatCount=rsCAT("ProductCount")
							tmp_CatID=rsCAT("idcategory")
							tmp_CatName=rsCAT("categoryDesc")						
							if tmp_CatCount>0 then						
								if tmp_CatCount > 1 then
									tmp_CatCountMessage = dictLanguage.Item(Session("language")&"_ShowSearch_6")
								else
									tmp_CatCountMessage = dictLanguage.Item(Session("language")&"_ShowSearch_7")
								end if
								%> 
								<li><a href="javascript:document.ajaxSearch.idcategory.value='<%=tmp_CatID%>';document.ajaxSearch.submit();" <%if scStoreUseToolTip="1" or scStoreUseToolTip="3" then%>onmouseover="javascript:document.ajaxSearch.idcategory.value='<%=tmp_CatID%>'; sav_callxml='1'; runXML('cat_<%=tmp_CatID%>');" onmouseout="javascript: sav_callxml=''; hidetip();"<%end if%>><%=tmp_CatName%></a> (<%=tmp_CatCount & tmp_CatCountMessage %>)</li> 
								<%
							end if				
			
							rsCat.MoveNext
						loop
						%>
						</ul>
					</div>
					</layer>
					</ilayer>
					<script language="JavaScript1.2">
					if (iens6){
					document.write('</div></div>')
					var crossobj=document.getElementById? document.getElementById("catcontent") : document.all.catcontent
					var cross1obj=document.getElementById? document.getElementById("catcontainer") : document.all.catcontainer
					var contentheight=crossobj.offsetHeight
					}
					else if (ns4){
					var crossobj=document.nscontainer.document.nscontent
					var contentheight=crossobj.clip.height
					}
					</script>
				</div>
				
			</div>
	
			<script language="JavaScript1.2">
			function movedown(){
			getcontent_height();
			if (window.moveupvar) clearTimeout(moveupvar)
			if (iens6&&parseInt(crossobj.style.top)>=(contentheight*(-1)+100))
			crossobj.style.top=parseInt(crossobj.style.top)-speed+"px"
			else if (ns4&&crossobj.top>=(contentheight*(-1)+100))
			crossobj.top-=speed
			movedownvar=setTimeout("movedown()",20)
			}
			
			function moveup(){
			getcontent_height();
			if (window.movedownvar) clearTimeout(movedownvar)
			if (iens6&&parseInt(crossobj.style.top)<=0)
			crossobj.style.top=parseInt(crossobj.style.top)+speed+"px"
			else if (ns4&&crossobj.top<=0)
			crossobj.top+=speed
			moveupvar=setTimeout("moveup()",20)
			}
			
			function stopscroll(){
			if (window.moveupvar) clearTimeout(moveupvar)
			if (window.movedownvar) clearTimeout(movedownvar)
			}
			
			function movetop(){
			stopscroll()
			if (iens6)
			crossobj.style.top=0+"px"
			else if (ns4)
			crossobj.top=0
			}
			
			function getcontent_height(){
			if (iens6)
			contentheight=crossobj.offsetHeight
			else if (ns4)
			document.nscontainer.document.nscontent.visibility="show"
			var ie=document.all
			if (ie)
			{	
				if (contentheight<=100) {
					document.getElementById("DownArrow").style.visibility="hidden"
					document.getElementById("UpArrow").style.visibility="hidden"
				}
				else
				{
					document.getElementById("DownArrow").style.visibility="visible"
					document.getElementById("UpArrow").style.visibility="visible"
				}
			}
			}
			
			</script>
			</td>
		</tr>
<%
END IF
set rsCAT=nothing
call closeDb()
%>