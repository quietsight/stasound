<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="../includes/languages.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<!--#include file="pcSeoFunctions.asp"-->
<%
'Check to see if store has been turned off by admin, then display message
If scStoreOff="1" then
	response.redirect "msg.asp?message=59"
End If

' Load affiliate ID
affVar=session("pc_idaffiliate")
if not validNum(affVar) then
	response.redirect "AffiliateLogin.asp"
end if
%>
<!--#include file="header.asp"-->
<%sMode=request("action")
	if sMode <> "" then
		sMode="1"
		idproduct=request.Form("product")
		idaffiliate=session("pc_IDAffiliate")
	end If %><%
Dim rstemp, connTemp, strSQL

call opendb()
%>
<script>
<!-- Hide me
var copytoclip=1

function HighlightAll(theField) {
var tempval=eval("document."+theField)
tempval.focus()
tempval.select()
if (document.all&&copytoclip==1){
therange=tempval.createTextRange()
therange.execCommand("Copy")
window.status="Contents highlighted and copied to clipboard!"
setTimeout("window.status=''",1800)
}
}
// -->
</script>
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1><%=dictLanguage.Item(Session("language")&"_AffgenLinks_1")%></h1>
			</td>
		</tr>
		<tr>
			<td>
				<form method="post" name="links" action="Affgenlinks.asp?action=1" class="pcForms">
					<table class="pcContent">
                        <tr><td class="pcSpacer" colspan="2"></td></tr> 
                        <tr><th colspan="2"><%=dictLanguage.Item(Session("language")&"_AffgenLinks_8")%></th></tr> 
                        <tr><td class="pcSpacer" colspan="2"></td></tr>
                        <tr>
							<td width="15%" valign="top" nowrap>
								<p><%=dictLanguage.Item(Session("language")&"_AffgenLinks_2")%></p>
							</td>
							<td width="85%">
								<select name="product">  
								<%
								query="SELECT idproduct,description FROM products WHERE active=-1 AND configOnly=0 AND removed=0 ORDER BY description ASC"
								set rsPrd=Server.CreateObject("adodb.recordset")
								set rsPrd=conntemp.execute(query)
								if err.number <> 0 then
									set rsPrd=nothing
									response.redirect "techErr.asp?error="& Server.Urlencode("Error in retreiving products from database: "&Err.Description) 
								end If
								
									do until rsPrd.eof
									intTempIdProduct=rsPrd("idproduct")
									strTempDescription=rsPrd("description")
									if sMode="1" And Cint(idproduct)= Cint(intTempIdProduct) then 
										pDescription=strTempDescription %>
										<option value="<%response.write intTempIdProduct%>" selected> 
									<% else %>
										<option value="<%response.write intTempIdProduct%>"> 
									<% end if %>
									<%=strTempDescription%>
									</option>
									<%
									rsPrd.movenext
									loop
									set rsPrd=nothing
									call closedb()
								%>
								</select>
							</td>
						</tr>
						<tr class="pcSpacer">
							<td></td>
						</tr>
						<tr>
							<td colspan="2"><input type="submit" name="submit1" value="<%=dictLanguage.Item(Session("language")&"_AffgenLinks_3")%>" id="submit" class="submit2"></td>	
						</tr>
						
						<% If sMode="1" then
						
							call openDb()
							query="SELECT idproduct, description FROM products WHERE idproduct="&idproduct
							set rsPrd=Server.CreateObject("adodb.recordset")
							set rsPrd=conntemp.execute(query)
							pProductDesc=rsPrd("description")
							set rsPrd=nothing
						
							'// SEO Links
							'// Build Navigation Product Link
							'// Get the first category that the product has been assigned to, filtering out hidden categories
							query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& idproduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=conntemp.execute(query)
							if not rs.EOF then
								pIdCategory=rs("idCategory")
							else
								pIdCategory=1
							end if
							set rs=nothing
							call closeDb()

							if scSeoURLs=1 then
								pcStrPrdLink=pProductDesc & "-" & pIdCategory & "p" & idproduct & ".htm"
								pcStrPrdLink=removeChars(pcStrPrdLink)
								pcStrPrdLink=pcStrPrdLink & "?"
							else
								pcStrPrdLink="viewPrd.asp?idproduct=" & idproduct &"&"
							end if
							'//
							
							tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrPrdLink&"idaffiliate="&idaffiliate),"//","/")
							tempURL=replace(tempURL,"http:/","http://")
						
						%>
                            <tr class="pcSpacer">
                                <td></td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <p><%=dictLanguage.Item(Session("language")&"_AffgenLinks_4")%><%=pProductDesc%><%=dictLanguage.Item(Session("language")&"_AffgenLinks_5")%><%=pAffiliateName%>:</p>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <a class="highlighttext" href="javascript:HighlightAll('links.link1')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
                                    <input type="text" name="link1" size="80" value="<%=tempURL%>">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2"><%=dictLanguage.Item(Session("language")&"_AffgenLinks_6")%><%=pAffiliateName%>:</td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <a class="highlighttext" href="javascript:HighlightAll('links.link2')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
                                    <%
                                    tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/home.asp?idaffiliate="&idaffiliate),"//","/")
                                    tempURL=replace(tempURL,"http:/","http://")
                                    %>
                                    <input type="text" name="link2" size="80" value="<%=tempURL%>">
                                </td>
                            </tr>
						<% end if %> 
                        
						<% If SNW_AFFILIATE="1" then %>
                            <tr><td class="pcSpacer" colspan="2"></td></tr> 
                            <tr><th colspan="2"><%=dictLanguage.Item(Session("language")&"_AffgenLinks_9")%></th></tr> 
                            <tr><td class="pcSpacer" colspan="2"></td></tr>
                            <tr>
                                <td colspan="2"><%=dictLanguage.Item(Session("language")&"_AffgenLinks_7")%>:</td>
                            </tr>
                            <tr>
                                <td colspan="2" align="left" valign="top">
                                    <%
                                    tempURL=replace((scStoreURL&"/"&scPcFolder),"//","/")
                                    tempURL=replace(tempURL,"http:/","http://")
                                    tempCode="<script language=""javascript"">idaffiliate="""& session("pc_IDAffiliate") &""";</script><script type=""text/javascript"" src="""&tempURL&"/pc/pcSyndication.js""></script>"									
									%>
                                    <a class="highlighttext" href="javascript:HighlightAll('links.link3')"><img src="images/edit2.gif" width="25" height="23" border="0"></a>
                                    <textarea name="link3" cols="50" rows="10"><%=tempCode%></textarea>						  
                              </td>

                            </tr>
						<% end if %> 
					</table>
				</form>
			</td>
		</tr>
		<tr>
			<td><a href="javascript:history.go(-1)"><img src="<%=rslayout("back")%>" border=0></a></td>
		</tr>
	</table>
</div>
<!--#include file="Footer.asp"-->