<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Cross Selling - Edit an existing relationship" %>
<% Section="products" %>
<%PmAdmin="2*3*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp" -->
<!--#include file="AdminHeader.asp"-->
<% dim mySQL, conntemp, rstemp, cnt
cnt=0
idmain=request.QueryString("idmain")
if request.Form("submitUpdate")<>"" then
	idmain=request.Form("idmain") 
	icnt=request.Form("icnt")
	call openDb()
	set rs=Server.CreateObject("ADODB.RecordSet")
	For i=0 to (cint(icnt)-1)
		idcrosssell=request.Form("idcrosssell"&i)
		order=request.Form("num"&i)
		cs_type=request.Form("type"&i)
		if (cs_type="Bundle") then
		    discount=replacecomma(request.Form("discount"&i))
		    if request("percent"&i)="on" then
			    ipercent="-1"
		    else
			    ipercent="0"
		    end if
		    irequired="0"
		else
		    if request("required"&i)="on" then
			    irequired="-1"
		    else
			    irequired="0"
		    end if
	        ipercent="0"
	        discount="0"
		end if
		query="UPDATE cs_relationships SET num="&order&",discount="&discount&",isPercent="&ipercent&",cs_type='"&cs_type&"',isRequired="&irequired&" "
		query=query&"WHERE idcrosssell="&idcrosssell&";"
		set rs=conntemp.execute(query)
		set rs=nothing
	Next
	call closeDb()
	
	response.Redirect "crossSellEdit.asp?idmain="&idmain
	response.end
end if %>
	<script language="javascript">
			function SetOptions(itemcnt)
			{
					var obj = eval("document.form1.elements.type"+itemcnt);
					var dobj = eval("document.form1.elements.discount"+itemcnt);
					var pobj = eval("document.form1.elements.percent"+itemcnt);
					var robj = eval("document.form1.elements.required"+itemcnt);
					if (obj.value=="Bundle")
					{
							dobj.disabled=false;
							pobj.disabled=false;
							robj.disabled=true;
					}
					else
					{
							dobj.disabled=true;
							pobj.disabled=true;
							robj.disabled=false;
							
					}
			}
			
			function CheckDiscount(itemcnt)
			{
					var i = 0;
					while (i < itemcnt)
					{
							var obj = eval("document.form1.elements.type"+i);
							var dobj = eval("document.form1.elements.discount"+i);
							if (obj.value=="Bundle")
							{
									if(dobj.value=="0")
									{
											alert("Discount is required for Bundles.");
											return false;
									}
							}
							i+=1;
					}
					return true;
			}
	</script>
<% 
	call openDb()
	query="SELECT idproduct, description, sku FROM products WHERE idproduct="&idmain&";"
	set rs=Server.CreateObject("ADODB.Recordset") 
	set rs=conntemp.execute(query)
	productName=rs("description")
	productSku=rs("sku")
	set rs = nothing
%>
<form name="form1" method="post" action="" class="pcForms">
	<input name="idmain" type="hidden" value="<%=idmain%>">
	<table class="pcCPcontent">
	<tr>       
		<td>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
        <h2>Primary product: <strong><%=productName%></strong> (<%=productSku%>) - <a href="FindProductType.asp?id=<%=idmain%>" target="_blank">Edit</a>&nbsp;|&nbsp;<a href="../pc/viewPrd.asp?idproduct=<%=idmain%>&adminPreview=1" target="_blank">View</a></h2></td>
	</tr>    
	<tr> 
		<td> 
			<table class="pcCPcontent">
			<tr> 
				<th width="50%">Related products:</th>
				<th width="10%" align="center" nowrap="nowrap">Type&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=205')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
				<th width="10%" align="right" nowrap="nowrap">Disc.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=206')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
				<th width="5%" align="right">%</th>
				<th width="10%" align="right" nowrap="nowrap">Required&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=207')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
				<th width="10%" align="center">Order</th>
				<th width="5%">&nbsp;</th>
			</tr>
			<tr>          
				<td colspan="7" class="pcCPspacer"></td>
			</tr> 
			<%
				query="SELECT cs_relationships.idcrosssell, cs_relationships.idrelation,cs_relationships.cs_type, cs_relationships.num,cs_relationships.discount,cs_relationships.isPercent,cs_relationships.isRequired, products.idproduct, products.description, products.sku, products.active, cs_relationships.idproduct FROM products INNER JOIN cs_relationships ON products.idProduct=cs_relationships.idrelation WHERE (((cs_relationships.idproduct)="&idmain&"));"
				set rs=Server.CreateObject("ADODB.Recordset") 
				set rs=conntemp.execute(query)
				cnt=0
				do until rs.eof				
			%>
			<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
				<td width="50%"><a href="FindProductType.asp?id=<%=rs("idrelation")%>" target="_blank"><%=rs("description")%></a> (<%=rs("sku")%>)<% if rs("active")=0 then%>&nbsp;<span class="pcCPnotes">Inactive</span><%end if%></td>
				<td width="10%" align="right"> 
				<select name="type<%=cnt%>" onchange="SetOptions('<%=cnt%>')">
						<option value="Bundle" <%if rs("cs_type")="Bundle" then %> selected <%end if %>>Sold as Bundle</option>
						<option value="Accessory" <%if rs("cs_type")="Accessory" then %> selected <%end if %>>Not a Bundle</option>    
				</select>
				</td>
				<td width="10%" align="right"> 
				<input name="discount<%=cnt%>" type="text" value="<%=money(rs("discount"))%>" size="3">
				</td>
				<td width="5%" align="right"> 
				<input name="percent<%=cnt%>" type="checkbox" <%if rs("isPercent")<>0 then %>checked<%end if%> class="clearBorder">
				</td>
				<td width="10%" align="center"> 
				<input name="required<%=cnt%>" type="checkbox" <%if rs("isRequired")<>0 then %>checked<%end if%> class="clearBorder">
				</td>
				<td width="10%" align="right"> 
				<input name="num<%=cnt%>" type="text" value="<%=rs("num")%>" size="3">
				<input name="idcrosssell<%=cnt%>" type="hidden" value="<%=rs("idcrosssell")%>">
				</td>
				<td width="5%" align="center"><a href="javascript:if (confirm('You are about to completely remove this relationship. Click OK to confirm the removal.')) location='delRelationship.asp?type=2&idmain=<%=idmain%>&idcrosssell=<%=rs("idcrosssell")%>'"><img src="images/pcIconDelete.jpg" width="12" height="12" alt="Remove" title="Remove from cross selling relationship"></a></td>
			</tr>
			<script language="javascript">SetOptions('<%=cnt%>')</script>								
			<%
					cnt=cnt+1
					rs.moveNext
					loop
					set rs=nothing
					call closeDb()
			%>
											
			<input name="icnt" type="hidden" value="<%=cnt%>">
			</table>
		</td>
	</tr>
	<tr>          
		<td class="pcCPspacer"></td>
	</tr> 
	<tr> 
		<td align="center">
		<input type="submit" name="submitUpdate" value="Update Relationship" onClick="javascript: return CheckDiscount('<%=cnt%>');" class="submit2">&nbsp;
		<input type="button" name="" value="Add more products" onClick="document.location.href='crossSellAddb.asp?idmain=<%=idmain%>'">&nbsp;
        <input type="button" name="" value="Copy to other products" onClick="document.location.href='crossSellMultiAdd.asp?idmain=<%=idmain%>'">&nbsp;
		<input type="button" name="" value="View other relationships" onClick="document.location.href='crossSellView.asp'">
		</td>
	</tr>
	<tr> 
		<td align="center">
			<input type="button" value="Edit Product-specific cross selling settings" onClick="document.location.href='crossSellSettings.asp?idmain=<%=idmain%>'">&nbsp;
		<input type="button" value="Edit Store-wide cross selling settings" onClick="document.location.href='crossSellSettings.asp?idmain=1'"> 
		</td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->