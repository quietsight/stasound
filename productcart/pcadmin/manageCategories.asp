<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Manage Product Categories" 
pcStrPageName="manageCategories.asp"
%>
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="xmlcat.asp"-->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
Dim MaxBufferCats,BufferCats

MaxBufferCats=300
BufferCats=0

Dim rs, conntemp, query
Dim iXML,iRoot,oXML,oRoot
Dim HaveSubCats
HaveSubCats=""
call opendb()

if xmlcat="" or request("action")="refresh" then%>
	<!--#include file="inc_genCatXML.asp"-->
<%end if

dim intCatExist
intCatExist=0


'--- Get Search Form parameters ---
IF request("action")="newsrc" THEN
	form_keyWord=getUserInput(request("keyWord"),0)
	form_exact=getUserInput(request("exact"),0)

	if form_exact <> "" then
		if form_keyWord <> "" then
			SSTR = " AND categoryDesc = " & form_keyWord & " "
		end if
	else
		if form_keyWord <> "" then
			SSTR = " AND categoryDesc like '%" & form_keyWord & "%' "
		end if
	end if
	query="SELECT idCategory,categoryDesc,iBTOhide,idparentcategory FROM categories WHERE idCategory>1 " & SSTR & " ORDER BY categoryDesc"
	Set rs=Server.CreateObject("ADODB.Recordset")
	SET rs=conntemp.execute(query)
	if Not rs.eof then
		intCatExist=1
		catarray=rs.getrows()
		intCatCount=ubound(catarray,2)
	end if
	SET rs=nothing
else
	Set iXML=Server.CreateObject("MSXML2.DOMDocument"&scXML)	
	iXML.async=false
	iXML.loadXML(xmlcat)
	If iXML.parseError.errorCode <> 0 Then
		response.write iXML.parseError.reason & "AAA<br>"
	End if

	Set iRoot=iXML.documentElement
	if iRoot.hasChildNodes then
		intCatExist=1
	end if
end if
%>
<!--#include file="AdminHeader.asp"-->
<script>
var imgopen = new Image();
imgopen.src = "images/btn_collapse.gif";
var imageclose = new Image();
imageclose.src = "images/btn_expand.gif";

function UpDown(tabid)
{
	try
	{
		var etab=document.getElementById('SUB' + tabid);
		if (etab.style.display=='')
		{
			etab.style.display='none';
			var etab=document.images['IMGCAT' + tabid];
			etab.src=imageclose.src;
		}
		else
		{	
			etab.style.display='';
			var etab=document.images['IMGCAT' + tabid];
			etab.src=imgopen.src;
		}
	}
	catch(err)
	{
	}
	
	
}

function ExpandAll()
{
	var tmpStr=HaveSubCats.split("*");
	var i=0;
	for (var i=0;i<tmpStr.length;i++)
	{
	try
	{
		if (tmpStr[i]!="")
		{
			var etab=document.getElementById('SUB' + tmpStr[i]);
			etab.style.display='';
			var etab=document.images['IMGCAT' + tmpStr[i]];
			etab.src=imgopen.src;
		}
	}
	catch(err)
	{
		return(false);
	}
	}
	
}

function CollapseAll()
{
	var tmpStr=HaveSubCats.split("*");
	var i=0;
	for (var i=0;i<tmpStr.length;i++)
	{
	try
	{
		if (tmpStr[i]!="")
		{
			var etab=document.getElementById('SUB' + tmpStr[i]);
			etab.style.display='none';
			var etab=document.images['IMGCAT' + tmpStr[i]];
			etab.src=imageclose.src;
		}
	}
	catch(err)
	{
		return(false);
	}
	}
	
}

</script>

<!--AJAX Functions-->
<script type="text/javascript" src="XHConn.js"></script>

<SCRIPT>
var myConn = new XHConn();

if (!myConn) alert("XMLHTTP not available. Try a newer/better browser.");

var fnWhenDone = function (oXML) {
var xmldoc = oXML.responseXML;
var root_node = xmldoc.getElementsByTagName('count').item(0);
document.getElementById("totalresults").innerHTML="Your search will return " + root_node.firstChild.data + " results.";
document.body.style.cursor='';
};

function runXML()
{
document.getElementById("totalresults").innerHTML="Please wait for processing ... ";
document.body.style.cursor='progress';
myConn.connect("xml_srcCATsCount.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
}

function GetAllValues(theForm){
var ValueStr="";

	var els = theForm.elements; 

	for(i=0; i<els.length; i++){ 

		switch(els[i].type){

			case "select-one" :
			
				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);
				break;

			case "text":

				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);	
				break;

			case "textarea":

				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);	
				break;
				
			case "hidden":
				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);	
				break;

			case "checkbox":

				if (els[i].checked == true)
				{
				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);	
				}
				break;
				

			case "radio":

				if (els[i].checked == true)
				{
				if (ValueStr!="") ValueStr=ValueStr + "&";
				ValueStr=ValueStr + els[i].name + "=" + URLEncode(els[i].value);	
				}
				break;

		}

	}
	return(ValueStr);

}

							// ====================================================================
							// URLEncode Functions
							// Copyright Albion Research Ltd. 2002
							// http://www.albionresearch.com/
							// ====================================================================
							function URLEncode(eStr)
							{
							// The Javascript escape and unescape functions do not correspond
							// with what browsers actually do...
							var SAFECHARS = "0123456789" +					// Numeric
											"ABCDEFGHIJKLMNOPQRSTUVWXYZ" +	// Alphabetic
											"abcdefghijklmnopqrstuvwxyz" +
											"-_.!~*'()";					// RFC2396 Mark characters
							var HEX = "0123456789ABCDEF";
						
							var plaintext = eStr;
							var encoded = "";
							for (var i = 0; i < plaintext.length; i++ ) {
								var ch = plaintext.charAt(i);
									if (ch == " ") {
										encoded += "+";				// x-www-urlencoded, rather than %20
								} else if (SAFECHARS.indexOf(ch) != -1) {
										encoded += ch;
								} else {
										var charCode = ch.charCodeAt(0);
									if (charCode > 255) {
											alert( "Unicode Character '" 
																		+ ch 
																		+ "' cannot be encoded using standard URL encoding.\n" +
															"(URL encoding only supports 8-bit characters.)\n" +
													"A space (+) will be substituted." );
										encoded += "+";
									} else {
										encoded += "%";
										encoded += HEX.charAt((charCode >> 4) & 0xF);
										encoded += HEX.charAt(charCode & 0xF);
									}
								}
							} // for
						
								return encoded;
							};
	

				function scrollWindow()
					{
					window.scrollBy(100,4000);
					}
</SCRIPT>

<!--End of AJAX Functions-->

<%Sub ShowCatInfo(parentNode)
Dim catNodes,catNode,NodeAtts,intAtt,subNodes,subNode

				Set catNodes = parentNode.childNodes

				For Each catNode In catNodes
					BufferCats=BufferCats+1
					If BufferCats>=MaxBufferCats then
						Response.Flush()
						Response.Clear()
						BufferCats=0
					End if
					Set NodeAtts = catNode.attributes 
					For Each intAtt in NodeAtts
   						if intAtt.name="catName" then
							tCategoryDesc=intAtt.value
						end if
						if intAtt.name="catID" then
							tIdCategory=intAtt.value
						end if
						if intAtt.name="catParentID" then
							parent=intAtt.value
						end if
						if intAtt.name="RetailHide" then
							tmpCatRetailHide=intAtt.value
						end if
						if intAtt.name="iBTOHide" then
							tIBTOhide=intAtt.value
						end if
					Next
					%>
					<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
                    	<td width="1%" cellpadding="10">
						<%if catNode.hasChildNodes then%><img align="absmiddle" name="IMGCAT<%=tIdCategory%>" border="0" src="images/btn_expand.gif" onclick="javascript:UpDown(<%=tIdCategory%>);" style="margin-right: 3px;"><%end if%></td>
						<td width="74%" cellpadding="10"><a href="modCata.asp?idcategory=<%=tIdCategory%>&top=<%=top%>&parent=<%=parent%>"><%=tCategoryDesc%></a></td>
						<td width="25%" align="right" nowrap="nowrap" class="cpLinksList"> 
                            <%
							Dim pcIntHidden
							pcIntHidden=0
							if (tIBTOhide<>"") and (tIBTOhide="1") then pcIntHidden=1
							if pcIntHidden=1 then
							%>
                            	<img src="images/hidden.gif" width="33" height="9" border=0 align="baseline">&nbsp;&nbsp;
							<%
							else
							%>
                            	<a href="../pc/viewcategories.asp?idcategory=<%=tIdCategory%>" target="_blank" title="View this category in the storefront"><img src="images/pcIconPreview.jpg" width="12" height="12" alt="View"></a>&nbsp;&nbsp;
							<%
							end if
							
							if catNode.hasChildNodes then
							%>
                            <a href="viewCat.asp?parent=<%=tIdCategory%>&hidden=<%=pcIntHidden%>" title="Order subcategories">Order</a> :  
							<%
							end if
							%>
                            <a href="editCategories.asp?nav=<%=request("nav")%>&lid=<%= tIdCategory %>" title="List products assigned to this category">Products</a> : <a href="modCata.asp?idcategory=<%=tIdCategory%>&top=<%=top%>&parent=<%=parent%>" title="Edit this category">Edit</a> : <a href="AddDupCat.asp?idcategory=<%=tIdCategory%>&top=<%=top%>&parent=<%=parent%>" title="Clone this category">Clone</a> : <a href="JavaScript: if (confirm('Are you sure you want to delete this category? If the category contains any products, you will first be prompted to remove them from it. If no products have been assigned, the category will be immediately deleted.')) location='modcatb.asp?idcategory=<%=tIdCategory%>';" title="Delete this category">Del</a>
						</td>
					</tr>
                    <%if catNode.hasChildNodes then
					HaveSubCats=HaveSubCats & "*" & tIdCategory%>
                    <tr>
                    	<td width="1%"></td>
                        <td colspan="2" width="99%" cellpadding="0">
                        	<table class="pcCPcontentCat" id="SUB<%=tIdCategory%>" style="display:none">
							<tr>
                            	<td>
                                	<%call ShowCatInfo(catNode)%>
                                </td>
							</tr>
                            </table>
                        </td>
					</tr>
					<%end if
				Next
End Sub%>

		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
		<% CategoryName="Top Level Category (Root)"	%>           
		<table class="pcCPcontentCat">
            <tr>
            	<td colspan="3" class="cpOtherLinks">
                    <a href="JavaScript:;" onClick="javascript:document.getElementById('FindProducts').style.display=''; scrollWindow()">Find</a> | 
                    <a href="instCata.asp">Add New</a> | 
					<a href="srcFreePrds.asp">Orphaned Products</a> | 
                    <a href="editCategories.asp?nav=&lid=1">Root-Level Products</a> | 
					<a href="../pc/viewCategories.asp" target="_blank">Browse</a> | 
                    <a href="viewCat.asp?parent=1&hidden=0">Order</a> | 
                    <a href="viewCatRecent.asp">Recently Edited</a> | 
					<a href="genCatNavigation.asp">Create/Update Storefront Navigation</a>
                </td>
            </tr>
            <tr> 
				<td colspan="3" class="pcCPspacer"></td>
			</tr>
            <tr> 
				<td colspan="3">
                <div style="float: right"><a href="javascript:ExpandAll();">Expand All</a> | <a href="javascript:CollapseAll();">Collapse All</a></div>
                If the list below does not seem current, please <a href="manageCategories.asp?action=refresh">click here</a> (updates XML cache).
                </td> 
			</tr>
				<tr>
					<td colspan="3" class="pcCPspacer"></td>
				</tr>
			<%Response.Flush()
			Response.Clear()
			If intCatExist=0 Then %>
				<tr> 
					<td colspan="3">No Categories Found - <a href="instCata.asp">Add new</a></td>
				</tr>
			<% Else 
				call ShowCatInfo(iRoot)
			End If
			call closeDb()
			%>
		<tr> 
			<td colspan="3" class="pcCPspacer">
            </td>
		</tr>
		<tr> 
			<td bgcolor="#F5F5F5" align="right" colspan="3" class="pcCPspacer">
            <a href="javascript:ExpandAll();">Expand All</a> | <a href="javascript:CollapseAll();">Collapse All</a>
            <script>
				HaveSubCats="<%=HaveSubCats%>";
			</script>
            </td>
		</tr>

		<tr> 
			<td colspan="3">
				<table id="FindProducts" class="pcCPcontent" style="display:none;">
					<tr>
						<td>
						<%
							src_FormTitle1="Find A Category"
							src_FormTitle2=""
							src_FormTips1="Use the following filters to look for categories in your store."
							src_FormTips2=""
							src_DisplayType=0
							src_ShowLinks=1
							src_ParentOnly=0
							src_FromPage="manageCategories.asp?"
							src_ToPage="manageCategories.asp?action=newsrc"
							src_Button1=" Search "
							src_Button2=""
							src_Button3=" Back "
							src_PageSize=15
							UseSpecial=1
							session("srcCat_from")=""
						%>
						<!--#include file="inc_srcCATs.asp"-->
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<!--#include file="AdminFooter.asp"-->