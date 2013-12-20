<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if session("cp_cmt_action")=1 then
pageTitle="Copy Selected Products To Categories"
end if
if session("cp_cmt_action")=2 then
pageTitle="Move Selected Products To Categories"
end if
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="xmlcat.asp"-->
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<%
Dim rs, conntemp, query
call opendb()

if request("action")="run" then
	Catlist=request("catlist")
	if Catlist<>"" then
		CatArr=split(catlist,",")
		PrdArr=split(session("cp_cmt_prdlist"),",")
		SourceCat=session("cp_cmt_idcat")
		For i=0 to ubound(PrdArr)
			if trim(PrdArr(i))<>"" then
			For j=0 to ubound(CatArr)
				if trim(CatArr(j))<>"" then
				query="SELECT idproduct FROM categories_products WHERE idProduct=" & PrdArr(i) & " AND idCategory=" & CatArr(j) & ";"
				set rs=connTemp.execute(query)
				if rs.eof then
					query="INSERT INTO categories_products (idProduct,idCategory) VALUES (" & PrdArr(i) & "," & CatArr(j) & ");"
					set rs=connTemp.execute(query)
				end if
				set rs=nothing
				end if
			Next
			If session("cp_cmt_action")=2 then
				query="DELETE FROM categories_products WHERE idProduct=" & PrdArr(i) & " AND idCategory=" & SourceCat & ";"
				set rs=connTemp.execute(query)
			End if
			end if
		Next
		set rs=nothing
		session("cp_cmt_prdlist")=""
		session("cp_cmt_idcat")=""
		if session("cp_cmt_action")="1" then
			msg="Selected products have been copied to new categories successfully!"
		end if
		if session("cp_cmt_action")="2" then
			msg="Selected products have been moved to new categories successfully!"
		end if
		session("cp_cmt_action")=""
	end if
end if
%>

<!--#include file="AdminHeader.asp"-->
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


		<% If msg <> "" Then %>
			<table class="pcCPcontent">        
				<tr>           
					<td align="center"><div class="pcCPmessage"><%=msg%></div></td>
				</tr>
                <tr>
                	<td>
                    	<br>
                        <br>
                    	<input type="button" name="Back" value="Back to manage categories" onclick="location='manageCategories.asp';" class="ibtnGrey">
                        <br>
                        <br>
					</td>
				</tr>
			</table>
		<% End If %>
        <%if request("action")<>"run" then%>
		<table class="pcCPcontentCat">
		<tr> 
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="3">
				<table id="FindProducts" class="pcCPcontent">
					<tr>
						<td>
						<%
							src_FormTitle1="Find Categories"
							src_FormTitle2=""
							src_FormTips1="Use the following filters to look for categories in your store to copy selected products to."
							src_FormTips2=""
							src_DisplayType=1
							src_ShowLinks=0
							src_ParentOnly=0
							src_FromPage="CopyMovePrds.asp?"
							src_ToPage="CopyMovePrds.asp?action=run"
							src_Button1=" Search "
							src_Button2=""
							src_Button3=" Back "
							src_PageSize=15
							UseSpecial=1
							session("srcCat_from")=""
							session("srcCat_where")=" AND categories.idcategory<>" & session("cp_cmt_idcat")
						%>
						<!--#include file="inc_srcCATs.asp"-->
						</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr> 
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3"><hr></td>
		</tr>
	</table>
    <%end if%>
<!--#include file="AdminFooter.asp"-->