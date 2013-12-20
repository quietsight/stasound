<%
conntemp=""
call openDB()

'--- Generate Seach Parameters ---

if UseSpecial="1" then
else
	session("srcCat_from")=""
	session("srcCat_where")=""
	session("srcCat_DiscArea")=""
end if	

if src_FormTitle1="" then
	src_FormTitle1="Search a category"
end if

if src_FormTitle2="" then
	src_FormTitle2="Search Results"
end if

if src_DisplayType="" then
	src_DisplayType="0"
end if

if src_ShowLinks="" then
	src_ShowLinks="0"
end if

if src_Button1="" then
	src_Button1=" Search "
end if

if src_Button2="" then
	src_Button2=" Continue "
end if

if src_Button3="" then
	src_Button3=" New Search "
end if

if src_IncNotShDesc="" then
	src_IncNotShDesc="0"
end if

if src_IncNotDisplay="" then
	src_IncNotDisplay="0"
end if

if src_IncNotFRetail="" then
	src_IncNotFRetail="0"
end if

if src_ParentOnly="" then
	src_ParentOnly="0"
end if

'--- End of Generate Seach Parameters ---

%>

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
	

</SCRIPT>

<!--End of AJAX Functions-->

<!--Validate Form-->
<script language="JavaScript">
<!--
	
function FormValidator(theForm)
	{
		var tmpStr=theForm.key1.value+theForm.key2.value+theForm.key3.value;
		if (tmpStr == "")
			{
				alert("Please enter at least one information for searching.");
				theForm.key1.focus();
				return (false);
			}
		return (true);
	}
	
//-->
</script>

<!--Begin Search Form-->

<form name="ajaxSearch" method="post" action="srcCATs.asp" class="pcForms">

<input type="hidden" name="referpage" value="NewSearch">
<input type="hidden" name="src_FormTitle1" value="<%=src_FormTitle1%>">
<input type="hidden" name="src_FormTitle2" value="<%=src_FormTitle2%>">
<input type="hidden" name="src_FormTips1" value="<%=src_FormTips1%>">
<input type="hidden" name="src_FormTips2" value="<%=src_FormTips2%>">
<input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
<input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
<input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
<input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
<input type="hidden" name="src_Button2" value="<%=src_Button2%>">
<input type="hidden" name="src_Button3" value="<%=src_Button3%>">
<input type="hidden" name="src_IncNotShDesc" value="<%=src_IncNotShDesc%>">
<input type="hidden" name="src_IncNotDisplay" value="<%=src_IncNotDisplay%>">
<input type="hidden" name="src_IncNotFRetail" value="<%=src_IncNotFRetail%>">
<input type="hidden" name="src_ParentOnly" value="<%=src_ParentOnly%>">
 
<input type="hidden" name="idSupplier" value="10">
<table class="pcCPsearch">
	<tr>
		<td colspan="2">
		<h2><%=src_FormTitle1%></h2>
		<%if src_FormTips1<>"" then%>
			<p><%=src_FormTips1%></p>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2"><span id="totalresults">&nbsp;</span></td>
	</tr>

	<tr> 
		<td width="20%" align="right">Category Name:</td>
		<td width="80%">  
			<input type=text name="key1" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr> 
		<td align="right">Short Description:</td>
		<td>  
			<input type=text name="key2" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td align="right">Long Description:</td>
		<td>  
			<input type=text name="key3" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>

	<% if src_ShowDiscTypes=1 then %>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<tr>
		<td valign="top" align="right">Filter by Discount:</td>
		<td valign="top"><input type="radio" name="CatDiscType" value="1" onclick="javascript:runXML();" class="clearBorder"> Only show categories with quantity discounts<br>
		<input type="radio" name="CatDiscType" value="2" onclick="javascript:runXML();" class="clearBorder"> Only show categories without quantity discounts<br>
		<input type="radio" name="CatDiscType" value="0" onclick="javascript:runXML();" checked class="clearBorder"> Include both categories in the search
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<% end if %>
    
    <% if src_ShowPromoTypes=1 then %>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<tr>
		<td valign="top" align="right">Filter by Promotion:</td>
		<td valign="top"><input type="radio" name="CatPromoType" value="1" onclick="javascript:runXML();" class="clearBorder"> Only show categories with promotion<br>
		<input type="radio" name="CatPromoType" value="2" onclick="javascript:runXML();" class="clearBorder"> Only show categories without promotion<br>
		<input type="radio" name="CatPromoType" value="0" onclick="javascript:runXML();" checked class="clearBorder"> Include both categories in the search
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<% end if %>
	
	<tr>
		<td align="right">Results per page:</td>
		<td> 
			<select name="resultCnt" id="resultCnt">
				<option value="5" <%if src_PageSize="5" then%>selected<%end if%>>5</option>
				<option value="10" <%if src_PageSize="10" then%>selected<%end if%>>10</option>
				<option value="15" <%if src_PageSize="15" then%>selected<%end if%>>15</option>
				<option value="20" <%if src_PageSize="20" then%>selected<%end if%>>20</option>
				<option value="25" <%if src_PageSize="25" then%>selected<%end if%>>25</option>
				<option value="50" <%if src_PageSize="50" then%>selected<%end if%>>50</option>
				<option value="100" <%if src_PageSize="100" then%>selected<%end if%>>100</option>
			</select>
		</td>
	</tr>
	<tr> 
		<td align="right">Sort by:</td>
		<td>
			<select name="order">
				<option value="2" selected>Name ascending</option>
				<option value="3">Name descending</option>
			</select>
		<td>
	</tr>

	<tr>
		<td colspan="2" align="center">&nbsp;</td>
	</tr>
	<tr class="normal"> 
		<td colspan="2" align="center">  
			<br>
			<input name="Submit" type="submit" value="<%=src_Button1%>" id="searchSubmit">
		</td>
	</tr>
</table>
</form>
<!--End of Search Form-->
<% call closeDb()%>