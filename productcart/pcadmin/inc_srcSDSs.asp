<%
conntemp=""
call openDB()

'--- Generate Seach Parameters ---

if src_PageType="" then
	src_PageType=0
end if

if src_PageType=0 then
	pcv_Title="Supplier"
else
	pcv_Title="Drop-Shipper"
end if

if UseSpecial="1" then
else
	session("srcSDS_from")=""
	session("srcSDS_where")=""
end if	

if src_FormTitle1="" then
	src_FormTitle1="Search a " & pcv_Title
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
myConn.connect("xml_srcSDSsCount.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
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
		var tmpStr=theForm.key1.value+theForm.key2.value+theForm.key3.value+theForm.key4.value+theForm.key5.value+theForm.key6.value;
		if (tmpStr == "")
			{
				alert("Please fill out at least one field to run a search.");
				theForm.key1.focus();
				return (false);
			}
		return (true);
	}
	
//-->
</script>

<!--Begin Search Form-->

<form name="ajaxSearch" method="post" action="srcSDSs.asp" class="pcForms">
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
<input type="hidden" name="src_PageType" value="<%=src_PageType%>">
 
<table class="pcCPsearch">
	<tr>
		<td colspan="2">
		<span class="pcCPsectionTitle"><%=src_FormTitle1%></span>
		<%if src_FormTips1<>"" then%>
			<div><%=src_FormTips1%></div>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2"><span id="totalresults">&nbsp;</span></td>
	</tr>

	<tr> 
		<td width="31%">First name:</td>
		<td width="69%">
			<input type=text name="key1" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td width="31%">Last name:</td>
		<td width="69%">
			<input type=text name="key2" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr> 
		<td width="31%" align="left">Company:</td>
		<td width="69%">
			<input type=text name="key3" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td width="31%">Phone number:</td>
		<td width="69%">  
			<input type=text name="key5" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td width="31%" align="left">E-mail address:</td>
		<td>
			<input type=text name="key4" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td>Results per page:</td>
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
		<td nowrap>Sort by:</td>
		<td>
			<select name="order">
				<option value="2"><%=pcv_Title%> ID ascending</option>
				<option value="3"><%=pcv_Title%> ID descending</option>
				<option value="4" selected>Last Name ascending</option>
				<option value="5">Last Name descending</option>
			</select>
		<td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
    	<td></td>
		<td>
			<input name="Submit" type="submit" value="<%=src_Button1%>" id="searchSubmit">
            <input name="Submit1" type="submit" value="View All" id="searchSubmit" onclick="javascript:document.ajaxSearch.key1.value='';document.ajaxSearch.key2.value='';document.ajaxSearch.key3.value='';document.ajaxSearch.key4.value='';document.ajaxSearch.key5.value='';">
		</td>
	</tr>
</table>
</form>
<!--End of Search Form-->
<% call closeDb()%>