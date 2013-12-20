<%
conntemp=""
call openDB()

'--- Generate Seach Parameters ---

if UseSpecial="1" then
else
	session("srcImg_from")=""
	session("srcImg_where")=""
end if	

if src_FormTitle1="" then
	'src_FormTitle1="Search for an Image"
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
document.getElementById("totalresults").innerHTML="<% response.write dictLanguage.Item(Session("language")&"_ShowSearch_3")%>" + root_node.firstChild.data + "<% response.write dictLanguage.Item(Session("language")&"_ShowSearch_4")%>";
document.body.style.cursor='';
};

function runXML()
{
document.getElementById("totalresults").innerHTML="<% response.write dictLanguage.Item(Session("language")&"_ShowSearch_2")%>";
document.body.style.cursor='progress';
myConn.connect("xml_srcImgsCount.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
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
	
function check_date(field)
{
    var checkstr = "0123456789";
    var DateField = field;
    var Datevalue = "";
    var DateTemp = "";
    var seperator = "/";
    var day;
    var month;
    var year;
    var leap = 0;
    var err = 0;
    var i;
    err = 0;
    DateValue = DateField.value;
    /* Delete all chars except 0..9 */
    for (i = 0; i < DateValue.length; i++) 
    {
        if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) 
        {
            DateTemp = DateTemp + DateValue.substr(i,1);
        }
        else
        {
            if (DateTemp.length == 1)
            {
                DateTemp = "0" + DateTemp
            }
            else
            {
                if (DateTemp.length == 3)
                {
                    DateTemp = DateTemp.substr(0,2) + '0' + DateTemp.substr(2,1);
                }
            }
        }
    }
    DateValue = DateTemp;
    /* Always change date to 8 digits - string*/
    /* if year is entered as 2-digit / always assume 20xx */
    if (DateValue.length == 6) 
    {
        DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); 
    }
    if (DateValue.length != 8) 
    {
        return(false);
    }
    /* year is wrong if year = 0000 */
    year = DateValue.substr(4,4);
    if (year == 0) 
    {
        err = 20;
    }
    /* Validation of month*/
    <%if scDateFrmt="DD/MM/YY" then%>
        month = DateValue.substr(2,2);
    <%else%>
        month = DateValue.substr(0,2);
    <%end if%>
    if ((month < 1) || (month > 12)) 
    {
        err = 21;
    }
    /* Validation of day*/
    <%if scDateFrmt="DD/MM/YY" then%>
        day = DateValue.substr(0,2);
    <%else%>
        day = DateValue.substr(2,2);
    <%end if%>
    if (day < 1) 
    {
        err = 22;
    }
    /* Validation leap-year / february / day */
    if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) 
    {
        leap = 1;
    }
    if ((month == 2) && (leap == 1) && (day > 29)) 
    {
        err = 23;
    }
    if ((month == 2) && (leap != 1) && (day > 28)) 
    {
        err = 24;
    }
    /* Validation of other months */
    if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) 
    {
        err = 25;
    }
    if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) 
    {
        err = 26;
    }
    /* if 00 ist entered, no error, deleting the entry */
    if ((day == 0) && (month == 0) && (year == 00)) 
    {
        err = 0; day = ""; month = ""; year = ""; seperator = "";
    }
    /* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
    if (err == 0) 
    {
        <%if scDateFrmt="DD/MM/YY" then%>
            DateField.value = day + seperator + month + seperator + year;
        <%else%>
            DateField.value = month + seperator + day + seperator + year;   
        <%end if%>  
        return(true);
    }
    /* Error-message if err != 0 */
    else 
    {
        return(false);   
    }
}


function FormValidator(theForm)
{
    if (theForm.key4.value == "")	
    {
        return (true);
    }
	
	if (check_date(theForm.key4) == false)
  	{
		alert("Please enter a valid date for this field.");
	    theForm.key4.focus();
	    return (false);
	}
	
	return (true);
}
	
//-->
</script>

<!--Begin Search Form-->

<form name="ajaxSearch" method="post" action="srcImgs.asp?fid=<%=fid%>&ffid=<%=ffid%>" onSubmit="return FormValidator(this)" class="pcForms">
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
 
<table class="pcCPsearch" style="color: #000;">
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
		<td width="31%">Image type:</td>
		<td width="69%">
			<select name="key2" size="1" onblur="javascript:runXML();">
				<option value="" selected="selected">Any</option>
				<option value="gif">Gif</option>
				<option value="jpe">Jpeg</option>
                <option value="png">Png</option>
			</select>
		</td>
	</tr>

	<tr> 
		<td width="31%">Image name:</td>
		<td width="69%">
			<input type=text name="key1" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr> 
		<td width="31%" align="left">Maximum Image size (bytes):</td>
		<td width="69%">
			<input type=text name="key3" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	
	<tr>
		<td width="31%" align="left">Images Uploaded Since (eg. 06/01/2005):</td>
		<td>
			<input type=text name="key4" size="50" value="" onblur="javascript:runXML();">
		</td>
	</tr>
	<tr>
		<td>Results per page:</td>
		<td> 
			<select name="resultCnt" id="resultCnt">
				<option value="5">5</option>
				<option value="10">10</option>
				<option value="15" selected>15</option>
				<option value="20">20</option>
				<option value="25">25</option>
				<option value="50">50</option>
				<option value="100">100</option>
				<%if src_PageSize<>"" then%>
				<option value="<%=src_PageSize%>" selected><%=src_PageSize%></option>
				<%end if%>
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>Sort by:</td>
		<td>
			<select name="order">
				<option value="2" selected>Image Name ascending</option>
				<option value="3">Image Name descending</option>
				<option value="4">Image Type ascending</option>
				<option value="5">Image Type descending</option>
				<option value="6">Image Size ascending</option>
				<option value="7">Image Size descending</option>
				<option value="8">Date Uploaded ascending</option>
				<option value="9">Date Uploaded descending</option>
			</select>
		<td>
	</tr>
	<tr>
		<td colspan="2" align="center">&nbsp;</td>
	</tr>

    <tr>
        <td colspan="2">Would you like to <b>display image thumbnails</b> on the selection page?</td>
    </tr>    
    <tr>
        <td colspan="2">
				<input type="radio" name="showimage" value="NO" class="clearBorder">No&nbsp;
				<input type="radio" name="showimage" value="YES" checked class="clearBorder">Yes
		</td>
    </tr>

	<tr>
		<td colspan="2" align="center">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<br>
			<input name="Submit" type="submit" class="submit2" value="<%=src_Button1%>" id="searchSubmit" />
			<input name="Close" type="button" value="Close" onclick="JavaScript:window.close()" />
			<% if session("srcImg_query") <> "" then %>
			    <input name="Submit2" type="submit" value="Repeat Last Search" id="Submit2" />
			<% end if %>
		</td>
	</tr>
</table>
</form>
<!--End of Search Form-->
<% call closeDb()%>