<%'**** Check MAC IE Browser ******************

UserBrowser=Request.ServerVariables("HTTP_USER_AGENT")

MACIEBrowser=0
OperaBrowser=0

if UserBrowser<>"" then
	if instr(ucase(UserBrowser),"MAC")>0 AND instr(ucase(UserBrowser),"MSIE")>0 then
		MACIEBrowser=1
	end if
	
	if instr(ucase(UserBrowser),"OPERA")>0 then
		OperaBrowser=1
	end if
end if

'**** End of Check MAC IE Browser ******************%>

<%IF MACIEBrowser=1 OR OperaBrowser=1 OR UserBrowser="" THEN%>
	<script>
	function runXML(new_callID)
	{
	}
	
	function runXML1(new_callID)
	{
	}
	
	function hidetip()
	{
	}
</script>
<%session("store_useAjax")="1"%>
<%ELSE%>
<%session("store_useAjax")=""%>
<link type="text/css" rel="stylesheet" href="../pc/ei-tooltip.css">

<script language="javascript" type="text/javascript" src="../pc/ei-tooltip.js"></script>

<script>
var sav_title=""
var sav_content=""
var sav_callxml=""
var save_callID=""
var sav_callID= new Array(100);
var sav_btitle= new Array(100);
var sav_bcontent= new Array(100);
var sav_bcount=-1;

var sav_CatPretitle=""
var sav_CatPrecontent=""
var sav_CatPrecallxml=""
var save_CatcallID=""
var sav_CatPrecallID= new Array(100);
var sav_CatPrebtitle= new Array(100);
var sav_CatPrebcontent= new Array(100);
var sav_CatPrebcount=-1;
</script>

<!--AJAX Functions-->
<script type="text/javascript" src="../pc/XHConn.js"></script>

<SCRIPT>
var myConn = new XHConn();

if (!myConn) alert("XMLHTTP not available. Try a newer/better browser.");

var fnWhenDoneCat = function (oXML) {
var xmldoc = oXML.responseXML;
var root_node = xmldoc.getElementsByTagName('bcontent').item(0);
var xml_data=root_node.firstChild.data;
sav_CatPretitle=""
sav_CatPrecontent="";
if (xml_data=="nothing")
{
	sav_CatPretitle="";
	sav_CatPrecontent="";
	hidetip();
}
else
{
	if (xml_data.indexOf("|||")>=0)
	{
		var tmp_data=xml_data.split("|||");
		sav_CatPretitle=tmp_data[0];
		sav_CatPrecontent=tmp_data[1];
	}
	else
	{
		sav_CatPretitle="<%=dictLanguage.Item(Session("language")&"_ShowSearch_8")%>";
		sav_CatPrecontent=xml_data;
	}
	if (((sav_CatPretitle!='') || (sav_CatPrecontent!='')) && (sav_CatPrecallxml=="1"))
	{
		sav_CatPrebcount=sav_CatPrebcount+1;
		sav_CatPrecallID[sav_CatPrebcount]=save_CatcallID;
		sav_CatPrebtitle[sav_CatPrebcount]=sav_CatPretitle;
		sav_CatPrebcontent[sav_CatPrebcount]=sav_CatPrecontent;
		showtip(sav_CatPretitle,sav_CatPrecontent);
	}
	else
	{
		sav_CatPretitle="";
		sav_CatPrecontent="";
		hidetip();
	}
}

};

function runPreCatXML(new_callID)
{
save_CatcallID=new_callID
if (sav_CatPrecallxml=="1")
{
var x=0; 
for (x=0; x<sav_CatPrebcount; x++)
{
	if (""+new_callID==""+sav_CatPrecallID[x])
	{
		sav_CatPretitle=sav_CatPrebtitle[x];
		sav_CatPrecontent=sav_CatPrebcontent[x];
		hidetip();
		showtip(sav_CatPretitle,sav_CatPrecontent);
		sav_CatPrecallxml="";
		return(true);
	}
}
myConn.connect("xml_getCatPreInfo.asp", "POST", GetAllValues(document.getCatPre), fnWhenDoneCat);
}
return(true);
}

var fnWhenDone = function (oXML) {
var xmldoc = oXML.responseXML;
var root_node = xmldoc.getElementsByTagName('bcontent').item(0);
var xml_data=root_node.firstChild.data;
sav_title=""
sav_content="";
if (xml_data=="nothing")
{
	sav_title="";
	sav_content="";
	hidetip();
}
else
{
	if (xml_data.indexOf("|||")>=0)
	{
		var tmp_data=xml_data.split("|||");
		sav_title=tmp_data[0];
		sav_content=tmp_data[1];
	}
	else
	{
		sav_title="<%=dictLanguage.Item(Session("language")&"_ShowSearch_8")%>";
		sav_content=xml_data;
	}
	if (((sav_title!='') || (sav_content!='')) && (sav_callxml=="1"))
	{
		sav_bcount=sav_bcount+1;
		sav_callID[sav_bcount]=save_callID;
		sav_btitle[sav_bcount]=sav_title;
		sav_bcontent[sav_bcount]=sav_content;
		showtip(sav_title,sav_content);
	}
	else
	{
		sav_title="";
		sav_content="";
		hidetip();
	}
}

};

function runXML(new_callID)
{
save_callID=new_callID
if (sav_callxml=="1")
{
var x=0; 
for (x=0; x<sav_bcount; x++)
{
	if (""+new_callID==""+sav_callID[x])
	{
		sav_title=sav_btitle[x];
		sav_content=sav_bcontent[x];
		hidetip();
		showtip(sav_title,sav_content);
		sav_callxml="";
		return(true);
	}
}
myConn.connect("xml_srcListPrds.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
}
return(true);
}

function runXML1(new_callID)
{
save_callID=new_callID
if (sav_callxml=="1")
{
var x=0; 
for (x=0; x<sav_bcount; x++)
{
	if (""+new_callID==""+sav_callID[x])
	{
		sav_title=sav_btitle[x];
		sav_content=sav_bcontent[x];
		hidetip();
		showtip(sav_title,sav_content);
		sav_callxml="";
		return(true);
	}
}
myConn.connect("xml_getPrdInfo.asp", "POST", GetAllValues(document.getPrd), fnWhenDone);
}
return(true);
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
<%END IF%>
<form name="getPrd" style="margin:0px;">
	<input type="hidden" name="idproduct" value="">
</form>
<form name="getCatPre" style="margin:0px;">
	<input type="hidden" name="idcategory" value="">
</form>

<!--End of AJAX Functions-->

