<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="inc_srcCATQuery.asp"-->
<% pageTitle = getUserInput(request("src_FormTitle2"),0) %>
<!--#include file="Adminheader.asp"--> 
<%totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
  	response.redirect "techErr.asp?error="&Server.UrlEncode("Error in page advSrcb. Error: "&err.description)
end If
iPageCount=0
if not rsTemp.eof then 
	totalrecords=clng(rstemp.RecordCount)
	iPageCount=rstemp.PageCount
end if

'--- Get Search Form parameters ---

src_FormTitle1=getUserInput(request("src_FormTitle1"),0)
src_FormTitle2=getUserInput(request("src_FormTitle2"),0)
src_FormTips1=getUserInput(request("src_FormTips1"),0)
src_FormTips2=getUserInput(request("src_FormTips2"),0)
src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
src_FromPage=getUserInput(request("src_FromPage"),0)
src_ToPage=getUserInput(request("src_ToPage"),0)
src_Button2=getUserInput(request("src_Button2"),0)
src_Button3=getUserInput(request("src_Button3"),0)
src_IncNotShDesc=getUserInput(request("src_IncNotShDesc"),0)
src_IncNotDisplay=getUserInput(request("src_IncNotDisplay"),0)
src_IncNotFRetail=getUserInput(request("src_IncNotFRetail"),0)
src_ParentOnly=getUserInput(request("src_ParentOnly"),0)
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_CatDiscType=getUserInput(request("CatDiscType"),0)
form_CatPromoType=getUserInput(request("CatPromoType"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)

'--- End of Search Form parameters ---
%>

<% IF rstemp.eof THEN %>
	<div class="pcCPmessage">
		Your search did not return any results.
	</div>
<% ELSE%>

	<%if src_FormTips2<>"" then%>
		<p><%=src_FormTips2%></p>
	<%end if%>

	<!--AJAX Functions-->
	<script type="text/javascript" src="XHConn.js"></script>

	<SCRIPT>
	var myConn = new XHConn();

	if (!myConn) alert("XMLHTTP not available. Try a newer/better browser.");

	var fnWhenDone = function (oXML) {
	var xmldoc = oXML.responseXML;
	var root_node = xmldoc.getElementsByTagName('data0').item(0);
	var tmpcount=parseInt(root_node.firstChild.data);
	var tmpdata="";
	if (tmpcount>0)
	{
		for (i=1;i<=tmpcount;i++)
		{
			var root_node = xmldoc.getElementsByTagName('data'+i).item(0);
			tmpdata=tmpdata+root_node.firstChild.data;
		}
		document.getElementById("resultarea").innerHTML=tmpdata;
	}

	document.getElementById("runmsg").innerHTML="";
	document.body.style.cursor='';
	<%if iPageCount>1 then%>
	document.getElementById("currentpage").innerHTML=document.ajaxSearch.iPageCurrent.value;
	<%end if%>
	<%if src_DisplayType="1" then%>
	document.getElementById("checkarea").innerHTML="<br /><a href='javascript:checkAll();'>Check All</a>&nbsp;|&nbsp;<a href='javascript:uncheckAll();'>Uncheck All</a>";
	<%end if%>
	presetF();
	};

	function runXML()
	{
	document.getElementById("runmsg").innerHTML="Please wait while we are processing your request ...";
		document.body.style.cursor='progress';
		myConn.connect("xml_srcCATs.asp", "GET", GetAllValues(document.ajaxSearch), fnWhenDone);
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
		<span id="runmsg"></span>
	<form name="ajaxSearch">
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
	<input type="hidden" name="key1" value="<%=form_key1%>">
	<input type="hidden" name="key2" value="<%=form_key2%>">
	<input type="hidden" name="key3" value="<%=form_key3%>">
	<input type="hidden" name="CatDiscType" value="<%=form_CatDiscType%>">
    <input type="hidden" name="CatPromoType" value="<%=form_CatPromoType%>">
	<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
	<input type="hidden" name="order" value="<%=form_order%>">
	<input type="hidden" name="iPageCurrent" value="1">
	<input type="hidden" name="catlist" value="">
	</form>
	<div id="resultarea"></div>
	<script>
	runXML();
	</script>
	<%
	pcv_HaveResults=1
END IF
set rstemp=nothing%>

<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= eval("document.srcresult.count.value"); j++)
{
	box = eval("document.srcresult.C" + j); 
	if (box.checked == false)
	{
		box.checked = true;
		updvalue(box);
	}
}
}

function uncheckAll() {
for (var j = 1; j <= eval("document.srcresult.count.value"); j++)
{
	box = eval("document.srcresult.C" + j); 
	if (box.checked == true)
	{
		box.checked = false;
		updvalue(box);
	}
}
}

//-->
</script>

<script>
var savelist="xml,";

function getcatlist()
{
var tmp2=savelist;
var pos=0;
	pos=tmp2.indexOf("xml,");
	var out="xml,";
	var temp = "" + (tmp2.substring(0, pos) + tmp2.substring((pos + out.length), tmp2.length));
	return(temp);
}
function updvalue(vfield)
{
var pos=0;
		var tmp1=vfield.value;
		pos=savelist.indexOf("," + tmp1 + ",");
		
<%if src_DisplayType="1" then%>		
		
		if (vfield.checked==true)
		{
			if (pos<0)
			{
				savelist=savelist + tmp1 + ",";
			}
		}
		else
		{
			if (pos>=0)
			{
			var out=tmp1+',';
					
			var temp = "" + (savelist.substring(0, pos) + savelist.substring((pos + out.length), savelist.length));
			savelist=temp;
			}
		}
<%else%>

		if (vfield.checked==true)
		{
			if (pos<0)
			{
				savelist="xml," + tmp1 + ",";
			}
		}
<%end if%>
}

function presetF()
{
var i=0;
var objElems = document.srcresult.elements;
var pos=0;
	for(i=0;i<objElems.length;i++)
	{
		var tmp1=objElems[i].value;
		pos=savelist.indexOf(tmp1 + ",");
		if (pos==0)
		{
		pos=1
		}
		else
		{
		pos=savelist.indexOf("," + tmp1 + ",");
		}
		if (pos>0)
		{
			objElems[i].checked=true;
		}
	}
}
</script>

<p>
	<span id="checkarea"></span>
</p>

<% If iPageCount>1 Then %>
	<p class="pcPageNav">
	Currently viewing page <span id="currentpage">1</span> of <%=iPageCount%><br>
		<%For I = 1 To iPageCount%>
		<a href="javascript:document.ajaxSearch.iPageCurrent.value='<%=I%>';runXML();"><%=I%></a> 
		<% Next %>
	</p>
<% End If %>

<% 
set rstemp=nothing
call closeDb()
%>

<p style="text-align: center;">
	<%if (src_ToPage<>"") and (pcv_HaveResults="1") and (src_DisplayType>"0") then
			if instr(src_ToPage,"?")<=0 then
			src_ToPage=src_ToPage & "?"
			else
				if Right(src_ToPage,1)<>"?" then
				src_ToPage=src_ToPage & "&"
				end if
			end if%>
		<input type="button" value="<%=src_Button2%>" onClick="javascript:if (savelist=='xml,') {alert('Please select a category before clicking on this button');} else {location.href='<%=src_ToPage%>catlist='+getcatlist(savelist);}" class="ibtnGrey">
	<%end if%>
	<%if src_FromPage<>"" then%>
	&nbsp;<input type="button" value="<%=src_Button3%>" onClick="location.href='<%=src_FromPage%>'" class="ibtnGrey">
	<%end if%>
</p>
<!--#include file="Adminfooter.asp"-->