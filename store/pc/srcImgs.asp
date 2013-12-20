<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/languages.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="inc_srcImgsQuery.asp"-->
<% pageTitle = getUserInput(request("src_FormTitle2"),0)

totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
  	response.redirect "techErr.asp?error="&Server.UrlEncode("Error in page srcImgs. Error: "&err.description)
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
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_key4=getUserInput(request("key4"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)
pshowimage=getUserInput(request("showimage"),0)
fid=request.QueryString("fid")
ffid=request.QueryString("ffid")
submit=request("Submit")
submit2=request("Submit2")


'--- End of Search Form parameters ---
Function URLDecode(tmpURL)
	Dim tmp1,tmpArr,i,icount	
	tmp1=tmpURL	
	if tmp1<>"" then
		tmp1=replace(tmp1,"+"," ")
		tmpArr=split(tmp1,"%")
		tmp1=tmpArr(0)
		icount=ubound(tmpArr)
		For i=1 to icount
			tmp1=tmp1 & Chr("&H" & Left(tmpArr(i),2)) & Right(tmpArr(i),len(tmpArr(i))-2)
		Next
	end if	
	URLDecode=tmp1
End Function

%>
<html>
<head>
<title>Locate an Image</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />

<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function openGalleryWindow(url) {
	if (document.all)
		var xMax = screen.width, yMax = screen.height;
	else
		if (document.layers)
			var xMax = window.outerWidth, yMax = window.outerHeight;
		else
			var xMax = 800, yMax=600;
	var xOffset = (xMax - 200)/2, yOffset = (yMax - 200)/2;
	var xOffset = 100, yOffset = 100;

	popupWin = window.open(url,'new_page','width=700,height=535,screenX='+xOffset+',screenY='+yOffset+',top='+yOffset+',left='+xOffset+',scrollbars=auto,toolbars=no,menubar=no,resizable=yes')
}
// done hiding -->
</script>
</head>
<body style="margin: 0;">
<div id="pcMain" style="background-color:#FFF; color:#000;">
<table class="pcMainTable" style="background-color:#FFF; color:#000;">

<% IF request("del")<>"" THEN  
        '=======================================
        ' START delete checked files
        '=======================================
	    IF session("admin")=-1 THEN '2 - Check for admin user
		    Count=request("dCount")
		    if Count>"0" then '3
			    pcv_TestP=0
			    pcv_ErrMsg=""

			    Set fso=server.CreateObject("Scripting.FileSystemObject")
			    PageName="catalog/testing.txt"
			    findit=Server.MapPath(PageName)
    			
			    Set f=fso.OpenTextFile(findit, 2, True)
			    f.Write "test done"
			    if Err.number>0 then
				    pcv_TestP=1
				    pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_10")
				    Err.number=0
			    end if
			    Set f=nothing

			    IF pcv_TestP=0 THEN
				    Set f=fso.GetFile(findit)
				    Err.number=0
				    f.Delete
				    if Err.number>0 then
					    pcv_TestP=1
					    pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_11")
					    Err.number=0
				    end if
			    END IF

			    Set f=nothing			
			    Set fso=nothing

			    IF pcv_TestP=0 THEN '4
			        pcv_ErrMsg =""
				    For i=1 to Count
					    if request("cimg"&i)<>"" then
						    Set fso = server.CreateObject("Scripting.FileSystemObject")
						    Set f = fso.GetFile(Server.MapPath(URLDecode(request("cimg"&i))))
						    Err.number=0
						    f.Attributes=vbArchive
						    f.Delete
					        if Err.number>0 then
					            pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_11")
					            Err.number=0
					            exit for
				            end if
					        call opendb()
					        fname=mid(request("cimg"&i),instr(1,request("cimg"&i),"/")+1)
							fname=URLDecode(fname)
                            query="DELETE from pcImageDirectory where pcImgDir_Name='"&fname&"'"
                            set rs=server.CreateObject("ADODB.RecordSet")
                            set rs=conntemp.execute(query)
                            set rs=nothing
					        call closedb()
				            Set f=nothing
				            Set fso = nothing
					        pcv_ErrMsg = pcv_ErrMsg & URLDecode(fname) & " has been deleted.<br>"
					    end if
				    Next
				    Set f=nothing
				    Set fso = nothing
			    END IF '4 %>
					<tr>
						<td>
			    		<div class="pcErrorMessage"><%=pcv_ErrMsg%></div>
						</td>
					</tr>
		        <%
		    end if '3
	    END IF '2
        '=======================================
        ' END delete checked files
        '=======================================
   'end if
%>    

<%
		elseIF rstemp.eof THEN
			Dim intNoResults
			intNoResults=1
%>
	<tr>
		<td>
			<div class="pcErrorMessage"><% response.write dictLanguage.Item(Session("language")&"_ShowSearch_5")%><br /><br /><a href="imageDir.asp?ffid=smallImageUrl&fid=hForm">Back</a></div>
		</td>
	</tr>

<% ELSE%>

	<%if src_FormTips2<>"" then%>
	<tr>
		<td><p><%=src_FormTips2%></p></td>
	</tr>
		
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
	};

	function runXML()
	{
		document.getElementById("runmsg").innerHTML="<% response.write dictLanguage.Item(Session("language")&"_ShowSearch_2")%>";
		document.body.style.cursor='progress';
		myConn.connect("xml_srcImgs.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
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
	<tr>
		<td><span id="runmsg"></span></td>
	</tr>
		
	
	<% 
    response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf
    for i=1 to iPageSize
        response.write "function setForm"&i&"() {"&vbCrlf
        response.write "opener.document."&fid&"."&ffid&".value = document.inputForm"&i&".inputField"&i&".value;"&vbCrlf
        response.write "opener.document."&fid&"."&ffid&".focus();"&vbCrlf
        response.write "self.close();"&vbCrlf
        response.write "return false;"&vbCrlf
        response.write "}"&vbCrlf
    next
    response.write "//--></SCRIPT>"&vbCrlf
    %>
		
	<tr>
		<td>
		
			<form name="ajaxSearch" class="pcForms">
				<input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
				<input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
				<input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
				<input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
				<input type="hidden" name="src_Button2" value="<%=src_Button2%>">
				<input type="hidden" name="src_Button3" value="<%=src_Button3%>">
				<input type="hidden" name="key1" value="<%=form_key1%>">
				<input type="hidden" name="key2" value="<%=form_key2%>">
				<input type="hidden" name="key3" value="<%=form_key3%>">
				<input type="hidden" name="key4" value="<%=form_key4%>">
				<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
				<input type="hidden" name="order" value="<%=form_order%>">
				<input type="hidden" name="fid" value="<%=fid%>">
				<input type="hidden" name="ffid" value="<%=ffid%>">
				<input type="hidden" name="showimage" value="<%=pshowimage%>">
				<input type="hidden" name="submit" value="<%=submit%>">
				<input type="hidden" name="submit2" value="<%=submit2%>">
				<input type="hidden" name="iPageCurrent" value="1">
				<input type="hidden" name="Imglist" value="">
			</form>

				<div id="resultarea"></div>
				<script>
				runXML();
				</script>
				<%
				pcv_HaveResults=1
				END IF
				set rstemp=nothing%>
				
				<script>
				<!--
				var savelist="xml,";
				
				function getImglist()
				{
				var tmp2=savelist;
				var pos=0;
					pos=tmp2.indexOf("xml,");
					var out="xml,";
					var temp = "" + (tmp2.substring(0, pos) + tmp2.substring((pos + out.length), tmp2.length));
					return(temp);
				}
				
				//-->
				</script>
				
				<% If iPageCount>1 and pcv_ErrMsg="" Then %>
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

			<%if src_FromPage<>"" and intNoResults<>1 then%>
            <div style="text-align: center; padding: 10px;">
			<p style="text-align: center;">
		     	<form class="pcForms">
					<input type="button" value="<%=src_Button3%>" onClick="location.href='<%=src_FromPage%>'">
				</form>
			</p>
            </div>
			<p>&nbsp;</p>
			<% end if%>
		</td>
	</tr>
</table>
</div>
</body>
</html>