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
<!--#include file="inc_srcPrdQuery.asp"-->
<% 
pcStrPageName="srcPrds.asp"
pageTitle = getUserInput(request("src_FormTitle2"),0)
src_PageType=getUserInput(request("src_PageType"),0)
'Start SDBA
if src_PageType="" then
	src_PageType=0
end if
src_IDSDS=getUserInput(request("src_IDSDS"),0)
src_sdsAssign=getUserInput(request("src_sdsAssign"),0)
if (src_IDSDS<>"") and (src_IDSDS<>"0") then
	if src_sdsAssign="1" then
		if src_PageType="0" then
			pageTitle=trim(pageTitle) & " Supplier"
		else
			pageTitle=trim(pageTitle) & " Drop-Shipper"
		end if
	end if
end if
'End SDBA
%>
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
IF request("action")="newsrc" THEN
	src_FormTitle1=getUserInput(request("src_FormTitle1"),0)
	src_FormTitle2=getUserInput(request("src_FormTitle2"),0)
	src_FormTips1=getUserInput(request("src_FormTips1"),0)
	src_FormTips2=getUserInput(request("src_FormTips2"),0)
	src_IncDown=getUserInput(request("src_IncDown"),0)
	src_IncGC=getUserInput(request("src_IncGC"),0)
	src_SM=getUserInput(request("src_SM"),0)
	src_IncNormal=getUserInput(request("src_IncNormal"),0)
	if src_IncNormal="" then
		src_IncNormal=0
	end if
	src_IncBTO=getUserInput(request("src_IncBTO"),0)
	if src_IncBTO="" then
		src_IncBTO=0
	end if
	src_IncItem=getUserInput(request("src_IncItem"),0)
	if src_IncItem="" then
		src_IncItem=0
	end if
	if (src_IncBTO="0") AND (src_IncItem="0") AND (src_IncDown="") AND (src_IncGC="") then
		src_IncNormal="1"
	end if
	
	src_Special=getUserInput(request("src_Special"),0)
	src_Featured=getUserInput(request("src_Featured"),0)
	src_DisplayType=getUserInput(request("src_DisplayType"),0)
	src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
	src_FromPage=getUserInput(request("src_FromPage"),0)
	src_ToPage=getUserInput(request("src_ToPage"),0)
	src_Button2=getUserInput(request("src_Button2"),0)
	src_Button3=getUserInput(request("src_Button3"),0)
	src_DiscType=getUserInput(request("src_DiscType"),0)
	src_PromoType=getUserInput(request("src_PromoType"),0)
	
	'Start SDBA
	src_PageType=getUserInput(request("src_PageType"),0)
	src_IDSDS=getUserInput(request("src_IDSDS"),0)
	src_IsDropShipper=getUserInput(request("src_IsDropShipper"),0)
	src_sdsAssign=getUserInput(request("src_sdsAssign"),0)
	src_sdsStockAlarm=getUserInput(request("src_sdsStockAlarm"),0)
	'End SDBA

	form_idcategory=getUserInput(request("idcategory"),0)
	if form_idcategory="" then
		form_idcategory="0"
	end if
	form_customfield=getUserInput(request("customfield"),0)
	if form_customfield="" then
		form_customfield="0"
	end if
	form_SearchValues=getUserInput(request("SearchValues"),0)
	form_priceFrom=getUserInput(request("priceFrom"),0)
	if form_priceFrom="" then
		form_priceFrom="0"
	end if
	form_priceUntil=getUserInput(request("priceUntil"),0)
	if form_priceUntil="" then
		form_priceUntil="9999999"
	end if
	form_withstock=getUserInput(request("withstock"),0)
	form_stocklevel=getUserInput(request("stocklevel"),0)
	form_sku=getUserInput(request("sku"),0)
	form_IDBrand=getUserInput(request("IDBrand"),0)
	form_keyWord=getUserInput(request("keyWord"),0)
	form_exact=getUserInput(request("exact"),0)
	form_pinactive=getUserInput(request("pinactive"),0)
	form_notforsale=getUserInput(request("notforsale"),4)
	form_resultCnt=getUserInput(request("resultCnt"),0)
	form_order=getUserInput(request("order"),0)
	
	'Save values to sessions
	session("cp_lct_src_FormTitle1")=src_FormTitle1
	session("cp_lct_src_FormTitle2")=src_FormTitle2
	session("cp_lct_src_FormTips1")=src_FormTips1
	session("cp_lct_src_FormTips2")=src_FormTips2
	session("cp_lct_src_IncNormal")=src_IncNormal
	session("cp_lct_src_IncBTO")=src_IncBTO
	session("cp_lct_src_IncItem")=src_IncItem
	session("cp_lct_src_IncDown")=src_IncDown
	session("cp_lct_src_IncGC")=src_IncGC
	session("cp_lct_src_SM")=src_SM
	session("cp_lct_src_Special")=src_Special
	session("cp_lct_src_Featured")=src_Featured
	session("cp_lct_src_DisplayType")=src_DisplayType
	session("cp_lct_src_ShowLinks")=src_ShowLinks
	session("cp_lct_src_FromPage")=src_FromPage
	session("cp_lct_src_ToPage")=src_ToPage
	session("cp_lct_src_Button2")=src_Button2
	session("cp_lct_src_Button3")=src_Button3
	session("cp_lct_src_DiscType")=src_DiscType
	session("cp_lct_src_PromoType")=src_PromoType
	
	'Start SDBA
	session("cp_lct_src_PageType")=src_PageType
	session("cp_lct_src_IDSDS")=src_IDSDS
	session("cp_lct_src_IsDropShipper")=src_IsDropShipper
	session("cp_lct_src_sdsAssign")=src_sdsAssign
	session("cp_lct_src_sdsStockAlarm")=src_sdsStockAlarm
	'End SDBA

	session("cp_lct_form_idcategory")=form_idcategory
	session("cp_lct_form_customfield")=form_customfield
	session("cp_lct_form_SearchValues")=form_SearchValues
	session("cp_lct_form_priceFrom")=form_priceFrom
	session("cp_lct_form_priceUntil")=form_priceUntil
	session("cp_lct_form_withstock")=form_withstock
	session("cp_lct_form_stocklevel")=form_stocklevel
	session("cp_lct_form_sku")=form_sku
	session("cp_lct_form_IDBrand")=form_IDBrand
	session("cp_lct_form_keyWord")=form_keyWord
	session("cp_lct_form_exact")=form_exact
	session("cp_lct_form_pinactive")=form_pinactive
	session("cp_lct_form_notforsale")=form_notforsale
	session("cp_lct_form_resultCnt")=form_resultCnt
	session("cp_lct_form_order")=form_order
	
ELSE

	src_FormTitle1=session("cp_lct_src_FormTitle1")
	src_FormTitle2=session("cp_lct_src_FormTitle2")
	src_FormTips1=session("cp_lct_src_FormTips1")
	src_FormTips2=session("cp_lct_src_FormTips2")
	src_IncNormal=session("cp_lct_src_IncNormal")
	src_IncBTO=session("cp_lct_src_IncBTO")
	src_IncItem=session("cp_lct_src_IncItem")
	src_IncDown=session("cp_lct_src_IncDown")
	src_IncGC=session("cp_lct_src_IncGC")
	src_Special=session("cp_lct_src_Special")
	src_Featured=session("cp_lct_src_Featured")
	src_DisplayType=session("cp_lct_src_DisplayType")
	src_ShowLinks=session("cp_lct_src_ShowLinks")
	src_FromPage=session("cp_lct_src_FromPage")
	src_ToPage=session("cp_lct_src_ToPage")
	src_Button2=session("cp_lct_src_Button2")
	src_Button3=session("cp_lct_src_Button3")
	src_DiscType=session("cp_lct_src_DiscType")
	src_PromoType=session("cp_lct_src_PromoType")
		
	'Start SDBA
	src_PageType=session("cp_lct_src_PageType")
	src_IDSDS=session("cp_lct_src_IDSDS")
	src_IsDropShipper=session("cp_lct_src_IsDropShipper")
	src_sdsAssign=session("cp_lct_src_sdsAssign")
	src_sdsStockAlarm=session("cp_lct_src_sdsStockAlarm")
	'End SDBA

	form_idcategory=session("cp_lct_form_idcategory")
	form_customfield=session("cp_lct_form_customfield")
	form_SearchValues=session("cp_lct_form_SearchValues")
	form_priceFrom=session("cp_lct_form_priceFrom")
	form_priceUntil=session("cp_lct_form_priceUntil")
	form_withstock=session("cp_lct_form_withstock")
	form_stocklevel=session("cp_lct_form_stocklevel")
	form_sku=session("cp_lct_form_sku")
	form_IDBrand=session("cp_lct_form_IDBrand")
	form_keyWord=session("cp_lct_form_keyWord")
	form_exact=session("cp_lct_form_exact")
	form_pinactive=session("cp_lct_form_pinactive")
	form_notforsale=session("cp_lct_form_notforsale")
	form_resultCnt=session("cp_lct_form_resultCnt")
	form_order=session("cp_lct_form_order")

END IF

'--- End of Search Form parameters ---
%>
	<%if (src_IDSDS<>"") and (src_IDSDS<>"0") then%>
		<form name="searchprds" method="post" action="sds_assignprds.asp">
		<input type="hidden" name="src_FormTitle1" value="">
		<input type="hidden" name="src_FormTitle2" value="Assign Products to the ">
		<input type="hidden" name="src_FormTips1" value="">
		<input type="hidden" name="src_FormTips2" value="Select one or more products that you want to add to the ">
		<input type="hidden" name="src_IncNormal" value="<%=src_IncNormal%>">
		<input type="hidden" name="src_IncBTO" value="<%=src_IncBTO%>">
		<input type="hidden" name="src_IncItem" value="<%=src_IncItem%>">
		<input type="hidden" name="src_Special" value="<%=src_Special%>">
		<input type="hidden" name="src_Featured" value="<%=src_Featured%>">
		<input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
		<input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
		<input type="hidden" name="src_FromPage" value="javascript:history.back()">
		<input type="hidden" name="src_ToPage" value="">
		<input type="hidden" name="src_Button2" value="Assign Products">
		<input type="hidden" name="src_Button3" value=" New Search ">
		<input type="hidden" name="src_DiscType" value="<%=src_DiscType%>">
        <input type="hidden" name="src_PromoType" value="<%=src_PromoType%>">
				
		<%'Start SDBA%>
		<input type="hidden" name="src_PageType" value="<%=src_PageType%>">
		<input type="hidden" name="src_IDSDS" value="<%=src_IDSDS%>">
		<input type="hidden" name="src_IsDropShipper" value="<%=src_IsDropShipper%>">
		<input type="hidden" name="src_sdsAssign" value="1">
		<input type="hidden" name="src_sdsStockAlarm" value="<%=src_sdsStockAlarm%>">	
		<%'End SDBA%>
	
		<input type="hidden" name="idcategory" value="<%=form_idcategory%>">
		<input type="hidden" name="customfield" value="<%=form_customfield%>">
		<input type="hidden" name="SearchValues" value="<%=form_SearchValues%>">
		<input type="hidden" name="priceFrom" value="<%=form_priceFrom%>">
		<input type="hidden" name="priceUntil" value="<%=form_priceUntil%>">
		<input type="hidden" name="withstock" value="<%=form_withstock%>">
        <input type="hidden" name="stocklevel" value="<%=form_stocklevel%>">
		<input type="hidden" name="sku" value="<%=form_sku%>">
		<input type="hidden" name="IDBrand" value="<%=form_IDBrand%>">
		<input type="hidden" name="keyWord" value="<%=form_keyWord%>">
		<input type="hidden" name="exact" value="<%=form_exact%>">
		<input type="hidden" name="pinactive" value="<%=form_pinactive%>">
        <input type="hidden" name="notforsale" value="<%=form_notforsale%>">
		<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
		<input type="hidden" name="order" value="<%=form_order%>">
		<input type="hidden" name="iPageCurrent" value="1">
		<input type="hidden" name="prdlist" value="">
		</form>
		<script>
		function AssignPrds()
		{
			idsds=document.searchprds.src_IDSDS.value;
			isdropshipper=document.searchprds.src_IsDropShipper.value;
			PageType=document.searchprds.src_PageType.value;
			document.searchprds.src_ToPage.value="sds_assignprds.asp?action=add&idsds="+idsds+"&pagetype=" + PageType + "&isdropshipper="+isdropshipper;
			document.searchprds.submit();
		}
		</script>
	<%end if%>
<% IF rstemp.eof THEN %>
		<div class="pcCPmessage">
			Your search did not return any results.
			<%if (src_IDSDS<>"") and (src_IDSDS<>"0") then%>
			<a href="javascript:AssignPrds();">Assign products to this <%if src_PageType="0" then%>supplier<%else%>drop-shipper<%end if%></a>
			<%end if%>
        </div>
<% ELSE%>
	
	<%if (src_IDSDS<>"") and (src_IDSDS<>"0") then%>
	<p>
		<span class="pcCPsectionTitle">
		<%if src_PageType="0" then%>
			<%="Supplier: "%>
			<%query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
			set rs1=ConnTemp.execute(query)%>
			<%=rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"%>
			<%set rs1=nothing
		else%>
			<%="Drop-Shipper: "%>
			<%if src_IsDropShipper="1" then
				query="SELECT pcSupplier_Company,pcSupplier_FirstName,pcSupplier_LastName FROM pcSuppliers WHERE pcSupplier_ID=" & src_IDSDS
				set rs1=ConnTemp.execute(query)%>
				<%=rs1("pcSupplier_Company") & " (" & rs1("pcSupplier_FirstName") & " " & rs1("pcSupplier_LastName") & ")"%>
				<%set rs1=nothing
			else
				query="SELECT pcDropShipper_Company,pcDropShipper_FirstName,pcDropShipper_LastName FROM pcDropShippers WHERE pcDropShipper_ID=" & src_IDSDS
				set rs1=ConnTemp.execute(query)%>
				<%=rs1("pcDropShipper_Company") & " (" & rs1("pcDropShipper_FirstName") & " " & rs1("pcDropShipper_LastName") & ")"%>
				<%set rs1=nothing
			end if%>
		<%end if%>
		</span>
	</p>
	<%end if%>

	<%if src_FormTips2<>"" then%>
		<p><%if (src_IDSDS<>"") and (src_IDSDS<>"0") then
				if src_PageType="0" then
					%><%=trim(src_FormTips2) & " supplier:"%>
				<%else%>
					<%=trim(src_FormTips2) & " drop-shipper:"%>
				<%end if
			else%>
				<%=src_FormTips2%>
			<%end if%>
		</p>
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
	
	var firstTimeRun=1;
	
	function runXML()
	{
	document.getElementById("runmsg").innerHTML="Please wait while we are processing your request ...";
	document.body.style.cursor='progress';
	myConn.connect("xml_srcPrds.asp", "GET", GetAllValues(document.ajaxSearch), fnWhenDone);
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
	<input type="hidden" name="act" value="">
	<input type="hidden" name="src_IncNormal" value="<%=src_IncNormal%>">
	<input type="hidden" name="src_IncBTO" value="<%=src_IncBTO%>">
	<input type="hidden" name="src_IncItem" value="<%=src_IncItem%>">
	<input type="hidden" name="src_SM" value="<%=src_SM%>">
	<input type="hidden" name="src_IncDown" value="<%=src_IncDown%>">
	<input type="hidden" name="src_IncGC" value="<%=src_IncGC%>">
	<input type="hidden" name="src_Special" value="<%=src_Special%>">
	<input type="hidden" name="src_Featured" value="<%=src_Featured%>">
	<input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
	<input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
	<input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
	<input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
	<input type="hidden" name="src_Button2" value="<%=src_Button2%>">
	<input type="hidden" name="src_Button3" value="<%=src_Button3%>">
	<input type="hidden" name="src_DiscType" value="<%=src_DiscType%>">
    <input type="hidden" name="src_PromoType" value="<%=src_PromoType%>">
		
	<%'Start SDBA%>
	<input type="hidden" name="src_PageType" value="<%=src_PageType%>">
	<input type="hidden" name="src_IDSDS" value="<%=src_IDSDS%>">
	<input type="hidden" name="src_IsDropShipper" value="<%=src_IsDropShipper%>">
	<input type="hidden" name="src_sdsAssign" value="<%=src_sdsAssign%>">
	<input type="hidden" name="src_sdsStockAlarm" value="<%=src_sdsStockAlarm%>">	
	<%'End SDBA%>

	<input type="hidden" name="idcategory" value="<%=form_idcategory%>">
	<input type="hidden" name="customfield" value="<%=form_customfield%>">
	<input type="hidden" name="SearchValues" value="<%=form_SearchValues%>">
	<input type="hidden" name="priceFrom" value="<%=form_priceFrom%>">
	<input type="hidden" name="priceUntil" value="<%=form_priceUntil%>">
	<input type="hidden" name="withstock" value="<%=form_withstock%>">
    <input type="hidden" name="stocklevel" value="<%=form_stocklevel%>">
	<input type="hidden" name="sku" value="<%=form_sku%>">
	<input type="hidden" name="IDBrand" value="<%=form_IDBrand%>">
	<input type="hidden" name="keyWord" value="<%=form_keyWord%>">
	<input type="hidden" name="exact" value="<%=form_exact%>">
	<input type="hidden" name="pinactive" value="<%=form_pinactive%>">
	<input type="hidden" name="notforsale" value="<%=form_notforsale%>">
	<input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
	<input type="hidden" name="order" value="<%=form_order%>">
	<input type="hidden" name="iPageCurrent" value="1">
	<input type="hidden" name="prdlist" value="">
	</form>
	<div id="resultarea"></div>
	<script>
	function ExportPrds()
	{
		document.ajaxSearch.action="sds_exportprds.asp";
		document.ajaxSearch.target="_blank";
		document.ajaxSearch.submit();
	}
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
<%tmplist=""
if session("sm_selectall")="1" then
	tmpquery="SELECT Products.idProduct " & mid(query,Instr(query,"FROM"),len(query))
	set rs=connTemp.execute(tmpquery)
	if not rs.eof then
		tmpArr=rs.getRows()
		intCount=ubound(tmpArr,2)
		For k=0 to intCount
		tmplist=tmplist & tmpArr(0,k) & ","
		Next
	end if
	set rs=nothing
end if%>
var savelist="xml,<%=tmplist%>";

function getprdlist()
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
<script>
	runXML();
</script>
<%if (totalrecords>"0") AND (SRCH_MAX>"0") then%>
<p>
	Showing Top <b><%=totalrecords%></b> Results
</p>
<%end if%>
<% If iPageCount>1 Then %>
	<p class="pcPageNav">
	Currently viewing page <span id="currentpage">1</span> of <%=iPageCount%><br>
		<%For I = 1 To iPageCount%>
		<a href="javascript:document.ajaxSearch.iPageCurrent.value='<%=I%>';document.ajaxSearch.act.value='newsrc';runXML();"><%=I%></a> 
		<% Next %>
	</p>
<% End If %>

<% 
set rstemp=nothing
call closeDb()
%>
<form name="formreturn" method="post" class="pcForms">
<input type="hidden" name="prdlist" value="">
<p style="text-align: center;">
	<%Select Case session("sm_ShowNext")
	Case "1":%>
		<input type="button" name="Go" value="Continue" onClick="location='sm_addedit_S3.asp';" class="submit2">&nbsp;
	<%Case "2":%>
		<input type="button" name="Go" value=" Save Sale " onClick="location='sm_addedit_S5.asp';" class="submit2">&nbsp;
	<%Case "3":%>
		<input type="button" name="Go" value=" Back " onClick="location='sm_manage.asp';" class="submit2">&nbsp;
	<%Case "4":%>
		<input type="button" name="Go" value=" Back " onClick="location='sm_start.asp?a=start&id=<%=session("sm_pcSaleID")%>';" class="submit2">&nbsp;
		
	<%End Select%>
	<%if (src_ToPage<>"") and (pcv_HaveResults="1") and (src_DisplayType>"0") then%>
		<input type="button" value="<%=src_Button2%>" onClick="javascript:if (savelist=='xml,') {alert('Please select a product before clicking on this button');} else {if (confirm('Are you sure you want to complete this action?')) {document.formreturn.action='<%=src_ToPage%>';document.formreturn.prdlist.value=getprdlist(savelist);document.formreturn.submit();}}" class="submit2">
	<%end if%>
	<%if src_FromPage<>"" and ((src_IDSDS="") or (src_IDSDS="0")) then%>
	&nbsp;<input type="button" value="<%=src_Button3%>" onClick="location.href='<%=src_FromPage%>'">
	<%end if%>
</p>
</form>
<%if (src_IDSDS<>"") and (src_IDSDS<>"0") then%>
	<%if (src_sdsAssign="") or (src_sdsAssign="0") then%>
	<table class="pcCPcontent">
		<tr>
			<td align="center">
            <form class="pcForms">
            	<input type="button" onClick="javascript:AssignPrds();" value="Assign products to the <%if src_PageType="0" then%>supplier<%else%>drop-shipper<%end if%>">&nbsp;
                <input type="button" onClick="location.href='globalChanges.asp'" value="Move products">&nbsp;
                <input type="button" onClick="javascript:ExportPrds();" value="Export products">
             </form>
             </td>
		</tr>
	</table>
<%end if
end if%><!--#include file="Adminfooter.asp"-->