<!--#include file="../includes/utilities.asp"-->
<%
conntemp=""
call openDB()

'--- Generate Seach Parameters ---

if UseSpecial="1" then
else
	session("srcprd_from")=""
	session("srcprd_where")=""
	session("srcprd_DiscArea")=""
end if	

if src_FormTitle1="" then
	src_FormTitle1="Search a product"
end if

if src_FormTitle2="" then
	src_FormTitle2="Search Results"
end if

if src_IncNormal="" then
	src_IncNormal="0"
end if

if src_IncBTO="" then
	src_IncBTO="0"
end if

if src_IncItem="" then
	src_IncItem="0"
end if

if scBTO<>"1" then
	src_IncBTO="0"
	src_IncItem="0"
end if

if (src_IncBTO="0") and (src_IncItem="0") then
	src_IncNormal="1"
end if

if src_DisplayType="" then
	src_DisplayType="0"
end if

if src_ShowLinks="" then
	src_ShowLinks="0"
end if

if src_Special="" then
	src_Special="0"
end if

if src_Featured="" then
	src_Featured="0"
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

'Start SDBA
if src_PageType="" then
	src_PageType="0"
end if

if src_IDSDS="" then
	src_IDSDS="0"
end if

if src_IsDropShipper="" then
	src_IsDropShipper="0"
end if

if src_sdsAssign="" then
	src_sdsAssign="0"
end if

if src_sdsStockAlarm="" then
	src_sdsStockAlarm="0"
end if

'End SDBA

if src_ShowDiscTypes="" then
	src_ShowDiscTypes=0
end if

if src_ShowPromoTypes="" then
	src_ShowPromoTypes=0
end if

if src_DontShowInactive="1" then
	src_ShowInactive="0"
else
	src_ShowInactive="1"
end if

'--- End of Generate Seach Parameters ---

'Clear sessions

	session("cp_lct_src_FormTitle1")=""
	session("cp_lct_src_FormTitle2")=""
	session("cp_lct_src_FormTips1")=""
	session("cp_lct_src_FormTips2")=""
	session("cp_lct_src_IncNormal")=""
	session("cp_lct_src_IncBTO")=""
	session("cp_lct_src_IncItem")=""
	session("cp_lct_src_SM")=""
	session("cp_lct_src_IncDown")=""
	session("cp_lct_src_IncGC")=""
	session("cp_lct_src_Special")=""
	session("cp_lct_src_Featured")=""
	session("cp_lct_src_DisplayType")=""
	session("cp_lct_src_ShowLinks")=""
	session("cp_lct_src_FromPage")=""
	session("cp_lct_src_ToPage")=""
	session("cp_lct_src_Button2")=""
	session("cp_lct_src_Button3")=""
	session("cp_lct_src_DiscType")=""
	session("cp_lct_src_PromoType")=""
	
	'Start SDBA
	session("cp_lct_src_PageType")=""
	session("cp_lct_src_IDSDS")=""
	session("cp_lct_src_IsDropShipper")=""
	session("cp_lct_src_sdsAssign")=""
	session("cp_lct_src_sdsStockAlarm")=""
	'End SDBA

	session("cp_lct_form_idcategory")=""
	session("cp_lct_form_customfield")=""
	session("cp_lct_form_SearchValues")=""
	session("cp_lct_form_priceFrom")=""
	session("cp_lct_form_priceUntil")=""
	session("cp_lct_form_withstock")=""
	session("cp_lct_form_stocklevel")=""
	session("cp_lct_form_sku")=""
	session("cp_lct_form_IDBrand")=""
	session("cp_lct_form_keyWord")=""
	session("cp_lct_form_exact")=""
	session("cp_lct_form_pinactive")=""
	session("cp_lct_form_notforsale")=""
	session("cp_lct_form_resultCnt")=""
	session("cp_lct_form_order")=""

'AJAX Functions 
%>
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
myConn.connect("xml_srcPrdsCount.asp", "POST", GetAllValues(document.ajaxSearch), fnWhenDone);
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

<%
'End of AJAX Functions

'Validate Form
%>
<script language="JavaScript">
<!--
	
function isDigitA(s)
{
var test=""+s;
if(test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigitA(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigitA(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

function FormValidator(theForm)
{
 	qtt= theForm.priceFrom;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	    
	qtt= theForm.priceUntil;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{
		    alert("Please enter a numeric value for this field.");
		    qtt.focus();
		    return (false);
		    }
	    }
	
return (true);
}
//-->
</script>

<%' Begin search form %>

<form name="ajaxSearch" method="post" action="srcPrds.asp?action=newsrc" onSubmit="return FormValidator(this)" class="pcForms">

<input type="hidden" name="referpage" value="NewSearch">
<input type="hidden" name="src_FormTitle1" value="<%=src_FormTitle1%>">
<input type="hidden" name="src_FormTitle2" value="<%=src_FormTitle2%>">
<input type="hidden" name="src_FormTips1" value="<%=src_FormTips1%>">
<input type="hidden" name="src_FormTips2" value="<%=src_FormTips2%>">
<%if src_ShowPrdTypeBtns<>"1" then%>
<input type="hidden" name="src_IncNormal" value="<%=src_IncNormal%>">
<input type="hidden" name="src_IncBTO" value="<%=src_IncBTO%>">
<input type="hidden" name="src_IncItem" value="<%=src_IncItem%>">
<%end if%>
<%if src_ShowDiscTypes="0" then%>
<input type="hidden" name="src_DiscType" value="">
<%end if%>
<input type="hidden" name="src_Special" value="<%=src_Special%>">
<input type="hidden" name="src_Featured" value="<%=src_Featured%>">
<input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
<input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
<input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
<input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
<input type="hidden" name="src_Button2" value="<%=src_Button2%>">
<input type="hidden" name="src_Button3" value="<%=src_Button3%>">

<%'Start SDBA%>
<input type="hidden" name="src_PageType" value="<%=src_PageType%>">
<input type="hidden" name="src_IDSDS" value="<%=src_IDSDS%>">
<input type="hidden" name="src_IsDropShipper" value="<%=src_IsDropShipper%>">
<input type="hidden" name="src_sdsAssign" value="<%=src_sdsAssign%>">
<input type="hidden" name="src_sdsStockAlarm" value="<%=src_sdsStockAlarm%>">
<%'End SDBA%>

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
		<td colspan="2"><p><span id="totalresults">&nbsp;</span></p></td>
	</tr>

	<% ' Get category list %>
	<tr> 
		<td width="25%"><p>Category:</p></td>
		<td width="75%">
			<%if cat_HadItem<>"" then 'Had a Category ID%>
				<%=cat_HadItem%>
			<%else
				cat_DropDownName="idcategory"
				cat_Type="1"
				cat_DropDownSize="1"
				cat_MultiSelect="0"
				cat_ExcBTOHide="0"
				cat_StoreFront="0"
				cat_ShowParent="1"
				cat_DefaultItem="All"
				cat_SelectedItems="0,"
				cat_ExcItems=""
				cat_ExcSubs="0"
				cat_EventAction="onchange='javascript:runXML();'"
				%>
				<!--#include file="../includes/pcCategoriesList.asp"-->
				<%call pcs_CatList()%>
			<%end if%>
	</td>
</tr>

<% ' search custom fields if any are defined %>
<%tmpJSStr=""
tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
intCount=-1
tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>

				<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldCPSearch=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				if not rs.eof then
					set pcv_tempFunc = new StringBuilder
					pcv_tempFunc.append "<script>" & vbcrlf
					pcv_tempFunc.append "function CheckList(cvalue,tmpvalue) {" & vbcrlf
					pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
					pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
					pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
					pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0"");" & vbcrlf
					pcv_tempFunc.append "SFID=new Array();" & vbcrlf
					pcv_tempFunc.append "SFNAME=new Array();" & vbcrlf
					pcv_tempFunc.append "SFVID=new Array();" & vbcrlf
					pcv_tempFunc.append "SFVALUE=new Array();" & vbcrlf
					pcv_tempFunc.append "SFVORDER=new Array();" & vbcrlf
					intCount=-1
					pcv_tempFunc.append "SFCount=" & intCount & ";" & vbcrlf
					pcv_tempFunc.append "CreateTable(tmpvalue);" & vbcrlf
					pcv_tempFunc.append "}" & vbcrlf
					
					set pcv_tempList = new StringBuilder
					pcv_tempList.append "<select name=""customfield1"" onchange=""javascript:CheckList(document.ajaxSearch.customfield1.value,0);"">" & vbcrlf
					pcv_tempList.append "<option value=""0"">All</option>" & vbcrlf
					pcArray=rs.getRows()
					intCount=ubound(pcArray,2)
					set rs=nothing
					
					For i=0 to intCount
						pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
						query="SELECT idSearchData,pcSearchDataName FROM pcSearchData WHERE idSearchField=" & pcArray(0,i) & " ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpArr=rs.getRows()
							LCount=ubound(tmpArr,2)
							pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
							pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
							pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
							pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0"");" & vbcrlf
							For j=0 to LCount
								pcv_tempFunc.append "SelectA.options[" & j+1 & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
							Next
							pcv_tempFunc.append "}" & vbcrlf
						else
							pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
							pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
							pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
							pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0""); }" & vbcrlf
						end if
					Next
			
					pcv_tempList.append "</select>" & vbcrlf
					pcv_tempFunc.append "}" & vbcrlf
					pcv_tempFunc.append "</script>" & vbcrlf
					
					pcv_tempList=pcv_tempList.toString
					pcv_tempFunc=pcv_tempFunc.toString
					%>
					<tr>
						<td nowrap><p>Filter by:</p></td>
						<td>
							<%=pcv_tempList%>&nbsp;
							<select name="SearchValues1" onchange="javascript:var testvalue=document.ajaxSearch.customfield.value; if ((testvalue.indexOf('||')==-1) && (document.ajaxSearch.customfield1.value!='0')) {document.ajaxSearch.customfield.value=document.ajaxSearch.customfield1.value;document.ajaxSearch.SearchValues.value=this.value;runXML()}">
							</select>
							<%=pcv_tempFunc%>
							&nbsp;<a href="javascript:AddSF(document.ajaxSearch.customfield1.value,document.ajaxSearch.customfield1.options[document.ajaxSearch.customfield1.selectedIndex].text,document.ajaxSearch.SearchValues1.value,document.ajaxSearch.SearchValues1.options[document.ajaxSearch.SearchValues1.selectedIndex].text,0);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
						</td>
					</tr>
                    
					<tr>
					<td nowrap></td>
					<td>
						<input type="hidden" name="customfield" value="0">
						<input type="hidden" name="SearchValues" value="">
						<span id="stable" name="stable"></span>
						<script>
						<%=tmpJSStr%>
						function CreateTable(tmpRun)
						{
							var tmp1="";
							var tmp2="";
							var tmp3="";
							var i=0;
							var found=0;
							tmp1='<br><table class="pcCPcontent">';
							for (var i=0;i<=SFCount;i++)
							{
								found=1;
								tmp1=tmp1 + '<tr><td align="right"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="Remove" border="0"></a></td><td width="100%" nowrap>'+SFNAME[i]+': '+SFVALUE[i]+'</td></tr>';
								if (tmp2=="") tmp2=tmp2 + "||";
								tmp2=tmp2 + SFID[i] + "||";
								if (tmp3=="") tmp3=tmp3 + "||";
								tmp3=tmp3 + SFVID[i] + "||";
							}
							tmp1=tmp1+'</table><br><br>';
							if (found==0) tmp1="";
							document.getElementById("stable").innerHTML=tmp1;
							if (tmp2=="") tmp2=0;
							document.ajaxSearch.customfield.value=tmp2;
							document.ajaxSearch.SearchValues.value=tmp3;
							if (tmp2==0)
							{
								document.ajaxSearch.customfield.value=document.ajaxSearch.customfield1.value;
								document.ajaxSearch.SearchValues.value=document.ajaxSearch.SearchValues1.value;
							}
							if (tmpRun!=1) runXML();
						}
						
						CheckList(document.ajaxSearch.customfield1.value,1);
						
						function ClearSF(tmpSFID)
						{
							var i=0;
							for (var i=0;i<=SFCount;i++)
							{
								if (SFID[i]==tmpSFID)
								{
									removedArr = SFID.splice(i,1);
									removedArr = SFNAME.splice(i,1);
									removedArr = SFVID.splice(i,1);
									removedArr = SFVALUE.splice(i,1);
									removedArr = SFVORDER.splice(i,1);
									SFCount--;
									break;
								}
							}
							CreateTable(0);
						}
						
						function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
						{
							if ((tmpSVID!="") && (tmpSFID!="") && (tmpSVID!="0") && (tmpSFID!="0"))
							{
								var i=0;
								var found=0;
								for (var i=0;i<=SFCount;i++)
								{
									if (SFID[i]==tmpSFID)
									{
										SFVID[i]=tmpSVID;
										SFVALUE[i]=tmpSValue;
										SFVORDER[i]=tmpSOrder;
										found=1;
										break;
									}
								}
								if (found==0)
								{
									SFCount++;
									SFID[SFCount]=tmpSFID;
									SFNAME[SFCount]=tmpSFName;
									SFVID[SFCount]=tmpSVID;
									SFVALUE[SFCount]=tmpSValue;
									SFVORDER[SFCount]=tmpSOrder;
								}
								CreateTable(0);
							}
						}
						</script>
					</td>
					</tr>
				<%pcv_HaveSearchFields=1
				else%>
					<input type="hidden" name="customfield" value="0">
				<%end if %>
<% 
'End of custom fields

'Product Prices
%>
<tr> 
	<td><p>Price:</p></td>
	<td>From:&nbsp;<input type="text" name="priceFrom" size="6" value="0" onblur="javascript:runXML();">&nbsp;To:&nbsp;<input type="text" name="priceUntil" size="6" value="999999999" onblur="javascript:runXML();">
	</td>
</tr>

<% ' Inventory %>
<%if src_StockChoices=1 then%>
<tr> 
	<td><p>Product Inventory Level:</p></td>
	<td>  
	<input type="radio" name="withstock" value="-1" onclick="javascript:runXML();" class="clearBorder">&nbsp;In Stock&nbsp;&nbsp;
    <input type="radio" name="withstock" value="2" onclick="javascript:runXML();" class="clearBorder">&nbsp;Out of Stock&nbsp;&nbsp;
    <input type="radio" name="withstock" value="0" onclick="javascript:runXML();" class="clearBorder" checked>&nbsp;Both
    
	</td>
</tr>
<%else%>       
<tr> 
	<td><p>In stock:</p></td>
	<td>  
	<input type="checkbox" name="withstock" value="-1" onclick="javascript:runXML();" class="clearBorder">
	</td>
</tr>
<%end if%>
<%if src_ShowStockLevel=1 then%>
<tr> 
	<td><p>Only show items whose inventory level is lower than:</p></td>
	<td>  
	<input type="text" name="stocklevel" size="4" value="" onblur="javascript:runXML();">
	</td>
</tr>
<%end if%>

<%if src_checkPrdType<>"2" then%>
<tr> 
	<td><p>Not For Sale status:</p></td>
	<td>  
	<input type="radio" name="notforsale" value="-1" onclick="javascript:runXML();" class="clearBorder">&nbsp;For sale&nbsp;&nbsp;
    <input type="radio" name="notforsale" value="2" onclick="javascript:runXML();" class="clearBorder">&nbsp;Not for sale&nbsp;&nbsp;
    <input type="radio" name="notforsale" value="0" onclick="javascript:runXML();" class="clearBorder" checked>&nbsp;Both
	</td>
</tr>
<% end if %>

<% 'Product SKU %>
          
<tr> 
	<td><p>Part number (SKU):</p></td>
	<td> 
	<input name="sku" type="text" size="15" maxlength="150" onblur="javascript:runXML();"></td>
</tr>

<% 'Get Brands %>
          
<%query="Select IDBrand,BrandName from Brands order by BrandName asc"
set rsTempFP=connTemp.execute(query)
if not rsTempFP.eof then
pcArr=rsTempFP.getRows()
intCount1=ubound(pcArr,2)%>
<tr> 
	<td><p>Brand:</p></td>
	<td>
	<select name="IDBrand" onchange="javascript:runXML();">
	<option value="0" selected>All</option>
	<%For m=0 to intCount1%>
	<option value="<%=pcArr(0,m)%>"><%=pcArr(1,m)%></option>
	<%Next%>
	</select></td>
</tr>
<%end if
set rsTempFP=nothing%>

<% 'Search Keywords %>
          
<tr> 
	<td valign="top"><p>Keyword(s):</p></td>
	<td> 
	<input type="text" name="keyWord" size="40" onblur="javascript:runXML();">
	<br>
	<input type="checkbox" name="exact" value="1" onclick="javascript:runXML();" class="clearBorder"> Search on exact phrase
	</td>
</tr>

<%
	if src_ShowPrdTypeBtns="1" then
		if (scBTO=1) OR (src_SM=1) then
%>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<tr>
		<td valign="top"><p>Filter by Product Type:</p></td>
		<td valign="top">Standard Products	<input type="checkbox" name="src_IncNormal" value="1" <%if src_checkPrdType="0" then%>checked<%end if%> onclick="javascript:runXML();" class="clearBorder">
			<%if src_SM<>"1" then%>
			&nbsp;&nbsp;&nbsp;&nbsp;BTO Products <input type="checkbox" name="src_IncBTO" value="1" <%if src_checkPrdType="1" then%>checked<%end if%> onclick="javascript:runXML();" class="clearBorder">&nbsp;&nbsp;&nbsp;&nbsp;BTO Items	<input type="checkbox" name="src_IncItem" value="1" <%if src_checkPrdType="2" then%>checked<%end if%> onclick="javascript:runXML();" class="clearBorder">
			<%end if%>
			<%if src_SM=1 then%>
			&nbsp;&nbsp;&nbsp;&nbsp;Downloadable Products	<input type="checkbox" name="src_IncDown" value="1" <%if src_checkPrdType="3" then%>checked<%end if%> onclick="javascript:runXML();" class="clearBorder">
			&nbsp;&nbsp;&nbsp;&nbsp;Gift Certificates 	<input type="checkbox" name="src_IncGC" value="1" <%if src_checkPrdType="4" then%>checked<%end if%> onclick="javascript:runXML();" class="clearBorder">
			<input type="hidden" name="src_SM" value="1" />
			<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
<%
		else
%>
		<input type="hidden" name="src_IncNormal" value="1">
		<input type="hidden" name="src_IncBTO" value="0">
		<input type="hidden" name="src_IncItem" value="0">
<%
		end if
	end if
%>

<%
	if src_ShowDiscTypes="1" then
%>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<tr>
		<td valign="top"><p>Filter by Discount:</p></td>
		<td valign="top">
        <div style="float: right; width: 230px; font-size: 11px; background-color: #FF9; padding: 6px; margin-right: 10px"><strong>NOTE</strong>: Promotions and Quantity Discounts cannot be applied at the same time. So products for which a promotion has been set will not appear in the search results.</div>
        <input type="radio" name="src_DiscType" value="1" onclick="javascript:runXML();" class="clearBorder"> Only products with quantity discounts<br>
		<input type="radio" name="src_DiscType" value="2" onclick="javascript:runXML();" class="clearBorder"> Only products without quantity discounts<br>
		<input type="radio" name="src_DiscType" value="0" onclick="javascript:runXML();" checked class="clearBorder"> Include both
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
<%
	end if
%>

<%
	if src_ShowPromoTypes="1" then
%>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
	<tr>
		<td valign="top"><p>Filter by Promotion:</p></td>
		<td valign="top">
        <div style="float: right; width: 230px; font-size: 11px; background-color: #FF9; padding: 6px; margin-right: 10px"><strong>NOTE</strong>: Promotions and Quantity Discounts cannot be applied at the same time. So products for which quantity discounts have been set will not appear in the search results.</div>
        <input type="radio" name="src_PromoType" value="1" onclick="javascript:runXML();" class="clearBorder"> Only products with a promotion<br>
		<input type="radio" name="src_PromoType" value="2" onclick="javascript:runXML();" class="clearBorder"> Only products without a promotion<br>
		<input type="radio" name="src_PromoType" value="0" onclick="javascript:runXML();" checked class="clearBorder"> Include both
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
<%
	end if
%>

<%
	if src_ShowInactive="1" then
%>
<tr> 
	<td><p>Include inactive products:</p></td>
	<td>
	<input type="checkbox" name="pinactive" value="-1" checked onclick="javascript:runXML();" class="clearBorder">
	</td>
</tr>
<%else%>
	<input type="hidden" name="pinactive" value="">
<%end if%>

<tr> 
	<td><p>Results per page:</p></td>
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
	<td nowrap><p>Sort by:</p></td>
	<td>
	<select name="order">
		<option value="2">SKU ascending</option>
		<option value="3">SKU descending</option>
		<option value="4" selected>Description ascending</option>
		<option value="5">Description descending</option>
	</select>
	<td>
</tr>
<tr>
    <td colspan="2" align="center"><hr /></td>
</tr>
<tr> 
	<td colspan="2">
	<input type="hidden" name="act" value="newsrc">
	<input name="runform" type="submit" value="<%=src_Button1%>" id="searchSubmit">
	</td>
</tr>  
</table>
</form>
<%if pcv_HaveSearchFields=1 then%>
<script>CreateTable(1);</script>
<%end if
' End of Search Form
call closeDb()
%>