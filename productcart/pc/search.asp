<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="../includes/utilities.asp"-->
<%
on error resume next

'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "search.asp"

%>
<!--#include file="pcStartSession.asp"-->
<%

Dim query, conntemp, rs, rstemp
call openDb()

%>

<!--#include file="header.asp"-->
<!--#include file="pcValidateHeader.asp"-->

<%
'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>

<%IF scStoreUseToolTip="1" or scStoreUseToolTip="4" THEN%>
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
	function runXML()
	{
	}
</script>
<%session("store_useAjax")="1"
ELSE%>
<%session("store_useAjax")=""%>
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
<!--End of AJAX Functions-->
<%END IF%>
<%END IF 'End of AJAX turns on%>

<!--Validate Form-->
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
 	qtt= document.ajaxSearch.priceFrom;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{		    
		    alert("<%response.write dictLanguage.Item(Session("language")&"_advSrca_26")%>");
		    qtt.focus();
			<% If SRCH_WAITBOX="1" Then %>
				CloseHS();
			<% End If %>
		    return (false);
		    }
	    }
	    
	qtt= document.ajaxSearch.priceUntil;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{
		    alert("<%response.write dictLanguage.Item(Session("language")&"_advSrca_26")%>");
		    qtt.focus();
			<% If SRCH_WAITBOX="1" Then %>
				CloseHS();
			<% End If %>
		    return (false);
		    }
	    }
		return (true);
}
<% If SRCH_WAITBOX="1" Then %>
function CloseHS() 
{
	var t=setTimeout("hs.close('pcMainSearch')",50)
}
function OpenHS() 
{
	document.getElementById('pcMainSearch').onclick()
}
<% End If %>
//-->
</script>

<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td colspan="2"> 
				<h1><%response.write dictLanguage.Item(Session("language")&"_advSrca_1")%></h1>
				<% ' Show search page description, if any
				pcStrSearchDesc = dictLanguage.Item(Session("language")&"_search_1")
				if trim(pcStrSearchDesc) <> "" then %>
				<div class="pcPageDesc"><%=pcStrSearchDesc%></div>
		<% 	end if %>
			</td>
		</tr>
		<tr>
			<td>
            <%
			'// Set Submit Action
			Dim pcv_strSubmitAction
			If SRCH_WAITBOX="1" Then
				pcv_strSubmitAction = "OpenHS(); return FormValidator(this);"
			Else
				pcv_strSubmitAction = "return FormValidator(this);"
			End If
			%>
			<form name="ajaxSearch" method="get" class="pcForms" action="showsearchresults.asp" onSubmit="<%=pcv_strSubmitAction%>">
				<table class="pcShowContent">
				<tr>
				<td colspan="2">
					<p><span id="totalresults" class="pcTextMessage">&nbsp;</span></p>
				</td>
				</tr>
					
					<%
					'// CATEGORY DROP DOWN - START
							select case schideCategory
							 case "0" ' // FIRST scenario: Category drop-down fully shown
					%>
								<tr> 
									<td width="25%" nowrap>
										<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_2")%></p>
									</td>
									<td width="75%">
										<%
										cat_DropDownName="idcategory"
										cat_Type="1"
										cat_DropDownSize="1"
										cat_MultiSelect="0"
										cat_ExcBTOHide="1"
										cat_StoreFront="1"
										cat_ShowParent="1"
										cat_DefaultItem=dictLanguage.Item(Session("language")&"_advSrca_4")
										cat_SelectedItems="0,"
										cat_ExcItems=""
										cat_ExcSubs="0"
										if scStoreUseToolTip="1" or scStoreUseToolTip="4" then
											cat_EventAction="onchange='javascript:runXML();'"
										else
											cat_EventAction=""
										end if
										%>
										<!--#include file="../includes/pcCategoriesList.asp"-->
										<%call pcs_CatList()%>
									</td>
							</tr>
						<%
							case "1" '// Only top-level categories are shown
								
								query="SELECT DISTINCT categories.idCategory,categories.categoryDesc,categories.idParentCategory "
								query=query&"FROM categories "
								query=query&"WHERE categories.iBTOhide=0 AND categories.pccats_RetailHide=0 AND idParentCategory=1 AND idCategory<>1 "
								query=query&"ORDER BY categories.categoryDesc ASC;"

								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=connTemp.execute(query)
								if not rs.eof then
									Dim categoryArray, categoryCount, categoryTotal
									categoryArray = rs.getRows()
									categoryCount = 0
									categoryTotal = ubound(categoryArray,2)
						%>
									<tr> 
										<td width="25%" nowrap>
											<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_2")%></p>
										</td>
										<td width="75%">
                                            <%
											if scStoreUseToolTip="1" or scStoreUseToolTip="4" then
                                                cat_EventAction="onchange='javascript:runXML();'"
                                            else
                                                cat_EventAction=""
                                            end if
											%>
											<select name="idcategory" <%=cat_EventAction%>>
												<option value="0" selected><%=dictLanguage.Item(Session("language")&"_advSrca_4")%></option>
													<%
													do while (categoryCount <= categoryTotal)
													%>
														<option value="<%=categoryArray(0, categoryCount)%>"><%=categoryArray(1, categoryCount)%></option>
													<%
													 categoryCount = categoryCount + 1
													loop
													%>
											</select>
										</td>
									 </tr>
							<%		
								end if
								set rs = nothing	
													
							case "-1" '// The category drop-down is hidden
							%>
							<tr>
							<td colspan="2" style="height:0px;">
								<input type="hidden" name="idcategory" value="0">
							</td>
							</tr>
							<%
							End Select
							'// CATEGORY DROP DOWN - END
							%>
							
							<tr>
							<td nowrap>
									<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_5")%></p>
							</td>
							<td> 
									<%response.write dictLanguage.Item(Session("language")&"_advSrca_6")%>
									<input type="text" name="priceFrom" value="0" size="6" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onblur="javascript:if (FormValidator(document.getElementById('ajaxSearch'))) {runXML();} else {return(false);}"<%end if%>>
									<%response.write dictLanguage.Item(Session("language")&"_advSrca_7")%>
									<input type=text name="priceUntil" value="999999999" size="10" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onblur="javascript:if (FormValidator(document.getElementById('ajaxSearch'))) {runXML();} else {return(false);}"<%end if%>>
							 </td>
							</tr>
							<tr> 
								<td nowrap> 
									<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_8")%></p>
								</td>
								<td>
									<input type="checkbox" name="withstock" value="-1" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onclick="javascript:runXML();"<%end if%> class="clearBorder">
								</td>
							</tr>
							<tr> 
								<td nowrap> 
									<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_11")%></p>
								</td>
								<td>
									<input name="sku" type="text" size="15" maxlength="150" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onblur="javascript:runXML();"<%end if%>>
								</td>
							</tr>
							
							<%
							'Show brands, if any
							query="Select IDBrand, BrandName from Brands order by BrandName asc"
							set rs=Server.CreateObject("ADODB.Recordset")
							set rs=connTemp.execute(query)
							if not rs.eof then
								Dim brandArray, brandCount, brandTotal
								brandArray = rs.getRows()
								brandCount = 0
								brandTotal = ubound(brandArray,2)
								%>
									<tr> 
										<td nowrap>
											<p><%=dictLanguage.Item(Session("language")&"_advSrca_13")%></p>
										</td>
										<td>
											<select name="IDBrand" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onchange="javascript:runXML();"<%end if%>>
												<option value="0" selected><%=dictLanguage.Item(Session("language")&"_advSrca_4")%></option>
													<%
													do while (brandCount <= brandTotal)
													%>
														<option value="<%=brandArray(0, brandCount)%>"><%=brandArray(1, brandCount)%></option>
													<%
													 brandCount = brandCount + 1
													 loop
													%>
											</select>
										</td>
									 </tr>
							 <%
							 end if
							 set rs = nothing
							 %>
							 
							<tr>
								<td nowrap>
									<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_9")%></p>
								</td>
								<td>
									<input type="text" name="keyWord" size="30" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onblur="javascript:runXML();"<%end if%>>
									<input type="checkbox" name="exact" value="1" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onclick="javascript:runXML();"<%end if%> class="clearBorder"><%response.write dictLanguage.Item(Session("language")&"_advSrca_14")%>
								</td>
							</tr>
							
				<!-- search custom fields if any are defined -->
				<%tmpJSStr=""
				tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
				tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
				tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
				tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
				tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
				intCount=-1
				tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>

				<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldSearch=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
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
						<td colspan="2">
							<hr>
						</td>
					</tr>								
					<tr>
						<td colspan="2">
							<p><%response.write dictLanguage.Item(Session("language")&"_advSrca_22")%></p>
						</td>
					</tr>
					<tr>
						<td nowrap><p><%response.write dictLanguage.Item(Session("language")&"_advSrca_12")%></p></td>
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
							tmp1='<table style="width: 100%; margin: 5px 0 5px 0;">';
							for (var i=0;i<=SFCount;i++)
							{
								found=1;
								tmp1=tmp1 + '<tr><td style="text-align: right;"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="" border="0"></a></td><td style="text-align: left; width: 100%;">'+SFNAME[i]+': '+SFVALUE[i]+'</td></tr>';
								if (tmp2=="") tmp2=tmp2 + "||";
								tmp2=tmp2 + SFID[i] + "||";
								if (tmp3=="") tmp3=tmp3 + "||";
								tmp3=tmp3 + SFVID[i] + "||";
							}
							tmp1=tmp1+'</table>';
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
				<!-- end of custom fields -->			
                    
							<%
							'SM-Start
							if UCase(scDB)="SQL" then
							tmpTargetType=0
							if session("customerCategory")<>"" AND session("customerCategory")<>"0" then
								tmpTargetType=session("customerCategory")
							else
								if session("customerType")="1" then
									tmpTargetType="-1"
								end if
							end if
							
							query="SELECT pcSales_Completed.pcSC_ID ,pcSales_Completed.pcSC_SaveName FROM pcSales_Completed INNER JOIN pcSales ON pcSales_Completed.pcSales_ID=pcSales.pcSales_ID WHERE pcSales_Completed.pcSC_Status=2 AND pcSales.pcSales_TargetPrice=" & tmpTargetType & " AND pcSales_Completed.pcSC_Archived=0 ORDER BY pcSC_SaveName ASC;"
							set rs=Server.CreateObject("ADODB.Recordset")
							set rs=connTemp.execute(query)
							if not rs.eof then
								saleArr=rs.getRows()
								intSale=ubound(saleArr,2)
								%>		
                                    <tr>
                                        <td colspan="2">
                                            <hr>
                                        </td>
                                    </tr>						
									<tr> 
										<td nowrap>
											<p><%=dictLanguage.Item(Session("language")&"_SaleSearch_1")%></p>
										</td>
										<td>
											<input type="checkbox" name="incSale" value="1" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onclick="javascript:runXML();"<%end if%> class="clearBorder">
											&nbsp;
											<%=dictLanguage.Item(Session("language")&"_SaleSearch_2")%>
											<select name="IDSale" <%if scStoreUseToolTip="1" or scStoreUseToolTip="4" then%>onchange="javascript:runXML();"<%end if%>>
												<option value="0" selected><%=dictLanguage.Item(Session("language")&"_SaleSearch_3")%></option>
													<%
													For k=0 to intSale
													%>
														<option value="<%=saleArr(0,k)%>"><%=saleArr(1,k)%></option>
													<%
													Next
													%>
											</select>
										</td>
									 </tr>
                                    <tr>
                                        <td colspan="2">
                                            <hr>
                                        </td>
                                    </tr>
							 <%
							 end if
							 set rs = nothing
							 end if
							 'SM-End
							 %>
                             
							<% 
                              '// Locate preferred results count and load as default
                                Dim pcIntPreferredCount
                                pcIntPreferredCount =(scPrdRow*scPrdRowsPerPage)
                                            if validNum(pcIntPreferredCount) then
                              %>				
                                    <tr> 
                                        <td nowrap>
                                            <p><%=dictLanguage.Item(Session("language")&"_advSrca_15")%></p>
                                        </td>
                                        <td>
                                            <select name="resultCnt" id="resultCnt">
                                                <option value="<%=pcIntPreferredCount%>" selected><%=pcIntPreferredCount%></option>
                                                <option value="<%=pcIntPreferredCount*2%>"><%=pcIntPreferredCount*2%></option>
                                                <option value="<%=pcIntPreferredCount*3%>"><%=pcIntPreferredCount*3%></option>
                                                <option value="<%=pcIntPreferredCount*4%>"><%=pcIntPreferredCount*4%></option>
                                                <option value="<%=pcIntPreferredCount*5%>"><%=pcIntPreferredCount*5%></option>
                                                <option value="<%=pcIntPreferredCount*10%>"><%=pcIntPreferredCount*10%></option>
                                            </select>
                                        </td>
                                    </tr>
                            <%
                             end if
                            %>
          
							<tr> 
								<td nowrap> 
									<p><%=dictLanguage.Item(Session("language")&"_advSrca_16")%></p>
								</td>
								<td>
									<select name="order">
										<option value="0"<% if PCOrd=0 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_18")%></option>
										<option value="1"<% if PCOrd=1 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_19")%></option>
										<option value="3"<% if PCOrd=3 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_20")%></option>
										<option value="2"<% if PCOrd=2 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_21")%></option>
									</select>
								</td>
							</tr>
                            <tr>
                                <td colspan="2">
                                    <hr>
                                </td>
                            </tr>
							<tr> 
								<td colspan="2" nowrap>  
									<input type="image" src="<%=rslayout("submit")%>" name="Submit" id="submit" value="<%response.write dictLanguage.Item(Session("language")&"_advSrca_10")%> ">
								 </td>
							</tr>
						</table>
					</form>
					<%
					If SRCH_WAITBOX="1" Then
						'// Loading Window
						'	>> Call Method with OpenHS();
						response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_advSrca_23"), "pcMainSearch", 200))
					End If
					%>
			</td>
		</tr>
	</table>
</div>
<%if pcv_HaveSearchFields=1 then%>
<script>CreateTable(1);</script>
<%end if%>
<% call closeDb()
set rstemp= nothing %><!--#include file="footer.asp"-->