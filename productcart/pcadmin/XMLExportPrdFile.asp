<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Product XML File" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/productcartfolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%Dim connTemp,query,rs
call openDB()%>

<script language="javascript">
function CalPop(sInputName)
{
	window.open('../Calendar/Calendar.asp?N=' + escape(sInputName) + '&DT=' + escape(window.eval(sInputName).value), 'CalPop','toolbar=0,width=378,height=225' );
}
</script>

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

<form name="ajaxSearch" method="post" action="XMLExportPrdFileA.asp?action=newsrc" onSubmit="return FormValidator(this)" class="pcForms">

<table class="pcCPsearch">
	<tr>
		<td colspan="2">
			<% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
			<h2>Find Products</h2>
			<p>Use the following filters to look for products to be included in the XML file.<br>The file will be saved in the folder: <strong>&quot;<%=scPcFolder%>/xml/export&quot; </strong>on your web server.</p>
		</td>
	</tr>
	<!--Get Category List-->
	<tr> 
		<td width="25%">Category:</td>
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
				cat_EventAction=""
				%>
				<!--#include file="../includes/pcCategoriesList.asp"-->
				<%call pcs_CatList()%>
			<%end if%>
	</td>
</tr>

<!-- search custom fields if any are defined -->
<% query="SELECT customfields.idcustom, customfields.custom, customfields.searchable FROM customfields ORDER BY customfields.custom;"
set rsTempFP=conntemp.execute(query)
if not rsTempFP.eof then
	pcv_tempFunc=""
	pcv_tempFunc=pcv_tempFunc & "<script>" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "function CheckList(cvalue) {" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "if (cvalue==0) {" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "SelectA.options.length = 0;" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "SelectA.options[0]=new Option(""All"",""""); }" & vbcrlf
					
	pcv_tempList=""
	pcv_tempList=pcv_tempList & "<select name=""customfield"" onchange=""javascript:CheckList(document.ajaxSearch.customfield.value);"">" & vbcrlf
	pcv_tempList=pcv_tempList & "<option value=""0"">All</option>" & vbcrlf
				
	pcArray=rsTempFP.getRows()
	intCount=ubound(pcArray,2)
	
	Dim AList(200)
	
	For i=0 to intCount
		pcv_checkvalue=0
		LCount=0
					
		if scDB="SQL" then
			query="SELECT content1 AS VContent1 FROM products WHERE content1 IS NOT NULL and custom1=" & pcArray(0,i) & " and active=-1 and removed=0;"
		else
			query="SELECT content1 AS VContent1 FROM products WHERE content1<>'' and custom1=" & pcArray(0,i) & " and active=-1 and removed=0 group by content1;"
		end if
		set rsTempFP=connTemp.execute(query)
					
		if not rsTempFP.eof then
			pcv_checkvalue=1
			pcArray1=rsTempFP.getRows()
			intCount1=ubound(pcArray1,2)
						
			For j=0 to intCount1
				m1=0
				For k=1 to LCount
					if pcArray1(0,j)=AList(k) then
						m1=1
						exit for
					end if
				Next
						
				if m1=0 then
					LCount=LCount+1
					AList(LCount)=pcArray1(0,j)
				end if
			Next
		end if
					
		if scDB="SQL" then
			query="SELECT content2 AS VContent2 FROM products WHERE content2 IS NOT NULL and custom2=" & pcArray(0,i) & " and active=-1 and removed=0;"
		else
			query="SELECT content2 AS VContent2 FROM products WHERE content2<>'' and custom2=" & pcArray(0,i) & " and active=-1 and removed=0 group by content2;"
		end if
		set rsTempFP=connTemp.execute(query)
					
		if not rsTempFP.eof then
			pcv_checkvalue=1
			pcArray1=rsTempFP.getRows()
			intCount1=ubound(pcArray1,2)
						
			For j=0 to intCount1
				m1=0
				For k=1 to LCount
					if pcArray1(0,j)=AList(k) then
						m1=1
						exit for
					end if
				Next
						
				if m1=0 then
					LCount=LCount+1
					AList(LCount)=pcArray1(0,j)
				end if
			Next
		end if
					
		if scDB="SQL" then
			query="SELECT content3 AS VContent3 FROM products WHERE content3 IS NOT NULL and custom3=" & pcArray(0,i) & " and active=-1 and removed=0;"
		else
			query="SELECT content3 AS VContent3 FROM products WHERE content3<>'' and custom3=" & pcArray(0,i) & " and active=-1 and removed=0 group by content3;"
		end if
		set rsTempFP=connTemp.execute(query)
					
		if not rsTempFP.eof then
			pcv_checkvalue=1
			pcArray1=rsTempFP.getRows()
			intCount1=ubound(pcArray1,2)
						
			For j=0 to intCount1
				m1=0
				For k=1 to LCount
					if pcArray1(0,j)=AList(k) then
						m1=1
						exit for
					end if
				Next
						
				if m1=0 then
					LCount=LCount+1
					AList(LCount)=pcArray1(0,j)
				end if
			Next
		end if
					
		if pcv_checkvalue=1 then
			pcv_tempList=pcv_tempList & "<option value=""" & pcArray(0,i) & """>" & pcArray(1,i) & "</option>" & vbcrlf
						
			For k=1 to LCount
				For m=k+1 to LCount
					if AList(k)>AList(m) then
						pcv_t1=AList(m)
						AList(m)=AList(k)
						AList(k)=pcv_t1
					end if
				Next
			Next

			pcv_tempFunc=pcv_tempFunc & "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
			pcv_tempFunc=pcv_tempFunc & "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
			pcv_tempFunc=pcv_tempFunc & "SelectA.options.length = 0;" & vbcrlf
			pcv_tempFunc=pcv_tempFunc & "SelectA.options[0]=new Option(""All"","""");" & vbcrlf
			For j=1 to LCount
				pcv_tempFunc=pcv_tempFunc & "SelectA.options[" & j & "]=new Option(""" & AList(j) & """,""" & AList(j) & """);" & vbcrlf
			Next
			pcv_tempFunc=pcv_tempFunc & "}" & vbcrlf
						
		end if
			
	Next
					
	pcv_tempList=pcv_tempList & "</select>" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "}" & vbcrlf
	pcv_tempFunc=pcv_tempFunc & "</script>" & vbcrlf
			
	set rsTempFP=nothing
	%>
  <tr>
		<td nowrap>Filter by:</td>
		<td>
	<%=pcv_tempFunc%>
	<%=pcv_tempList%>
	&nbsp;
	<select name="SearchValues">
	<option value="">All</option>
	</select>
	<script>
	CheckList(document.ajaxSearch.customfield.value);
	</script>
	</font>
	</td>
	</tr>
<%else%>
<input type="hidden" name="customfield" value="0">
<%end if %>
<!-- end of custom fields -->

<!--Product Prices-->

<tr> 
	<td>Price:</td>
	<td>From:&nbsp;<input type="text" name="priceFrom" size="6" value="0">&nbsp;To:&nbsp;<input type="text" name="priceUntil" size="6" value="999999999">
	</td>
</tr>

<!--Product Inventory-->
          
<tr> 
	<td>In stock:</td>
	<td>  
	<input type="checkbox" name="withstock" value="1"  class="clearBorder">
	</td>
</tr>

<!--Product SKU-->
          
<tr> 
	<td>Part number (SKU):</td>
	<td> 
	<input name="sku" type="text" size="15" maxlength="150"></td>
</tr>

<!--Get Brands-->
          
<%query="Select IDBrand,BrandName from Brands order by BrandName asc"
set rsTempFP=connTemp.execute(query)
if not rsTempFP.eof then
pcArr=rsTempFP.getRows()
intCount1=ubound(pcArr,2)%>
<tr> 
	<td>Brand:</td>
	<td>
	<select name="IDBrand">
	<option value="0" selected>All</option>
	<%For m=0 to intCount1%>
	<option value="<%=pcArr(0,m)%>"><%=pcArr(1,m)%></option>
	<%Next%>
	</select></td>
</tr>
<%end if
set rsTempFP=nothing%>

<!--Search Keywords-->
          
<tr> 
	<td valign="top">Keyword(s):</td>
	<td> 
	<input type="text" name="keyWord" size="40">
	<br>
	<input type="checkbox" name="exact" value="1"  class="clearBorder"> Search on exact phrase
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td>Filter by Product Type:</td>
	<td>Standard Products	<input type="checkbox" name="src_IncNormal" value="1" checked class="clearBorder">
		&nbsp;&nbsp;&nbsp;&nbsp;BTO Products <input type="checkbox" name="src_IncBTO" value="1" class="clearBorder">&nbsp;&nbsp;&nbsp;&nbsp;BTO Items	<input type="checkbox" name="src_IncItem" value="1" class="clearBorder">
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input type="checkbox" name="pinactive" value="1" class="clearBorder"> Include inactive products
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<input type="checkbox" name="pdeleted" value="1" class="clearBorder"> Include deleted products
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td valign="top">Products set as &quot;Specials&quot;:</td>
	<td valign="top">
		<input type="radio" name="pSpecial" value="1" class="clearBorder"> Only &quot;Specials&quot;<br>
		<input type="radio" name="pSpecial" value="2" class="clearBorder"> Exclude &quot;Specials&quot;<br>
		<input type="radio" name="pSpecial" value="0" checked class="clearBorder"> Any (disregard this filter)
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td valign="top">Products set as &quot;Featured&quot;:</td>
	<td valign="top"><input type="radio" name="pFeatured" value="1" class="clearBorder"> Only &quot;Featured&quot; products<br>
		<input type="radio" name="pFeatured" value="2" class="clearBorder"> Exclude &quot;Featured&quot; products<br>
		<input type="radio" name="pFeatured" value="0" checked class="clearBorder"> Any (disregard this filter)
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td colspan="2">Date Created</td>
</tr>
<tr>
	<td valign="top">From Date:</td>
	<td valign="top">
		<input type="text" name="pFromDate" value="" size="13"> <a href="javascript:CalPop('document.ajaxSearch.pFromDate');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td valign="top">To Date:</td>
	<td valign="top">
		<input type="text" name="pToDate" value="" size="13"> <a href="javascript:CalPop('document.ajaxSearch.pToDate');"><img SRC="../Calendar/icon_Cal.gif" border="0"></a>
	</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr> 
	<td nowrap>Sort by:</td>
	<td>
	<select name="order">
		<option value="0">Description ascending</option>
		<option value="1">Description descending</option>
		<option value="2">SKU ascending</option>
		<option value="3">SKU descending</option>
		<option value="4" selected>ProductID ascending</option>
		<option value="5">ProductID descending</option>
	</select>
	<td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<tr>
	<td colspan="2">
		<b>Include/Exclude Previously Exported Products</b>
	</td>
</tr>
<tr>
	<td colspan="2">
		Would you like to export <u><b>only</b></u> products that <u><b>have not</b></u> previously been exported? <input type="radio" name="pHideExported" value="1" class="clearBorder" checked> Yes    <input type="radio" name="pHideExported" value="0" class="clearBorder"> No</td>
</tr>
<tr>
	<td colspan="2"><hr /></td>
</tr>
<%query="SELECT pcXP_ID,pcXP_PartnerID,pcXP_Name From pcXMLPartners WHERE pcXP_Status=1 AND pcXP_FTPHost<>'' AND pcXP_FTPDirectory<>'' AND pcXP_FTPUsername<>'' AND pcXP_FTPPassword<>'';"
set rs=connTemp.execute(query)
if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	%>
	<tr>
	<td colspan="2">
		<h3>Upload Exported XML File to FTP Server</h3>
		Please choose a XML Partner that you want to upload exported XML file to their FTP Server
	</td>
	</tr>
	<tr>
		<td>XML Partner:</td>
		<td>
			<select name="pFTPPartner">
				<option value="0"></option>
			<%For i=0 to intCount%>
				<option value="<%=pcArr(0,i)%>"><%=pcArr(1,i)%><%if trim(pcArr(2,i))<>"" then%>&nbsp;(<%=pcArr(2,i)%>)<%end if%></option>
			<%Next%>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right">
			<input type="checkbox" name="pRmvFile" value="1" class="clearBorder">
		</td>
		<td>
			Remove Exported XML File after uploading it to the Partner FTP Server
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr /></td>
	</tr>
<%end if
set rs=nothing%>
<tr>
	<td colspan="2">
		<input name="runform" type="submit" value="Export Product XML File" id="searchSubmit">
	</td>
</tr>  
</table>
</form>
<!--End of Search Form-->
<%call closedb()%><!--#include file="AdminFooter.asp"-->