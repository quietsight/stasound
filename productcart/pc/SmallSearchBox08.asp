<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

  '// Locate preferred results count and load as default
    Dim pcIntPreferredCountSearch
    pcIntPreferredCountSearch =(scPrdRow*scPrdRowsPerPage)
%>
<form action="showsearchresults.asp" name="search" method="get">
	<input type="hidden" name="pageStyle" value="<%=bType%>">
	<input type="hidden" name="resultCnt" value="<%=pcIntPreferredCountSearch%>">
	<input type="Text" name="keyword" size="14" value="" id="smallsearchbox" >
    <a href="javascript:document.search.submit()" onclick="pcf_CheckSearchBox();" title="Search"><img src="images/pc2009-search.png" border="0" alt="Search" align="absbottom"></a>
    <div style="margin-top: 3px;">
		<a href="search.asp">More search options</a>
	</div>
</form>
<script language="JavaScript">
<!--
function pcf_CheckSearchBox() {
 	pcv_strTextBox = document.getElementById("smallsearchbox").value;
	if (pcv_strTextBox != "") {
		document.getElementById('small_search').onclick();
	}
}
//-->
</script>
<%
Response.Write(pcf_InitializePrototype())
response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_advSrca_23"), "small_search", 200))
%>