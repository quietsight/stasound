<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'******************************************************************
'// USE CACHE
'// To change this value from the default value of false
'// you will need to change the variable below to the value of true.
'//
'// For Example: 
'// pcv_strCache = "true"

'******************************************************************
pcv_strCache = "false"
'******************************************************************
%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryData.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryNestedXMLDataSet.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/xpath.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryEffects.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryTooltip.js"></script>
<script language="JavaScript" type="text/javascript" src="../includes/spry/SpryPagedView.js"></script>
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<link type="text/css" rel="stylesheet" href="pcSyndication.css" />
<script type="text/javascript">
<%
'// Get the Affiliate ID (if any)
idaffiliate=""
idaffiliate=Request("idaffiliate")
%>
<% if SNW_TYPE="1" then %>
	var pcMasterSet = new Spry.Data.XMLDataSet("pcSyndication_GetItems.asp", "Products/Product", {useCache: <%=pcv_strCache%>});
	var pcItems = new Spry.Data.PagedView( pcMasterSet ,{pageSize: 3}, {useCache: <%=pcv_strCache%>});
<% else %>
	var pcMasterSet = new Spry.Data.XMLDataSet("pcSyndication.xml", "Products/Product", {useCache: <%=pcv_strCache%>});
	var pcItems = new Spry.Data.PagedView( pcMasterSet ,{pageSize: 3}, {useCache: <%=pcv_strCache%>});
<% end if %>
</script>
<html>
    <body style="background-color:transparent">
        <div id="pcSyndication" class="pcSyndicationRegion">
            <div id="pcProductRegion" spry:region="pcItems" class="SpryHiddenRegion">		
                <div spry:state="ready">
                    <p id="pcSyndicationBox" align="center" onMouseOver="SetRow('{ds_RowNumber}','{ds_PageNumber}','{ds_PageSize}');" spry:repeat="pcItems">  
                        <span spry:if="'{SmallImage}'.indexOf('http') != -1;">
                            <img src="{SmallImage}"><br/>
                        </span>				
                        <span class="pcSyndicationName"><a href="{URL}&idaffiliate=<%=idaffiliate%>" target="_blank" class="SyndicationImage">{Description}</a></span>
                        <br/><span class="pcSyndicationPrice"><%=scCurSign%>{Price}</span>
                    </p>
              	</div>
                <div spry:state="loading">
                    <div>loading...</div>			
                </div>	
                <div spry:state="error">
                    <div>No Items</div>			
                </div>	
            </div>
        
            <div id="PagingContainer" style="display:none">		
                <div align="center" spry:region="pcItems" class="SpryHiddenRegion">
                    <p spry:if="{ds_UnfilteredRowCount} > 0">				 
                        Showing {ds_PageFirstItemNumber} - {ds_PageLastItemNumber} of {ds_UnfilteredRowCount} 
                    </p>
                    <p spry:if="{ds_UnfilteredRowCount} == 0"></p>
                    <p spry:if="{ds_UnfilteredRowCount} > 0">				
                        <a href="JavaScript:;" onClick="pcItems.previousPage();">Previous</a>&nbsp; | &nbsp;<a href="JavaScript:;" onClick="pcItems.nextPage();">Next</a>
                    </p>
                </div>
            </div>
        </div>
    </body>
</html>
<script type="text/javascript">			
	var myObserver = new Object;
	myObserver.onPostUpdate = function(notifier, data) {
		InitPaging();
	};	
	Spry.Data.Region.addObserver('pcProductRegion', myObserver);	
	function InitPaging() {
		document.getElementById("PagingContainer").style.display = '';
	}
	function SetRow(a, b, c) {
		RowNumber=parseInt(a);
		PageNumber=parseInt(b);
		PageSize=parseInt(c);
		pActualPage = (PageSize * (PageNumber-1));
		pActualRow = RowNumber + pActualPage;
		pcMasterSet.setCurrentRow(pActualRow);
	}	
</script>