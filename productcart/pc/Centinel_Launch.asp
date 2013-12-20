<%@ language="vbscript" %>
<% 'option explicit %>
<%response.expires=-1%>
<%response.buffer=true%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!-- #include file="Centinel_Config.asp"-->
<% dim conntemp, rs, query
'==========================================================================================
'= CardinalCommerce (http://www.cardinalcommerce.com)
'= Page used to POST the transaction request to the ACSURL
'==========================================================================================
%>

<html>
<head><TITLE>Processing Your Transaction</TITLE>
<script language="javascript">
	function onLoadHandler(){
		document.frmLaunchACS.submit();
	}
</script>
</head>
<body onLoad="onLoadHandler();">
<center>
<%
	'=====================================================================================
	' The inline frame window must be a minimum of 400 pixel width by 400 pixel height.
	'=====================================================================================
%>
<form name="frmLaunchACS" method="post" action="<%=Session("Centinel_ACSURL")%>">
<noscript>
	<br><br>
	<center>
	<font color="red">
	<h1>Processing your transaction</h1>
	<h2>Javascript is currently disabled or is not supported by your browser.<br></h2>
	<h3>Please click submit to continue	the processing of your transaction.</h3>
	</font>
	<input type="submit" value="submit">
	</center>
</noscript>
<input type=hidden name="PaReq" value="<%=Session("Centinel_Payload")%>">
<input type=hidden name="TermUrl" value="<%=Session("Centinel_TermURL")%>">
<input type=hidden name="MD" value="">
</form>
</center>
</body>
</html>