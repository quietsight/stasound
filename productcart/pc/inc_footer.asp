<%
private const scIncFooter="1"

dim tempFooterURL
tempFooterURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
tempFooterURL=replace(tempFooterURL,"https:/","https://")
tempFooterURL=replace(tempFooterURL,"http:/","http://")
%>
<%
'// Restore Cart
if session("NeedToShowRSCMsg")="1" then
	session("NeedToShowRSCMsg")=""

	'Get Path Info
	pcv_filePath = Request.ServerVariables("PATH_INFO")
	do while instr(pcv_filePath,"/")>0
		pcv_filePath = mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
	loop

	pcv_Query = Request.ServerVariables("QUERY_STRING")
	If len(pcv_Query)>0 Then
		If instr(pcv_filePath,"404.asp")>0 AND instr(pcv_Query,"404;")>0 Then
			pcv_filePath = Right(pcv_Query,Len(pcv_Query)-4)
		Else
			pcv_filePath = pcv_filePath & "?" & pcv_Query
		end if
	End If

	session("SFClearCartURL")=pcv_filePath
	%>
	<div id="GlobalMsgDialog" title="<%=dictLanguage.Item(Session("language")&"_CustSavedCarts_6")%>" style="display:none">
		<div id="GlobalMsg" style="width: 80%;">
			<div class="ui-main">
				<div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_CustSavedCarts_5")%></div>
			</div>
		</div>
	</div>
	<script>
		$(document).ready(function()
		{
			$("#GlobalMsgDialog").dialog({
					bgiframe: true,
					autoOpen: false,
					resizable: false,
					width: 400,
					minHeight: 50,
					modal: true,
					buttons: {
						' <%=dictLanguage.Item(Session("language")&"_opc_js_77")%> ': function() {
								location='<%=tempFooterURL%>CustLOb.asp';
								$(this).dialog('close');
						},
						' <%=dictLanguage.Item(Session("language")&"_opc_js_65")%> ': function() {
								location='<%=tempFooterURL%>viewcart.asp';
								$(this).dialog('close');
						},
						' <%=dictLanguage.Item(Session("language")&"_opc_js_78")%> ': function() {
								$(this).dialog('close');
						}
					}
			});
			$("#GlobalMsgDialog").dialog('open');
		});
	</script>
<%else
	session("SFClearCartURL")=""
end if
session("MobileURL")=""
session("idProductRedirect")=""
%>
<!--#include file="inc-GoogleAnalytics.asp"-->
<% if pcInterest=1 then %>
<script type="text/javascript" src="//assets.pinterest.com/js/pinit.js"></script>
<% end if %>
<% 
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing %>