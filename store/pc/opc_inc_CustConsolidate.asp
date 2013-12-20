<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
'// 0 and has 1 or 2
tmpEmail=pEmail
if Instr(pEmail,"'")>0 then
	pEmail=replace(pEmail,"'","''")
end if
if Session("CustomerGuest")="0" then
	query = "SELECT idCustomer FROM customers WHERE email = '" & pEmail & "' AND (pcCust_Guest=1 OR pcCust_Guest=2)"
	set rsCustC = Server.CreateObject("ADODB.Recordset")
	set rsCustC = conntemp.execute(query)
	if NOT rsCustC.EOF then
		pcs_CustConsolidate()
	end if	
	set rsCustC = nothing
end if

'// 1 and has a 0
if Session("CustomerGuest")="1" then
	query = "SELECT idCustomer FROM customers WHERE email = '" & pEmail & "' AND (pcCust_Guest=0)"
	set rsCustC = Server.CreateObject("ADODB.Recordset")
	set rsCustC = conntemp.execute(query)
	if NOT rsCustC.EOF then
		pcs_CustConsolidate()
	end if
	set rsCustC = nothing
end if

'// 2 because we know it has a 0
if Session("CustomerGuest")="2" then
	query = "SELECT idCustomer FROM customers WHERE email = '" & pEmail & "' AND (pcCust_Guest=0)"
	set rsCustC = Server.CreateObject("ADODB.Recordset")
	set rsCustC = conntemp.execute(query)
	if NOT rsCustC.EOF then
		pcs_CustConsolidate()
	end if
	set rsCustC = nothing
end if
pEmail=tmpEmail
%>
<% Public Sub pcs_CustConsolidate %>
<div id="ConArea">
  <table class="pcShowContent">
    <tr>
        <td>
        <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_opc_cons_1")%></div>
        </td>
    </tr>
    <tr>
        <td>
        <div style="padding-left: 10px;"><input type="button" id="ConSubmit" value="<%=dictLanguage.Item(Session("language")&"_opc_cons_2")%>" class="submit2">&nbsp;&nbsp;<input type="button" id="LogOut" value="<%=dictLanguage.Item(Session("language")&"_opc_cons_3")%>" onclick="location='CustLO.asp';" class="submit2"></div>
        </td>
    </tr>
  </table>
</div>
<div id="ConLoader" style="display:none"></div>
<script>
	$(document).ready(function()
	{
		jQuery.validator.setDefaults({
			success: "valid"
		});
		
		//*Ajax Global Settings
		$("#GlobalAjaxErrorDialog").ajaxError(function(event, request, settings){
			$(this).dialog('open');
			$("#ConLoader").hide();
		});
	
		
		//*Dialogs
		$("#GlobalAjaxErrorDialog").dialog({
				bgiframe: true,
				autoOpen: false,
				resizable: false,
				width: 450,
				height: 230,
				modal: true,
				buttons: {
					' OK ': function() {
							$(this).dialog('close');
						}
				}
		});
		
		//* Consolidate Account
		$('#ConSubmit').click(function(){
				$("#ConLoader").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_cons_4"))%>');
				$("#ConLoader").show();	
				$.ajax({
					type: "POST",
					url: "opc_sendConMail.asp",
					data: "{}",
					timeout: 45000,
					success: function(data, textStatus){
						if (data=="SECURITY")
						{
							$("#ConLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_cons_5"))%>');
							
						} else {
							
							if (data=="OK")
							{
								$("#ConLoader").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_cons_6"))%>');
								var callbackCon=function (){}
							
							} else {
								
								$("#ConLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+ data);
							}
						}
					}
				});
				return(false);
		});
		
	});
</script>
<% End Sub %>