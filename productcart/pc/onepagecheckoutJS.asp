<%
pcv_IsBillingPhoneReq = false
%>
<script>
var CountriesRequireZipCode="US, CA, GB, " 
var CountriesRequireStateCode="US, CA, "
var CountriesRequireProvince=""
var OPCCheck = 1;
var OPCFinal = 0;
var OPCFree = 0;
var AskEnterPass=0;
var AcceptEnterPass=0;
// JavaScript Document

function btnShow1(tmpName,tmpTab)
{
	if (tmpName=="OK")
	{
		$("#btnOK"+tmpTab).show();
		$("#btnEdit"+tmpTab).hide();
		$("#btnError"+tmpTab).hide();
	}
	if (tmpName=="Edit")
	{
		$("#btnOK"+tmpTab).hide();
		$("#btnEdit"+tmpTab).show();
		$("#btnError"+tmpTab).hide();
	}
	if (tmpName=="Error")
	{
		$("#btnOK"+tmpTab).hide();
		$("#btnEdit"+tmpTab).hide();
		$("#btnError"+tmpTab).show();
	}
}

function btnShow2(tmpTab)
{
	$("#btnOK"+tmpTab).show();
	$("#btnEdit"+tmpTab).show();
	$("#btnError"+tmpTab).hide();
}

function btnHide2(tmpTab)
{
	$("#btnEdit"+tmpTab).hide();
	$("#btnError"+tmpTab).hide();
}

function btnHideAll(tmpTab)
{
	$("#btnOK"+tmpTab).hide();
	$("#btnEdit"+tmpTab).hide();
	$("#btnError"+tmpTab).hide();
}

$(document).ready(function()
{
	 window.getSC = function(){  
          getShipContents("");
     }
	jQuery.validator.setDefaults({
		success: function(element) {
			$(element).parent("td").children("input, textarea").addClass("success")
		}
	});

	jQuery.validator.addMethod("credit_card", function(value, element) {
		return  this.optional(element) || /^[0-9\ ]+$/i.test(value);	
	},"<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_1"))%>");
	
	//*Ajax Global Settings
	$("#GlobalAjaxErrorDialog").ajaxError(function(event, request, settings){
		$(this).dialog('open');
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
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_66"))%> ': function() {
						location='OnePageCheckout.asp';
						$(this).dialog('close');						
				},
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_65"))%> ': function() {
						location='viewcart.asp';
						$(this).dialog('close');						
				}
			},
			close: function() {}
	});
	
	$("#GlobalErrorDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 500,
			height: 260,
			modal: true,
			buttons: {
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_66"))%> ': function() {						
						location='OnePageCheckout.asp';
						$(this).dialog('close');						
				}
			},
			close: function() {}
	});
	
	$("#ValidationDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 500,
			height: 260,
			modal: true,
			buttons: {
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_66"))%> ': function() {						
						$(this).dialog('close');						
				}
			},
			close: function() {}
	});
	
	//*Terms Dialog
	$("#TermsDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			width: 450,
			height: 250,
			modal: true
	});
	
	// 'SB S - Agreement Dialog
	$("#sb_TermsDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			width: 450,
			height: 250,
			modal: true
	});
	// 'SB E - Agreement Dialog

	//*Please Wait Dialog
	$("#PleaseWaitDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 250,
			minHeight: 50,
			modal: true
	});
	
	//*Ask to Enter password Dialog
	$("#AskEnterPassDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 450,
			minHeight: 50,
			modal: true,
			buttons: {
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_51"))%> ': function() {						
						AcceptEnterPass=0;
						$(this).dialog('close');
						$('#PwdLoader').hide();
						$('#PwdWarning').hide();
						$('#PwdArea').hide();
				},
				' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_51a"))%> ': function() {						
						AcceptEnterPass=1;
						$(this).dialog('close');
						$('#newPass1').focus();
				}				
			},
			close: function() {AcceptEnterPass=1;}
	});
	//MailUp-S
	$("#PleaseWaitDialogMU").dialog({
			bgiframe: true,
			autoOpen: false,
			resizable: false,
			width: 460,
			height: 80,
			minHeight: 80,
			modal: true
	});	
	//MailUp-E
	
	//*Gift Wrapping Dialog
	$("#GWDialog").dialog({
			bgiframe: true,
			autoOpen: false,
			width: 650,
			minHeight: 50,
			modal: true,
			open: function(event,ui)
			{
				$('#GWframeloader').show();
				$('#GWframe').hide();
			},
			close: function() {
				
			}
	});
	
	$('#GWframe').load( function() {
		$('#GWframeloader').hide();
		$('#GWframe').show();
	} );
	
	//*Validate LoginForm
	var validator0=$("#loginForm").validate({
		rules: {
			email: {
				required: true,
				email: true
		     },
			password: 
			{
				required: true
			}
		},
		messages: {
			email: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_2"))%>",
				email: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_3"))%>"
			},
			password: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_4"))%>"
			}
		}
	})
	
	$('#LoginSubmit').click(function(){
		if ($('#loginForm').validate().form())
		{

			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						$("#LoginLoader").hide();
						location="viewcart.asp";
					}
				}
	 		});

			$.ajax({
				type: "POST",
				url: "opc_checklogin.asp",
				data: "email=" + $('#email').val()+"&password=" + $('#password').val()+"&securityCode=" + $('#securityCode').val()+"&CAPTCHA_Postback=" + $('#CAPTCHA_Postback').val(),
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="OK")
					{
						$("#LoginLoader").hide();
						location="onepagecheckout.asp";
					}
					else
					{
						$("#LoginLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> ' + data);
						$("#LoginLoader").show();
						var CaptchaTest = $("#securityCode").val();
						if (CaptchaTest != null) { 
							reloadCAPTCHA(); 
						}
						validator0.resetForm();
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
	
	//*Validate Billing Form
	var validator1=$("#BillingForm").validate({
		rules: {
			billfname: "required",
			billlname: "required",
			billaddr: "required",
			billcity: "required",
			billcountry: "required",
			<% if pcv_IsBillingPhoneReq then %>
			billphone: "required",
			<% end if %>
			<% if pcv_isVatIdRequired then %>
			billVATID: "required",
			<% end if %>
			<% if pcv_isSSNRequired then %>
			billSSN: "required",
			<% end if %>
			billzip:
			{
				required: function(element)
				{
					var str=document.BillingForm.billcountry.value;
					if (CountriesRequireZipCode.indexOf(str + ",")>=0)
					{ return(true) }
					else
					{ return(false) }
				}
			},
			billstate:
			{
				required: function(element)
				{
					var str=document.BillingForm.billcountry.value;
					if (CountriesRequireStateCode.indexOf(str + ",")>=0)
					{ return(true) }
					else
					{ return(false) }
				}
			},			

			<%if Not (Session("idCustomer")>0 AND session("CustomerGuest")="0") then%>
			billemail: {
				required: true,
				email: true<%if scGuestCheckoutOpt="2" then%>,
				remote: {
    				    url: "opc_checkEmail.asp",
			    	    type: "POST"
	        	}
				<%end if%>
			},
			<%if scGuestCheckoutOpt="2" then%>
			billpass: 
			{
				required: true
			},
			billrepass:
			{
				<%if scGuestCheckoutOpt="2" then%>
				required: true,
				<%end if%>
				equalTo: "#billpass"
			},
			<%end if%>
			<%end if%>
			
			billprovince:
			{
				required: function(element)
				{
					var str=document.BillingForm.billcountry.value;
					if (CountriesRequireProvince.indexOf(str + ",")>=0)
					{ return(true) }
					else
					{ return(false) }
				}
			}

		},
		messages: {
			billfname: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_8"))%>"
			},
			billlname: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_9"))%>"
			},
			billaddr: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_10"))%>"
			},
			billcity: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_11"))%>"
			},
			billcountry: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_12"))%>"
			},
			<% if pcv_IsBillingPhoneReq then %>
			billphone: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_16"))%>"
			},
			<% end if %>
			<% if pcv_isVatIdRequired then %>
			billVATID: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_Custmoda_27"))%>"
			},
			<% end if %>
			<% if pcv_isSSNRequired then %>
			billSSN: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_Custmoda_25"))%>"
			},
			<% end if %>
			billzip: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_13"))%>"
			},
			billstate: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_15"))%>"
			},
			billprovince: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_63"))%>"
			},
			billemail: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_2"))%>",
				email: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_3"))%>",
				remote: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_5a"))%>"
			},
			billpass: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_4"))%>"
			},
			billrepass: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_47"))%>",
				equalTo: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_48"))%>"
			}

		}
	});
	$('#BillingSubmit').click(function(){		
		
		// 'SB S
		if (pcCustomerRegAgreed=='0') 
		{ 
			if ($("#sb_AgreeTerms").is(':checked')) 
			{
				updTermsSB();
			}
			else
			{
				$("#BillingLoaderSB").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_login_5"))%>');
				$("#BillingLoaderSB").show();
				btnShow1("Error","CO");
			  	return(false);
			}
		}
		// 'SB E
		
		if (pcCustomerTermsAgreed=='0') 
		{ 
			if ($("#AgreeTerms").is(':checked')) 
			{
				updTerms();
			}
			else
			{
				$("#BillingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_login_5"))%>');
				$("#BillingLoader").show();
				btnShow1("Error","CO");
			  	return(false);
			}
		}	
		
		if ($('#BillingForm').validate().form())
		{
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

			<%'MailUp-S
			if session("SF_MU_Setup")="1" then%>
			if (tmpNListChecked) $("#PleaseWaitDialogMU").dialog('open');
			<%end if
			'MailUp-E%>
			$.ajax({
				type: "POST",
				url: "opc_UpdBillAddr.asp",
				data: $('#BillingForm').formSerialize(),
				timeout: 45000,
				success: function(data, textStatus){	
					<%'MailUp-S
					if session("SF_MU_Setup")="1" then%>
					if (tmpNListChecked) $("#PleaseWaitDialogMU").dialog('close');
					<%end if
					'MailUp-E%>
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if ((data=="ZIPLENGTH"))					
					{	
						$("#ValidationErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_74"))%>');
						$("#ValidationDialog").dialog('open');				
					}
					else if (data=="ERROR")
					{
						$("#BillingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_57"))%>');
						$("#BillingLoader").show();
						btnShow1("Error","CO");
						validator1.resetForm();
					}
					<%'MailUp-S
					if session("SF_MU_Setup")="1" then%>
					else if ((data.indexOf("OK")>=0) || (data.indexOf("NEW")>=0))
					<%else%>
					else if ((data=="OK") || (data=="NEW"))
					<%end if
					'MailUp-E%>
					{						
						
						if (HaveShipArea==1)
						{

							// SHIPPING:  Check the options, open shipping panel
							
							  if (NeedLoadShipContent==1)
							  {
								  getShipContents("");
							  }
							  else
							  {
								  $('#ShippingArea').show();
							  }

						} else {
							
							// NO SHIPPING: Do tax, save order, and goto payment panel
							$("#ShipLoadContentMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_22"))%>');
							$("#ShipLoadContentMsg").show();
							getShipChargeContents("");
						}
						
						// Show Password?
						<%if scGuestCheckoutOpt=0 OR scGuestCheckoutOpt="" OR scGuestCheckoutOpt=1 then%>
						<%'MailUp-S
						if session("SF_MU_Setup")="1" then%>
						if ((data.indexOf("NEW")>=0) || ('<%=session("CustomerGuest")%>'=='1'))
						<%else%>
						if ((data=="NEW") || ('<%=session("CustomerGuest")%>'=='1'))
						<%end if
						'MailUp-E%>
						{							  
							$("#PwdArea").show(); // Display Optional password selection (in payment area)
						}
						else
						{	
							$("#PwdArea").hide(); // No optional password selection
						}
						<%end if%>

						// Hide Billing Panel / Display Billing Edit Link
						$("#BillingLoader").hide();
						//'SB S
						$("#BillingLoaderSB").hide();
						//'SB E
						$("#BillingArea").hide();
						<%if scGuestCheckoutOpt=2 then%>
						$("#billArea1").hide();
						$("#billArea2").hide();
						$("#billArea3").hide();
						<%end if%>
						<%'MailUp-S
						if session("SF_MU_Setup")="1" then%>
						var tmpArr=data.split("|s|");
						if (tmpArr[1]!="") $("#MailUpArea").html(tmpArr[1]);
						<%end if
						'MailUp-E%>
						getBillAddress("");	 
						
					}
					else
					{
						$("#BillingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> ' + data);
						$("#BillingLoader").show();
						btnShow1("Error","CO");
						validator1.resetForm();
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
	
	$('#BillingCancel').click(function(){	
									   
		// Hide Billing Panel / Show Login Area
		$("#BillingLoader").hide();
		//'SB S
		$("#BillingLoaderSB").hide();
		//'SB E
		$("#BillingArea").hide();
		$('#LoginOptions').show();
		$('#acc1').hide(); 

	});


	//* Get Billing Address
	function getBillAddress(tmpvalue)
	{
		$("#ShippingLoader").html('');
		$("#ShippingLoader").hide();
		$.ajax({
				type: "POST",
				url: "opc_getbilladdr.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if (data=="ERROR")
					{
						$("#BillingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_57"))%>');
						$("#BillingLoader").show();
						btnShow1("Error","CO");
						return(false);
					}
					else
					{

						$("#BillingAddress").html("" + data);
						$("#BillingAddress").show();
						btnShow1("OK","CO");
						
					}
				}
	 		});
	}


	//* Get Shipping Address
	function getShipAddress(tmpArr)
	{
		$("#ShipChargeLoadContentMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_54"))%>');
		$("#ShipChargeLoadContentMsg").show();
			$.ajax({
			type: "GET",
			async: false,
			url: "opc_cartcheck.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data!="OK")
				{
					location="viewcart.asp";
				}
			}
		});

		$.ajax({
				type: "POST",
				url: "opc_getshipaddr.asp",
				data: "{}",
				timeout: 45000,
				error: function (XMLHttpRequest, textStatus, errorThrown) {
					$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_76"))%>');
					$("#GlobalErrorDialog").dialog('open');
					return false;
				},
				global: false,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if (data=="ERROR")
					{

						$("#ShipChargeLoadContentMsg").hide();
						$("#ShippingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_57"))%>');
						$("#ShippingLoader").show();
						btnShow1("Error","CO");
						return(false);
						
					}
					else
					{
						if (tmpArr!="") {
							getShipChargeContents(tmpchkFree);
							$("#ShippingLoader").hide();
						}
						$("#ShippingAddress").html("" + data);
						$("#ShippingAddress").show(); 
						btnShow1("OK","CO");
						
					}
				}
	 		});
	}

	//*Prepare Contents of Shipping Form
	function getShipContents(tmpvalue)
	{
			$.ajax({
			type: "GET",
			async: false,
			url: "opc_cartcheck.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data!="OK")
				{
					location="viewcart.asp";
				}
			}
		});

		$('#ShippingArea').hide();
		$.ajax({
				type: "POST",
				url: "opc_genshipselect.asp",
				data: "{}",
				timeout: 90000,
				error: function (XMLHttpRequest, textStatus, errorThrown) {
					$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_76"))%>');
					$("#GlobalErrorDialog").dialog('open');
					return false;
				},
				global: false,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if (data=="ERROR")
					{
						btnShow1("Error","Ship");
						$('#ShippingArea').hide(); 
						$("#ShipLoadContentMsg").hide();
						
					}
					else
					{
						ShipContents=data;
						generateShipDrop(tmpvalue);
					}
				}
	 		});
	}

	//*Generate Shipping Drop-down
	function generateShipDrop(tmpvalue)
	{
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// Start: Radio Buttons
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		count=0
		
		var strRadioBtns = '';
		strRadioBtns = strRadioBtns + '<input id="rad_'+count+'" type="radio" name="ShipArrOpts" value="-1" onclick="javascript:FillShipForm(this.value,0);" />' + ' ' + '<label for="rad_'+count+'">' + "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_20"))%>" + '</label><br />';

		if (HaveGRAddress==1) {
			count=count+1
			strRadioBtns = strRadioBtns + '<input id="rad_'+count+'" type="radio" name="ShipArrOpts" value="-2" onclick="javascript:FillShipForm(this.value,0);" />' + ' ' + '<label for="rad_'+count+'">' +  "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_21"))%>" + '</label><br />';
		}
		
		var tmpHaveShipAddr=0;
		var SelectedShip="";
		var SelectedC="";
		if (ShipContents!="")
		{
			var tmpShipList=ShipContents.split("|$|");
			if (tmpShipList.length>0)
			{
				SelectedShip=tmpShipList[0];
				for (var i=1;i<tmpShipList.length;i++)
				{
					if (tmpShipList[i]!="")
					{
						tmpShipRe=tmpShipList[i].split("|*|")
						count=count+1
						strRadioBtns = strRadioBtns + '<input id="rad_'+count+'" type="radio" name="ShipArrOpts" value="' + tmpShipRe[0] + '" onclick="javascript:FillShipForm(this.value,0);" />' + ' ' + '<label for="rad_'+count+'">' + tmpShipRe[1] + '</label><br />';
						tmpHaveShipAddr=tmpHaveShipAddr+1;
						if (SelectedShip==tmpShipRe[0]) SelectedC=count;
					}
				}
			}
		}
		if ((tmpHaveShipAddr<2) || (tmpNewCust==0))
		if (CanCreateNewShip==1)
		{
			count=count+1
			strRadioBtns = strRadioBtns + '<input id="rad_'+count+'" type="radio" name="ShipArrOpts" value="ADD" onclick="javascript:FillShipForm(this.value,0);" />' + ' ' + '<label for="rad_'+count+'">' +  "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_38"))%>" + '</label><br />';
			if (SelectedShip=="ADD") SelectedC=count;
		}

		$("#radios").html(strRadioBtns);		
		
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		// End: Radio Buttons
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

		//if (tmpvalue!="") SelectA.value=tmpvalue;
		//FillShipForm(tmpvalue,1);
		<% If session("PPSA")="1" Then %>
			$("#rad_1").attr("checked", "checked");
			FillShipForm(<%=session("PPSAID")%>,0);
		<% Else %>
			if (SelectedShip!="")
			{
				if (SelectedShip=="-1") SelectedC=0;
				if (SelectedShip=="-2") SelectedC=1;
				$("#rad_"+SelectedC).attr("checked", "checked");
				FillShipForm(SelectedShip,1);
			}
			else
			{
			$("#rad_0").attr("checked", "checked");
			FillShipForm(-1,1);
			}
		<% End If %>
		
		$("#ShipLoadContentMsg").hide();
		$('#ShippingArea').show();
		
	}
	
	//*Pre-Load Shipping Address Contents
	if (NeedPreLoadShipContent==1) { getShipContents("");}	



	//*Validate Shipping Form
	var validator2=$("#ShippingForm").validate({
		// success: "valid",
		rules: {
			shipfname:
			{
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))}
			},
			shiplname:
			{
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))}
			},
			shipaddr: 
			{
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))}
			},
			shipcity:
			{
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))}
			},
			shipcountry:
			{
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))}
			},
			shipzip: 
			{
				required: function(element)
				{
					if (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))
					{
						var str=document.ShippingForm.shipcountry.value;
						if (CountriesRequireZipCode.indexOf(str + ",")>=0)
						{ return(true) }
						else
						{ return(false) }
					}
					else
					{ return(false) }
				}
			},
			shipstate:
			{
				required: function(element)
				{
					if (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))
					{
						var str=document.ShippingForm.shipcountry.value;
						if (CountriesRequireStateCode.indexOf(str + ",")>=0)
						{ return(true) }
						else
						{ return(false) }
					}
					else
					{ return(false) }
				}
			},
			shipprovince:
			{
				required: function(element)
				{
					if (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))
					{
						var str=document.ShippingForm.shipcountry.value;
						if (CountriesRequireProvince.indexOf(str + ",")>=0)
						{ return(true) }
						else
						{ return(false) }
					}
					else
					{ return(false) }
				}
			},

			shipemail: {
				required: function(element)
				{ return (($("input[name='ShipArrOpts']:checked").val()!="-1") && ($("input[name='ShipArrOpts']:checked").val()!="-2"))},
				email: true
			}
		},
		messages: {
			shipfname: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_25"))%>"
			},
			shiplname: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_26"))%>"
			},
			shipaddr: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_27"))%>"
			},
			shipcity: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_28"))%>"
			},
			shipcountry: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_29"))%>"
			},
			shipzip: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_30"))%>"
			},
			shipstate: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_32"))%>"
			},
			shipprovince: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_64"))%>"
			},
			shipemail: {
				email: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_3"))%>"
			},

			DF1: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_35"))%>"
			},
			TF1: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_36"))%>"
			}
			
		}
	});

	//*Submit Shipping Form
	$('#ShippingSubmit').click(function(){		

		if ($('#ShippingForm').validate().form())
		{
            var qtyok = true;
			$.ajax({
				type: "GET",
				async: false,
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});
			$.ajax({
				type: "POST",
				url: "opc_UpdShipAddr.asp",
				data: $('#ShippingForm').formSerialize(),
				timeout: 45000,
				error: function (XMLHttpRequest, textStatus, errorThrown) {
					if (textStatus=='timeout') {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_75"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					} else {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_76"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					}
				},
				global: false,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if ((data=="ZIPLENGTH"))					
					{	
						$("#ValidationErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_74"))%>');
						$("#ValidationDialog").dialog('open');				
					}
					else if (data=="ERROR")
					{
						$("#ShippingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_57"))%>');
						$("#ShippingLoader").show();
						btnShow1("Error","Ship");
						validator2.resetForm();
					}
					else
					{
						if (data.indexOf("OK")>=0)
						{
							var tmpArr=data.split("|*|")
							getShipAddress(tmpArr);							
							acc1.openPanel('opcShipping');
							GoToAnchor('opcShippingAnchor');

							
						}
						else
						{
							$("#ShippingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> ' + data);
							$("#ShippingLoader").show();
							btnShow1("Error","Ship");
							validator2.resetForm();
						}
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
	
	//*Submit Discounts & Reward Points Form
	$('#DiscSubmit').click(function(){	
				$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

			$.ajax({
				type: "POST",
				url: "opc_OrderVerify.asp",
				data: $('#DiscForm').formSerialize() + '&rtype=1',
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						if (data.indexOf("|***|OK|***|")>=0)
						{
							var tmpArr=data.split("|***|")
	
							$("#DiscLoader").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_42"))%>');
							$("#DiscountCode").val(tmpArr[2]);
							$("#UseRewards").val(tmpArr[4]);
							$("#pcOPCtotalAmount").text(tmpArr[7]);
							$("#OPRWarning").hide();
							$('#PayArea').show();
							OPCReady=tmpArr[8];
							tmpchkFree=tmpArr[9];
							if (tmpArr[8]=="NO") {
	
								$("#OPRWarning").hide();
								$("#OPRArea").show();
							}
							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
							}
							// Free Order - End
							ValidateGroup3();
							return;
						}
						else
						{
							var tmpArr=data.split("|***|")
							$("#DiscLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+tmpArr[3]);
							$("#DiscLoader").show();
							btnShow1("Error","DC");
							$("#DiscountCode").val(tmpArr[2]);
							$("#UseRewards").val(tmpArr[4]);
							$("#pcOPCtotalAmount").text(tmpArr[7]);
							$("#OPRArea").html(tmpArr[0]);
							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
							}
							// Free Order - End
							$("#OPRWarning").hide();
							$("#OPRArea").show();
						}
					}
					
					// 'SB S
					$.post("opc_OrderVerify.asp", { sbTax: "1" } );
					// 'SB E
					
					$("#PleaseWaitDialog").dialog('close');
				}
	 		});
			return(false);
	});
	
	//*Recalculate Discounts & Reward Points Form
	$('#DiscRecal').click(function(){						  
			$.ajax({
				type: "POST",
				url: "opc_OrderVerify.asp",
				data: $('#DiscForm').formSerialize() + '&rtype=1',
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
					if (data.indexOf("|***|OK|***|")>=0)
					{
						var tmpArr=data.split("|***|")

						$("#DiscLoader").html("");
						$("#DiscLoader").hide();
						$("#DiscountCode").val(tmpArr[2]);
						$("#UseRewards").val(tmpArr[4]);
						$("#pcOPCtotalAmount").text(tmpArr[7]);
						$("#OPRWarning").hide();
						$('#PayArea').show();
						OPCReady=tmpArr[8];
						tmpchkFree=tmpArr[9];						
						if (tmpArr[8]=="NO")
						{
							$("#OPRWarning").hide();
							$("#OPRArea").show();
						}
							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
								CheckPlaceOrder();
							}
							// Free Order - End
					}
					else
					{
						var tmpArr=data.split("|***|")
						$("#DiscLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+tmpArr[3]);
						$("#DiscLoader").show();
						btnShow1("Error","DC");
						$("#DiscountCode").val(tmpArr[2]);
						$("#UseRewards").val(tmpArr[4]);
						$("#pcOPCtotalAmount").text(tmpArr[7]);
						$("#OPRArea").html(tmpArr[0]);
							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
								CheckPlaceOrder();
							}
							// Free Order - End
						$("#OPRWarning").hide();
						$("#OPRArea").show();
					}
					
					// 'SB S
					$.post("opc_OrderVerify.asp", { sbTax: "1" } );
					// 'SB E
					
					GenOrderPreview(data,0);
					$("#DiscLoader1").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_42"))%>');
					$("#DiscLoader1").show();
					RunE("DiscLoader1");
					}
				}
	 		});
			return(false);
	});
	
	$('#RewardsRecal').click(function(){						  
			$.ajax({
				type: "POST",
				url: "opc_OrderVerify.asp",
				data: $('#DiscForm').formSerialize() + '&rtype=1',
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
					if (data.indexOf("|***|OK|***|")>=0)
					{
						var tmpArr=data.split("|***|")

						$("#DiscLoader").html("");
						$("#DiscLoader").hide();
						$("#DiscountCode").val(tmpArr[2]);
						$("#UseRewards").val(tmpArr[4]);
						$("#pcOPCtotalAmount").text(tmpArr[7]);
						$("#OPRWarning").hide();
						$('#PayArea').show();
						OPCReady=tmpArr[8];
						tmpchkFree=tmpArr[9];						
						if (tmpArr[8]=="NO")
						{

							$("#OPRWarning").hide();
							$("#OPRArea").show();
						}
					}
					else
					{
						var tmpArr=data.split("|***|")
						$("#DiscLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+tmpArr[3]);
						$("#DiscLoader").show();
						btnShow1("Error","DC");
						$("#DiscountCode").val(tmpArr[2]);
						$("#UseRewards").val(tmpArr[4]);
						$("#pcOPCtotalAmount").text(tmpArr[7]);
						$("#OPRArea").html(tmpArr[0]);
						$("#OPRWarning").hide();
						$("#OPRArea").show();
					}
					
					// 'SB S
					$.post("opc_OrderVerify.asp", { sbTax: "1" } );
					// 'SB E
							
					GenOrderPreview(data,0);
					$("#DiscLoader1").html('<img src="images/pcv4_st_icon_success_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_42"))%>');
					$("#DiscLoader1").show();
					RunE("DiscLoader1");
					}
				}
	 		});
			return(false);
	});

	//*Validate Other Order Information Form
	var validator3=$("#OtherForm").validate({
		rules:
		{
			shipemail: {
				email: true
			}
		},
		messages: {
			GcReEmail: {
				email: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_3"))%>"
			}
		}
	});
	
	//*Submit Other Order Information Form
	$('#OtherSubmit').click(function(){
				$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

			$.ajax({
				type: "POST",
				url: "opc_updotherinfo.asp",
				data: $('#OtherForm').formSerialize() + '&rtype=1',
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						if (data=="OK")
						{
							$("#OtherLoader").hide();
							ValidateGroup4();
							return;
						}
						else
						{
							$("#OtherLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> ' + data);
							$("#OtherLoader").show();
							btnShow1("Error","OT");
							validator3.resetForm();
							$("#PleaseWaitDialog").dialog('close');
						}
					}
					$("#PleaseWaitDialog").dialog('close');
				}
	 		});
			return(false);
	});
	
	
	//*Load the Payment Panel
	$('#LoadPaymentPanel').click(function(){
		$("#PleaseWaitMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_79"))%>');
		$("#PleaseWaitDialog").dialog('open');	
		$(".ui-dialog-titlebar").css({'display' : 'none'});
		$("#PleaseWaitDialog").css({'min-height' : '50px'});
		getBillAddress("");
		getShipAddress("");
		if (HaveShipArea==1) {
			getShipMethod();
		}
		if (NeedLoadShipChargeContent==1) {
			getShipContents();
			getShipChargeContents("");	
		} else {
			acc1.openPanel('opcPayment');
			$("#PleaseWaitDialog").dialog('close');	
		}
		GetOrderInfo("","#TaxLoadContentMsg",1,'');	
		
	});

	//*Submit Gift Wrapping Area
	$('#GWSubmit').click(function(){
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

		var tmpdata=getGWPrdList();	
		$.ajax({
			type: "POST",
			url: "opc_getGiftWrap.asp?action=setup&list=" + tmpdata,
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data=="LOAD")
				{
					document.getElementById("GWframe").src="opc_giftwrap.asp?action=setup&list=" + tmpdata;
					$("#GWDialog").dialog("open");
				}
				else
				{
					parent.tmpCheckedList="";
					parent.updGWPrdList();
				}
			}
		});
		return(false);	
	});
	
	function getGWPrdList()
	{
		var tmpdata="";
		var i=0;
		for (var i=1;i<=PrdCanGW;i++)
		{
			if (eval("document.GWForm.PrdGW"+i).checked == true)
			{
				if (tmpdata!="") tmpdata=tmpdata + ",";
				tmpdata=tmpdata + eval("document.GWForm.PrdGW"+i).value;
			}
		}
		return(tmpdata);
	}
	
	//*Validate Payment Form
	var validator4=$("#PayForm").validate({
		rules: {
			cardNumber:
			{
				required: true,
				credit_card: true,
				remote: {
    				    url: "opc_checkCC.asp",
			    	    type: "POST",
			        	data: {
								cardType: function() {
						            return $("#cardType").val();
								}
				          }
	        	}
			}
		},
		messages: {
			chkPayment: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_43"))%>"
			},
			cardNumber:
			{
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_44"))%>",
				minlength: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_45"))%>",
				remote: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_46"))%>"
			}
		}
	});
	
	//*Validate Password Form
	var validator5=$("#PwdForm").validate({
		rules: {
			newPass1: 
			{
				required: true
			},
			newPass2:
			{
				required: true,
				equalTo: "#newPass1"
			}
		},
		messages: {
			newPass1: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_4"))%>"
			},
			newPass2: {
				required: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_47"))%>",
				equalTo: "<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_48"))%>"
			}
		}
	})
	
	//*Submit Password
	$('#PwdSubmit').click(function(){
		<%if Not (Session("idCustomer")>0 AND session("CustomerGuest")="0") then
		if scGuestCheckoutOpt="1" then%>
		if ((AskEnterPass==0) && (AcceptEnterPass==0) && ($("#newPass1").val()==""))
		{
			AcceptEnterPass=0;
			$("#AskEnterPassDialog").dialog('open');
			AskEnterPass=1;
			return(false);	
		}
		<%end if
		end if%>						   
		
		if ($('#PwdForm').validate().form())
		{
			$.ajax({
				type: "POST",
				url: "opc_createacc.asp",
				data: $('#PwdForm').formSerialize() + "&action=create",
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
					if ((data=="OK") || (data=="REG") || (data=="OKA") || (data=="REGA"))
					{

						$("#PwdLoader").hide();	
						$("#PwdWarning").hide();						
						$("#PwdArea").hide();

					}
					else
					{
						$("#PwdLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> '+data);
						$("#PwdLoader").show();
						btnShow1("Error","PW");
						validator5.resetForm();
					}
					}
				}
	 		});
			return(false);
		}
		return(false);
	});
	
	//*Submit Agree Terms Area
	function updTerms(tmpvalue)
	{			
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

		$.ajax({
			type: "POST",
			url: "opc_updAgree.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data=="OK")
				{
					$("#BillingLoader").hide();

				}
				else
				{
					$("#BillingLoader").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_53"))%>');
					$("BillingLoader").show();
				}
			}
		});
	}

	
	//'SB S
	function updTermsSB(tmpvalue)
	{			
		
		$.ajax({
			type: "GET",
			url: "opc_cartcheck.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data!="OK")
				{
					location="viewcart.asp";
				}
			}
		});

		$.ajax({
			type: "POST",
			url: "sb_updAgree.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data=="OK")
				{
					$("#BillingLoaderSB").hide();
				}
				else
				{
					$("#BillingLoaderSB").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_53"))%>');
					$("BillingLoaderSB").show();
				}
			}
		});
	}
	//'SB E

	$('#ViewTerms').click(function(){
		$("#TermsDialog").dialog('open');
		return ;
	});

	// 'SB S
	$('#sb_ViewTerms').click(function(){
		$("#sb_TermsDialog").dialog('open');
		return ;
	});
	// 'SB E

});

//*Global Javascript Functions

	//*Submit Gift Wrapping Link
	function GWAdd(pid, index) {
			$.ajax({
				type: "GET",
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

		var tmpdata=pid;	
		$.ajax({
			type: "POST",
			url: "opc_getGiftWrap.asp?list=" + tmpdata,
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data=="LOAD")
				{
					document.getElementById("GWframe").src="opc_giftwrap.asp?list=" + tmpdata + "&index=" + index;
					$("#GWDialog").dialog("open");
				}
				else
				{
					parent.tmpCheckedList="";
				}
			}
		});
		return(false);	
	}


	//* Fill information to Shipping Form
	function FillShipForm(tmpvalue,cback)
	{
		if ((tmpvalue=="-1") || (tmpvalue=="") || (tmpvalue=="-2"))
		{
			$("#shippingAddressArea").hide();
			if (tmpvalue=="-2")
			{
				if (HaveShipTypeArea==1) $("#shipAddrTypeArea").hide();
			}
		}
		else
		{
		if (tmpvalue=="ADD")
		{
			$("#shipnickname").val("");
			$("#shipfname").val("");
			$("#shiplname").val("");
			$("#shipcompany").val("");
			$("#shipaddr").val("");
			$("#shipaddr2").val("");
			$("#shipcity").val("");
			$("#shipzip").val("");
			$("#shipprovince").val("");
			$("#shipstate").val("");
			$("#shipcountry").val($("#billcountry").val());
			$("#shipphone").val("");
			$("#shipfax").val("");
			$("#shipemail").val("");
			//SwitchStates('ShippingForm',document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', $("#billstate").val(), '');
			if (HaveShipTypeArea==1) {document.ShippingForm.pcAddressType[0].checked=true};
			$("#shipnicknameArea").show();
			$("#shipnameArea").show("");
			$("#shippingAddressArea").show();
		}
		else
		{
			if (ShipContents!="")
			{
				var tmpShipList=ShipContents.split("|$|");
				if (tmpShipList.length>0)
				{
					for (var i=0;i<tmpShipList.length;i++)
					{
						if (tmpShipList[i]!="")
						{
							tmpShipRe=tmpShipList[i].split("|*|");
							if (tmpShipRe[0]==tmpvalue)
							{
								$("#shipnickname").val(tmpShipRe[1]);
								$("#shipfname").val(tmpShipRe[2]);
								$("#shiplname").val(tmpShipRe[3]);
								$("#shipemail").val(tmpShipRe[4]);
								$("#shipphone").val(tmpShipRe[5]);
								$("#shipfax").val(tmpShipRe[6]);
								$("#shipcompany").val(tmpShipRe[7]);
								$("#shipaddr").val(tmpShipRe[8]);
								$("#shipaddr2").val(tmpShipRe[9]);
								$("#shipcity").val(tmpShipRe[10]);
								$("#shipprovince").val(tmpShipRe[11]);
								$("#shipstate").val(tmpShipRe[12]);
								$("#shipzip").val(tmpShipRe[13]);
								$("#shipcountry").val(tmpShipRe[14]);
								SwitchStates('ShippingForm',document.ShippingForm.shipcountry.options.selectedIndex, 'shipcountry', 'shipstate', 'shipprovince', tmpShipRe[12], '');
								if (tmpShipRe[15]=="") tmpShipRe[15]="1";
								if (HaveShipTypeArea==1)
								{
									if (tmpShipRe[15]=="0")
									{document.ShippingForm.pcAddressType[1].checked=true}
									else
									{document.ShippingForm.pcAddressType[0].checked=true}
								}
								if (tmpShipRe[0]==0)
								{
									$("#shipnicknameArea").hide();
									$("#shipnameArea").hide("");
								}
								else
								{
									$("#shipnicknameArea").show();
									$("#shipnameArea").show("");
								}
								$("#shippingAddressArea").show();
							}
						}
					}
				}
			}
		}
		}
		if (tmpvalue=="") 
		{
			if (HaveShipTypeArea==1) $("#shipAddrTypeArea").hide();
			if (HaveDeliveryArea==1) $("#shipDeliveryArea").hide();
		}
		else
		{
			if ((tmpvalue=="-1"))
			{
				if (cback!=1)
				if (HaveShipTypeArea==1) {document.ShippingForm.pcAddressType[0].checked=true};
			}
			if (tmpvalue!="-2")
			{
				if (HaveShipTypeArea==1) $("#shipAddrTypeArea").show();
			}
			if (HaveDeliveryArea==1) $("#shipDeliveryArea").show();
		}
	}
	
	//*Prepare Contents of Shipping Charge Form
	function getShipChargeContents(tmpchkFree)
	{
		$("#ShipChargeLoadContentMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"><%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_54"))%>');
		$("#ShipChargeLoadContentMsg").show();
		$('#ShippingChargeArea').hide();
		var tmpdata="";
		if (tmpchkFree!="") tmpdata="pSubTotalCheckFreeShipping=" + tmpchkFree;
			$.ajax({
			type: "GET",
			url: "opc_cartcheck.asp",
			data: "{}",
			timeout: 45000,
			success: function(data, textStatus){
				if (data!="OK")
				{
					location="viewcart.asp";
				}
			}
		});

		$.ajax({
				type: "POST",
				url: "opc_chooseShpmnt.asp",
				data: tmpdata + '&{}',
				timeout: 120000,
				error: function (XMLHttpRequest, textStatus, errorThrown) {
					if (textStatus=='timeout') {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_72"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					} else {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_73"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					}
				},
				global: false,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						$("#ShipLoadContentMsg").hide(); // Hide Msg after results are res
						var tmpArr=data.split("|*|")
						if (tmpArr[0]=="STOP")
						{
							acc1.openPanel('opcShipping');
							$("#ShipChargeLoadContentMsg").html(tmpArr[1]);
							$("#ShipChargeLoadContentMsg").show();
							return ;
						}
						if (tmpArr[0]=="OK")
						{
							
							<%
							'// NO SHIPPING
							'   Shipping charges are not required.  
							'	Calculate Tax and redirect to Payment panel
							%>
							$('#ShippingChargeArea').html(tmpArr[1]);
							getTaxContents();                                        
							btnShow1("OK","Ship");  
							
						}
						else
						{
							
							<%
							'// SHIPPING
							'   Shipping charges are required.  
							'	Open the Shipping panel and display the selections
							%>	
							$('#ShippingChargeArea').html(data);
							acc1.openPanel('opcShipping');
							GoToAnchor('opcShippingAnchor');
							
						}
						
						btnShow1("OK","CO"); // Mark Billing Complete
						$('#ShippingChargeArea').show();
						$("#ShipChargeLoadContentMsg").hide();
						
						<% If pcv_strPayPanel = "1" Then %>
						acc1.openPanel('opcPayment');
						GoToAnchor('opcPaymentAnchor');
						$("#PleaseWaitDialog").dialog('close');	
						<% End If %>
						
					}
				}
	 		});
	}
	
	//* Get Shipping Method
	function getShipMethod(tmpvalue)
	{

				$.ajax({
				type: "GET",
				async: false,
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

		$.ajax({
				type: "POST",
				url: "opc_getshipselection.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if (data=="ERROR")
					{
						$("#ShippingMethod").html('<img src="images/pcv4_st_icon_error_small.png" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_57"))%>');
						$("#ShippingMethod").show();
						return(false);
					}
					else
					{

						$("#ShippingMethod").html("" + data);
						$("#ShippingMethod").show();
						
					}
				}
	 		});
	}


	//*Save Incomplete Order
	function saveOrd()
	{
		$.ajax({
				type: "POST",
				url: "SaveOrd.asp?opc=true",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						acc1.openPanel('opcPayment');
						GoToAnchor('opcPaymentAnchor');
						CheckPlaceOrder();
					}
				}
	 		});
	}
	
	
	<%
	'//////////////////////////////////////////////////////
	'// START: CALCULATE TAX
	'   This function is called when shipping methods are loaded 
	'	and when shipping methods are submit (opc_chooseShpmnt.asp).
	'	It calculates taxes and creates a new order preview.
	'//////////////////////////////////////////////////////
	%>
	function getTaxContents()
	{
		$('#TaxContentArea').html("");
		$('#TaxContentArea').hide();
		$.ajax({
				type: "POST",
				url: "opc_tax.asp",
				data: "{}",
				timeout: 45000,
				error: function (XMLHttpRequest, textStatus, errorThrown) {
					if (textStatus=='timeout') {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_70"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					} else {
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_71"))%>');
						$("#GlobalErrorDialog").dialog('open');
						return false;
					}
				},
				global: false,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else if ((data=="ZIPLENGTH"))					
					{	
						$("#GlobalErrorMsg").text('<%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_74"))%>');
						$("#GlobalErrorDialog").dialog('open');				
					}
					else
					{
						if ((data=="OK") || (data=="OKA"))
						{
							
							<%
							'// TAX APPLIED
							'   Generate new Order Preview
							'	Open Payment Panel
							%>
							GetOrderInfo("","#TaxLoadContentMsg",1,'Y');
							
						}
						else
						{
							
							<%
							'// TAX NOT APPLIED
							'   Mulptiple zip codes. Must select an option
							%>
							acc1.openPanel('opcPayment');
							GoToAnchor('opcPaymentAnchor');
							$('#TaxContentArea').html(data);
							$('#TaxContentArea').show();
							$("#PaymentContentArea").hide();
							$("#TaxLoadContentMsg").hide();
							$("#ShippingChargeArea").hide();
							return;
						}
						
					}
				}
	 		});
	}
	<%
	'//////////////////////////////////////////////////////
	'// END: CALCULATE TAX
	'//////////////////////////////////////////////////////
	%>
	
	// 'SB S - Agreement Functions
	function closeRegDialog() {
		$("#RegTermsDialog").dialog("close");
	}
	// 'SB E

	//*Gift Wrapping Functions
	function closeGWDialog() {
		$("#GWDialog").dialog("close");
	}
	
	function updGWPrdList()
	{
		if (tmpCheckedList!="")
		{
			var tmpArr=tmpCheckedList.split(",")
			var i=0;
			for (var i=1;i<=PrdCanGW;i++)
			{
				eval("document.GWForm.PrdGW"+i).checked=false;
				for (var j=0;j<tmpArr.length;j++)
				{
					if (tmpArr[j]!="")
					{
						if (eval("document.GWForm.PrdGW"+i).value==tmpArr[j])
						{
							eval("document.GWForm.PrdGW"+i).checked=true;
						}
					}
				}
			}
			
		}
		else
		{
			var i=0;
			for (var i=1;i<=PrdCanGW;i++)
			{
				eval("document.GWForm.PrdGW"+i).checked=false;
			}
		}
	}
	
	//*Prepare Contents of Pay Details Form
	function getPayDetails(tmpid,tmpURL)
	{
			$.ajax({
				type: "GET",
				async: false,
				url: "opc_cartcheck.asp",
				data: "{}",
				timeout: 45000,
				success: function(data, textStatus){
					if (data!="OK")
					{
						location="viewcart.asp";
					}
				}
	 		});

		$('#PayFormArea').html("");
		$('#PayFormArea').hide();
		$.ajax({
				type: "POST",
				url: tmpURL,
				data: "idpayment=" + tmpid,
				timeout: 45000,
				success: function(data, textStatus){
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						$('#PayFormArea').html(data);
						$('#PayFormArea').show();
						$("#PayLoader").hide();
						$('.chkPay').attr('disabled',false);
					}
				}
 		});
	}
	function GenOrderPreview(data,ctype)
	{
						if (data != "") {	
							var tmpArr=data.split("|***|")
							$("#DiscountCode").val(tmpArr[2]);
							$("#UseRewards").val(tmpArr[4]);
							$("#pcOPCtotalAmount").text(tmpArr[7]);
							$("#OPRArea").html(tmpArr[0]);
							$("#OPRWarning").hide();
							$("#OPRArea").show();
							if (ctype==1)
							{
								if (tmpArr[8]=="YES")
								{
									if (tmpArr[5]=="FREE")
									{
										CustomPayment=1;
										NeedToUpdatePay=0;
										$("#PayNoNeed").show();
										$("#PayAreaSub").hide();
										OPCFree=1;
									} else {
										PreSelectPayType(tmpArr[6]);
										$("#PayNoNeed").hide();
										$("#PayAreaSub").show();
										OPCFree=0;
									}
								}
							}
							OPCReady=tmpArr[8];		
							tmpchkFree=tmpArr[9];
							if (tmpArr[8]=="NO")
							{
								$("#OPRWarning").hide();
								$("#OPRArea").show();
								
								getShipChargeContents(tmpArr[9]);
								acc1.openPanel('opcShipping');
								GoToAnchor('opcShippingAnchor');
							}

							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
							}
							// Free Order - End			

							CheckPlaceOrder();
						}
		
	}
	
	//Get Order Details
	function GetOrderInfo(tmpid,dloader,ctype,tmpsaveOrd)
	{
		var tmpdata="";
		if (tmpid!="") tmpdata="idpayment=" + tmpid;
		if (dloader=="#TaxLoadContentMsg")
		{
			if ($("#DiscountCode")!="")
			{
				if (tmpdata!="") tmpdata=tmpdata + "&";
				tmpdata=tmpdata + "DiscountCode=" + $("#DiscountCode").val() + "&rtype=1";
			}
		}
		$.ajax({
				type: "POST",
				url: "opc_OrderVerify.asp",
				data: tmpdata + '&{}',
				timeout: 45000,
				success: function(data, textStatus){
					if (dloader=="#PayLoader1") $('.chkPay').attr('disabled',false);
					if (data=="SECURITY")
					{
						// Session Expired
						window.location="msg.asp?message=1";
					}
					else
					{
						if (data != "") {	
							var tmpArr=data.split("|***|")
							$(dloader).hide();
							$("#DiscountCode").val(tmpArr[2]);
							$("#UseRewards").val(tmpArr[4]);
							$("#pcOPCtotalAmount").text(tmpArr[7]);
							$("#OPRArea").html(tmpArr[0]);
							$("#OPRWarning").hide();
							$("#OPRArea").show();
							if (ctype==1)
							{
								if (tmpArr[8]=="YES")
								{
									if (tmpArr[5]=="FREE")
									{
										CustomPayment=1;
										NeedToUpdatePay=0;
										$("#PayNoNeed").show();
										$("#PayAreaSub").hide();
										OPCFree=1;
									}
									else
									{
										PreSelectPayType(tmpArr[6]);
										$("#PayNoNeed").hide();
										$("#PayAreaSub").show();
										OPCFree=0;
									}
								}
							}
							OPCReady=tmpArr[8];		
							tmpchkFree=tmpArr[9];
							if (tmpArr[8]=="NO")
							{
								$("#OPRWarning").hide();
								$("#OPRArea").show();
								
								getShipChargeContents(tmpArr[9]);
								acc1.openPanel('opcShipping');
								GoToAnchor('opcShippingAnchor');
							}
							
							// Free Order - Start
							if (tmpArr[8]=="YES")
							{
								if (tmpArr[5]=="FREE")
								{
									CustomPayment=1;
									NeedToUpdatePay=0;
									$("#PayNoNeed").show();
									$("#PayAreaSub").hide();
									OPCFree=1;
								} else {
									$("#PayNoNeed").hide();
									$("#PayAreaSub").show();
									OPCFree=0;
								}
							}
							// Free Order - End	
							
							if (tmpArr[8]!="NO")
							{
								// 'SB S
								$.post("opc_OrderVerify.asp", { sbTax: "1" } );
								// 'SB E
								
								if (tmpsaveOrd=='Y') {
									saveOrd();
								} else {
									CheckPlaceOrder();
								}
							}
							
						}
					}
				}
	 		});
	}
	
	//Check to show or hide Place Order buttons
	function CheckPlaceOrder()
	{
		OPCFinal = 0;
		if ((OPCReady=="") || (OPCReady=="NO"))
		{
			//$("#ButtonArea").hide();
			//$("#PlaceOrderButton").hide();
			//$("#ContinueButton").hide();
		}
		else
		{

			if ((CustomPayment==1) && (NeedToUpdatePay==1))
			{

				if (CustomPayment==1)
				{
					$("#PlaceOrderButton").show();
					$("#ContinueButton").hide();
				}
				else
				{
					$("#PlaceOrderButton").hide();
					$("#ContinueButton").show();
				}
				OPCFinal = 0;
				$("#ButtonArea").show();
			}
			else
			{

				if (CustomPayment==1)
				{
					$("#PlaceOrderButton").show();
					$("#ContinueButton").hide();
				}
				else
				{
					$("#PlaceOrderButton").hide();
					$("#ContinueButton").show();
				}
				OPCFinal = 1;
				$("#ButtonArea").show();

			}
		}
	}
	
	//* Submit All Payment Features
	function ValidateGroup1(tmpID)
	{	

		<% if session("ExpressCheckoutPayment") <> "YES" then %>
			if ($("#PaySubmit").length > 0) { 
				if (OPCFree==0) {
					try {
						$('#PaySubmit').click();
					} catch(err) {
						ValidateGroup2();
					}
				} else {
					ValidateGroup2();
				}
			} else {
				ValidateGroup2();
			}			
		<% else %>
			ValidateGroup2();
		<% end if %>
   	}
	function ValidateGroup2(tmpID)
	{	
	
		<%if Not (Session("idCustomer")>0 AND session("CustomerGuest")="0") then
		if scGuestCheckoutOpt="1" then%>
		if ((AskEnterPass==0) && (AcceptEnterPass==0) && ($("#newPass1").val()==""))
		{
			AcceptEnterPass=0;
			$("#AskEnterPassDialog").dialog('open');
			AskEnterPass=1;
			return(false);	
		}
		<%end if
		end if%>
	
		$("#PleaseWaitMsg").html('<img src="images/ajax-loader1.gif" width="20" height="20" align="absmiddle"> <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_81"))%>');
		$("#PleaseWaitDialog").dialog('open');	
		$(".ui-dialog-titlebar").css({'display' : 'none'});
		$("#PleaseWaitDialog").css({'min-height' : '50px'});
		if ($("#DiscSubmit").length > 0) {
			$('#DiscSubmit').click();
		} else {
			ValidateGroup3();
		}
		
   	}
	function ValidateGroup3(tmpID)
	{	
		if ($("#OtherSubmit").length > 0) {
			$('#OtherSubmit').click();
		} else {
			ValidateGroup4();
		}
   	}
	function ValidateGroup4(tmpID)
	{	
		CheckPlaceOrder();
		if (OPCFinal == 1) 
		{
			setTimeout('window.location="SaveOrd.asp"',2000);
		}
		else
		{
			$("#PleaseWaitDialog").dialog('close');
		}
   	}
	
	function pcf_LoadPaymentPanel()
	{	
		$('#LoadPaymentPanel').click();
   	}

	function RunE(tmpID)
	{
		var LoaderFade = new Spry.Effect.Fade(tmpID, {duration: 3000, from: 100, to: 0, toggle:true, finish:finishFunc});
		LoaderFade.start();
	}

	function finishFunc(elmentObj , FadeObj)
   	{
	   elmentObj.style.display="none";
   	}
	function GoToAnchor(a) {
		
		<% if NOT instr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE")>0 then %>
			  //document.getElementsByName(a)[0].focus()
			  window.location.hash=a;
			  //window.location.href = window.location.href.replace(/#[\S]+/,"") + '#'+a;
			  //document.getElementsByName(a)[0].scrollIntoView(true);	
		<% end if %>
	}

</script>