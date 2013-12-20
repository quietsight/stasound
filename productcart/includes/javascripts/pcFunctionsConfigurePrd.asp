<script language="JavaScript">
<!--

var stralert1="<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n";
var stralert2="<%=dictLanguage.Item(Session("language")&"_alert_2")%>.\n";

function qttverify(fname,QtyValidate,MinQty,MultiQty,CheckStock,FromPQField)
{
	if (fname.value == "")
	{
	    alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
	    setTimeout(function() {fname.focus();}, 0);
	    return (false);
    }
	if (allDigit(fname.value) == false)
	{
	    alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
	    setTimeout(function() {fname.focus();}, 0);
	    return (false);
	}
	if (fname.value>"0")
	{
		if (QtyValidate=="1")
		{
			TempValue=eval(fname.value);
			TempV1=(TempValue/MultiQty);
			TempV1a=TempValue*TempV1;
			TempV2=Math.round(TempValue/MultiQty);
			TempV2a=TempValue*TempV2;
			if ((TempV1a != TempV2a) || (TempV1<1))
			{
			    alert("<% response.write(dictLanguage.Item(Session("language")&"_alert_3"))%>"+" "+ MultiQty);
			    setTimeout(function() {fname.focus();}, 0);
			    return (false);
			}
		}
		if (MinQty>0)
		{
			TempValue=eval(fname.value);
			if (TempValue < MinQty)
			{
				alert("<% response.write(dictLanguage.Item(Session("language")&"_alert_8"))%>" + " " +MinQty+" " +"<%=dictLanguage.Item(Session("language")&"_alert_9")%>");
				setTimeout(function() {fname.focus();}, 0);
				return (false);
			}
		}
		if (undefined != CheckStock)
		{
			TempValue=eval(fname.value)*eval(document.getElementById('quantity').value);
			if (TempValue > CheckStock)
			{
				alert("<%=dictLanguage.Item(Session("language")&"_instConfQty_2a")%>" + TempValue + "<%=dictLanguage.Item(Session("language")&"_instConfQty_2b")%>" + CheckStock + "<%=dictLanguage.Item(Session("language")&"_instConfQty_2c")%>");
				if (undefined == FromPQField) setTimeout(function() {fname.focus();}, 0);
				return (false);
			}
		}
	}
	return (true);
}

function checkproqty(fname)
{
	<%if strQtyCheck<>"" then%>
	<%=strQtyCheck%>
	<%end if%>
	
	if (fname.value == "")
	{
	    alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
	    setTimeout(function() {fname.focus();}, 0);
	    return (false);
    }
	if (allDigit(fname.value) == false)
	{
	    alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
	    setTimeout(function() {fname.focus();}, 0);
	    return (false);
	}
	if (fname.value == "0")
	{
	    alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
	    setTimeout(function() {fname.focus();}, 0);
	    return (false);
	}
	<%if pcv_intQtyValidate="1" then%>
		TempValue=eval(fname.value);
		TempV1=(TempValue/<%=clng(pcv_lngMultiQty)%>);
		TempV1a=TempValue*TempV1;
		TempV2=Math.round(TempValue/<%=clng(pcv_lngMultiQty)%>);
		TempV2a=TempValue*TempV2;
		if ((TempV1a != TempV2a) || (TempV1<1))
		{
		    alert("<% response.write(dictLanguage.Item(Session("language")&"_alert_3") & clng(pcv_lngMultiQty))%>");
		    setTimeout(function() {fname.focus();}, 0);
		    return (false);
		}
	<%end if
	if clng(pcv_lngMinimumQty)>0 then%>
		TempValue=eval(fname.value);
		if (TempValue < <%=pcv_lngMinimumQty%>)
		{
			alert("<% response.write(dictLanguage.Item(Session("language")&"_alert_8") & clng(pcv_lngMinimumQty) & dictLanguage.Item(Session("language")&"_alert_9")) %>");
			setTimeout(function() {fname.focus();}, 0);
			return (false);
		}
	<%end if%>
	return (true);
}
				
function viewWin(file)
{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=<%=iBTOPopWidth%>,height=<%=iBTOPopHeight%>')
	myFloater.location.href = file;
}

function checkDropdown(choice, option)
{
	if (choice == 0)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
	}
	return true;
}

function checkDropdowns(choice1, choice2, option1, option2)
{
	if (choice1 == 0)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option1 + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
	}
	if (choice2 == 0)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option2 + "<%=dictLanguage.Item(Session("language")&"_alert_6")%>.\n");
		return false;
	}
	return true;
}
//-->
</script>