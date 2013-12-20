<script language="JavaScript">
<!--
function validateNumber(field)
{
  var val=field.value;
  if(!/^\d*$/.test(val)||val==0)
  {
      alert("<%response.write dictLanguage.Item(Session("language")&"_showcart_2")%>");
      field.focus();
      field.select();
  }
}


	function isDigit(s)
	{
	var test=""+s;
	if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
	{
	return(true) ;
	}
	return(false);
	}

	function allDigit(s)
	{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (isDigit(c)==false)
	{
	return (false);
	}
	}
	return (true);
	}

	function checkproqty(fname)
	{
	if (fname.value == "")
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_4")%>");
		fname.focus();
		return (false);
		}
	if (allDigit(fname.value) == false)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
		fname.focus();
		return (false);
	}
	if (fname.value == "0")
	{
		alert("<%=dictLanguage.Item(Session("language")&"_alert_5")%>");
		fname.focus();
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
		fname.focus();
		return (false);
	}
	<%end if%>
	<%if ((pNoStock="0") OR (IsNull(pNoStock))) AND ((pcv_intBackOrder="0") OR (IsNull(pcv_intBackOrder))) AND (scOutOfStockPurchase = -1) then%>
	TempValue=eval(fname.value);
	if (TempValue > <%=pStock%>)
	{
		alert("<%=dictLanguage.Item(Session("language")&"_instPrd_2")%><%=replace(pDescription,"""","\""")%><%=dictLanguage.Item(Session("language")&"_instPrd_3")%><%=pStock%><%=dictLanguage.Item(Session("language")&"_instPrd_4")%>");
		fname.focus();
		return (false);
	}
	<%end if%>
	<%if clng(pcv_lngMinimumQty)>0 then%>
	TempValue=eval(fname.value);
	if (TempValue < <%=pcv_lngMinimumQty%>)
	{
		alert("<% response.write(dictLanguage.Item(Session("language")&"_alert_8") & clng(pcv_lngMinimumQty) & dictLanguage.Item(Session("language")&"_alert_9"))%>");
		fname.focus();
		return (false);
	}
	<%end if%>
	return (true);
	}
	

function optwin(fileName)
	{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=300')
	myFloater.location.href = fileName;
	}
//-->
</script>