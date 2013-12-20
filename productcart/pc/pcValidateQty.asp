<script language="javascript">
/* Validate quantity field for product display option "m" */
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
	
	function checkproqty(fname,qty,type,MultiQty)
	{
	if (fname.value == "")
	{
		alert("<% Response.write(dictLanguage.Item(Session("language")&"_alert_4"))%>");
		setTimeout(function() {fname.focus();}, 0);
		return (false);
	}
	if (allDigit(fname.value) == false)
	{
		alert("<% Response.write(dictLanguage.Item(Session("language")&"_alert_5"))%>");
		setTimeout(function() {fname.focus();}, 0);
		return (false);
	}
	if ((MultiQty > 0) && (eval(fname.value) != 0) && (type==1))
	{
	TempValue=eval(fname.value);
	TempV1=(TempValue/MultiQty);
	TempV1a=TempValue*TempV1;
	TempV2=Math.round(TempValue/MultiQty);
	TempV2a=TempValue*TempV2;
		if ((TempV1a != TempV2a) || (TempV1<1))
		{
				alert("<% Response.write(dictLanguage.Item(Session("language")&"_alert_3"))%>" + MultiQty);
				setTimeout(function() {fname.focus();}, 0);
				return (false);
		}
	}
	if ((qty > 0) && (eval(fname.value) != 0))
	{
		TempValue=eval(fname.value);
		if (TempValue < qty)
		{
			alert("<%=dictLanguage.Item(Session("language")&"_alert_8")%>" + " " + qty + " " + "<%=dictLanguage.Item(Session("language")&"_alert_9")%>");
			setTimeout(function() {fname.focus();}, 0);
			return (false);
		}
	}
	return (true);
	}
	
	/*
	Clear default form value script- By Ada Shimar (ada@chalktv.com)
	Featured on JavaScript Kit (http://javascriptkit.com)
	Visit javascriptkit.com for 400+ free scripts!
	*/
	
	function clearText(thefield){
	if (thefield.defaultValue==thefield.value)
	thefield.value = ""
	} 
</script>