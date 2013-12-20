<script language="JavaScript">
<!--
var imagename = '';

function enlrge(imgnme) {
	lrgewin=window.open("about:blank","","height=200,width=200,status=no,titlebar=yes")
	imagename=imgnme;
	setTimeout('update()',500)
}

function win(fileName)
{
	myFloater = window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
	myFloater.location.href = fileName;
}

function viewWin(file)
{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=<%=iBTOPopWidth%>,height=<%=iBTOPopHeight%>')
	myFloater.location.href = file;
}

function update() {
	<%
	'**** Check MAC IE Browser ******************
	UserBrowser=Request.ServerVariables("HTTP_USER_AGENT")
	if UserBrowser<>"" then
		MACBrowser=instr(ucase(UserBrowser),"MAC")
	end if
	'**** End of Check MAC IE Browser ******************
	%>
	doc=lrgewin.document;
	doc.open('text/html');
	doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if  (self.resizeTo)self.resizeTo((document.images[0].width+60),(document.images[0].height+150)); return false;" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table width=""' + document.images[0].width + '" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>');
	doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><%if MACBrowser=0 then%><form name="viewn"><input type="image" src="images/close.gif" align="right" value="Close Window" onClick="self.close(); return false;"><%end if%><\/td><\/tr><\/table>');
	doc.write('<\/form><\/BODY><\/HTML>');
	doc.close();
}

function besubmit()
{
	document.getElementById("show_1").style.display = '';;
	return (false);
}
//-->
</script>