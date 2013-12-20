<script language="JavaScript">
<!--
		imagename='';
		pcv_jspIdProduct='';
		pcv_jspCurrentUrl='';
		
		
		function enlrge(imgnme) {
			lrgewin=window.open("about:blank","","height=200,width=200,status=no,titlebar=yes")
			imagename=imgnme;
			setTimeout('update()',500)
		}
		
		function pcAdditionalImages(jspCurrentUrl,jspIdProduct) {
			lrgewin=window.open("about:blank","","height=600,width=800,status=no,titlebar=yes")
			pcv_jspIdProduct = jspIdProduct;
			pcv_jspCurrentUrl = jspCurrentUrl;
			setTimeout('updateAdditionalImages()',500)
		}
		
		
		function win(fileName)
			{
			myFloater=window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
			myFloater.location.href=fileName;
			}
		
		
		function updateAdditionalImages() {
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
		doc.write('<HTML><HEAD><TITLE>Loading Image Viewer<\/TITLE>')
		doc.write('<link type="text/css" rel="stylesheet" href="pcStorefront.css" /><\/HEAD>')
		doc.write('<BODY bgcolor="white" topmargin="4" leftmargin="0" rightmargin="0" onload="document.viewn.submit();" bottommargin="0">')
		doc.write('<div id="pcMain">');
		doc.write('<form name="viewn" action="viewPrdPopWindow.asp?idProduct=' + pcv_jspIdProduct + '" method="post" class="pcForms">');
		doc.write('<table class="pcMainTable"><tr><td>');		
		doc.write('<input name="idProduct" type="hidden" value="' + pcv_jspIdProduct + '" \/>');
		doc.write('<input name="pcv_strCurrentUrl" type="hidden" value="' + pcv_jspCurrentUrl + '" \/>');		
		doc.write('<\/td><\/tr><tr><td align="center">' + "<%=dictLanguage.Item(Session("language")&"_PrdError_3")%>" + '<\/td><\/tr><\/table>');
		doc.write('</form>');
		doc.write('</div>');
		doc.write('<\/BODY><\/HTML>');
		doc.close();
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
		
//-->
</script>