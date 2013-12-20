<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript" src="../htmleditor/innovaeditor.js"></script>
<style>
html, body {
	margin: 0;
	padding: 0;
}
</style>
</head>
<body onLoad="javascript:window.resizeTo(700,530);">
<%if request("iform")="" then
	pcv_strForm="hForm"
else
	pcv_strForm=request("iform")
end if%>
<script language="JavaScript">

	function LoadDoc()
	{
		<%if request("fi")="" then%>
		document.HTMLForm.elements.txtContent.value=opener.document.<%=pcv_strForm%>.elements.details.value;
		<%else%>
		abc=eval("opener.document.<%=pcv_strForm%>.elements." + "<%=request("fi")%>")
		document.HTMLForm.elements.txtContent.value=abc.value;
		<%end if%>
	}

	function Save()
	{
		document.HTMLForm.elements.txtContent.value=oEdit1.getHTMLBody();
		
		<%if request("fi")="" then%>
		opener.document.<%=pcv_strForm%>.elements.details.value = document.HTMLForm.elements.txtContent.value;
		<%else%>
		abc=eval("opener.document.<%=pcv_strForm%>.elements." + "<%=request("fi")%>")
		abc.value = document.HTMLForm.elements.txtContent.value;
		<%end if%>
		
		self.close();
	   	return false;
	}
</script>
<form name="HTMLForm" ID="HTMLForm">

<textarea id="txtContent" name="txtContent" rows=4 cols=30></textarea>

<script>
	LoadDoc();
	var oEdit1 = new InnovaEditor("oEdit1");

    oEdit1.width=680;
    oEdit1.height=450;

    /***************************************************
      RECONFIGURE TOOLBAR BUTTONS
    ***************************************************/

    /*Set toolbar mode: 0: standard, 1: tab toolbar, 2: group toolbar */
    oEdit1.toolbarMode = 1;

    oEdit1.tabs=[
    ["tabHome", "Home", ["grpEdit", "grpFont", "grpPara"]],
    ["tabStyle", "Insert", ["grpInsert", "grpTables"<%if session("admin")=-1 then%>, "grpMedia"<%end if%>]]
    ];

    oEdit1.groups=[
    ["grpEdit", "", ["XHTMLSource", "FullScreen", "Search", "RemoveFormat", "BRK", "Undo", "Redo", "Cut", "Copy", "Paste", "PasteWord", "PasteText"]],
    ["grpFont", "", ["FontName", "FontSize", "Strikethrough", "Superscript", "BRK", "Bold", "Italic", "Underline", "ForeColor", "BackColor"]],
    ["grpPara", "", ["Paragraph", "Indent", "Outdent", "Styles", "StyleAndFormatting", "BRK", "JustifyLeft", "JustifyCenter", "JustifyRight", "JustifyFull", "Numbering", "Bullets"]],
    ["grpInsert", "", ["Hyperlink", "Bookmark", "BRK", "Image"]],
    ["grpTables", "", ["Table", "BRK", "Guidelines", "AutoTable"]]
	<%if session("admin")=-1 then%>,
    ["grpMedia", "", ["Media", "Flash", "YoutubeVideo", "BRK", "CustomTag", "Characters", "Line"]]
	<%end if%>
    ];

    /***************************************************
      OTHER SETTINGS
    ***************************************************/
    oEdit1.css="../../style.css";//Specify external css file here. If Table Auto Format is enabled, the table autoformat css rules must be defined in the css file.

    <%if session("admin")=-1 then%>
    oEdit1.cmdAssetManager = "modalDialogShow('../htmleditor/assetmanager/assetmanager.asp',640,465)"; //Command to open the Asset Manager add-on.
    <%end if%>
    //oEdit1.cmdInternalLink = "modelessDialogShow('links.htm',365,270)"; //Command to open your custom link lookup page.
    //oEdit1.cmdCustomObject = "modelessDialogShow('objects.htm',365,270)"; //Command to open your custom content lookup page.

    oEdit1.arrCustomTag=[["First Name","{%first_name%}"],
        ["Last Name","{%last_name%}"],
        ["Email","{%email%}"]];//Define custom tag selection

    oEdit1.customColors=["#ff4500","#ffa500","#808000","#4682b4","#1e90ff","#9400d3","#ff1493","#a9a9a9"];//predefined custom colors

    oEdit1.mode="XHTMLBody"; //Editing mode. Possible values: "HTMLBody" (default), "XHTMLBody", "HTML", "XHTML"

    oEdit1.REPLACE("txtContent");
  </script>


	<div style="position: absolute; top: 3px; left: 620px;">
	<input type="button" value="Save" onClick="Save()" id="btnSave" name="btnSave">
    </div>
</form>
</body>
</html>