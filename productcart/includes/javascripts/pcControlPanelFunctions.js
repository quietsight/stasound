function checkTheBoxesB()
{
	with (document.myForm) {
		for(i=0; i < elements.length -1; i++) {
			if ( elements[i].name== "optionDescrip" ) {
				elements[i].checked=true;
			}
		}
	}
}	

function checkTheBoxesA()
{
	with (document.myForm) {
		for(i=0; i < elements.length -1; i++) {
			if ( elements[i].name== "product" ) {
				elements[i].checked=true;
			}
		}
	}
}	

function checkTheBoxes()
{
	with (document.myForm) {
		for(i=0; i < elements.length -1; i++) {
			if ( elements[i].name== "Address" ) {
				elements[i].checked=true;
			}
		}
	}
}	

function win(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=600,height=550')
	myFloater.location.href=fileName;
	}

function Graph(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=No,status=no,width=320,height=250')
	myFloater.location.href=fileName;
	}

if ((!(navigator.appVersion.indexOf('MSIE') != -1) && 
	(parseInt(navigator.appVersion)==4))) 
{
	document.write("<STYLE TYPE=\"text/css\">");
	document.write("BODY { margin-top: -8px; margin-left: -8px; }"); 
	document.write("</style>");
};

var bV=parseInt(navigator.appVersion);
NS4=(document.layers) ? true : false;
IE4=((document.all)&&(bV>=4))?true:false;
hasDOM=(document.getElementById) ? true : false;

if (NS4) {
    origWidth=innerWidth;
    origHeight=innerHeight;
}

function open_marketing_win (lang) {
    var url='preview/' + lang;
    var mWin=window.open(url, 'Welcome','menubar=no,toolbar=no,scrollbars=no,resizable=no,width=760,height=500');
    mWin.focus();
}

function open_pdf_win(url) {
    var pdfWin=window.open(url,'PDF','scrollbars=yes,resizable=yes');
    pdfWin.focus();
}

function reDo() {
    if (innerWidth != origWidth || innerHeight != origHeight) 
        location.reload();
}

if (NS4) onresize=reDo;

/*
Select and Copy form element script- By Dynamicdrive.com
For full source, Terms of service, and 100s DTHML scripts
Visit http://www.dynamicdrive.com
*/

//specify whether contents should be auto copied to clipboard (memory)
//Applies only to IE 4+
//0=no, 1=yes
var copytoclip=1

function HighlightAll(theField) {
var tempval=eval("document."+theField)
tempval.focus()
tempval.select()
if (document.all&&copytoclip==1){
therange=tempval.createTextRange()
therange.execCommand("Copy")
window.status="Contents highlighted and copied to clipboard!"
setTimeout("window.status=''",1800)
}
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

// Hide/Show inactive products of the category.
function reloadFormToShowHideInactiveProducts()
{ 
    document.myForm.frmHideInactiveProducts.value = document.myForm.chkHideInActiveProducts.checked;
	document.myForm.submit();
}

// Select all content of a text area
function selectFieldContent(id)

{
    document.getElementById(id).focus();
    document.getElementById(id).select();
}

// Count-down of remaining characters
function testchars(tmpfield,idx,maxlen)
{
	var tmp1=tmpfield.value;
	if (tmp1.length>maxlen)
	{
		alert("Maximum " + maxlen + " characters allowed.");
		tmp1=tmp1.substr(0,maxlen);
		tmpfield.value=tmp1;
		document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
		tmpfield.focus();
	}
	document.getElementById("countchar" + idx).innerHTML=maxlen-tmp1.length;
}