<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#INCLUDE FILE="../includes/settings.asp"-->
<!--#INCLUDE FILE="../includes/storeconstants.asp"-->
<!--#INCLUDE FILE="../includes/opendb.asp"-->
<!--#INCLUDE FILE="../includes/languages.asp"-->
<!--#INCLUDE FILE="../includes/currencyformatinc.asp"-->
<!--#INCLUDE FILE="../includes/adovbs.inc"-->
<!--#INCLUDE FILE="../includes/stringfunctions.asp"-->
<!--#include FILE="../includes/ErrorHandler.asp"--> 
<!--#include FILE="../includes/SQLFormat.txt"--> 
<%
Response.Buffer = True

'-------------------------------
' declare local variables
'-------------------------------
Dim CurFile, PopFile,ShowSub, ShowPic, PictureNo, ref, pShowButton
Dim strPathInfo, strPhysicalPath, query, rs, conntemp, doIndex, imgName, lastIndex
Dim intTotPics, intPicsPerRow, intPicsPerPage, intTotPages, intPage, strPicArray()

fid=request.QueryString("fid")
ffid=request.QueryString("ffid")
imgIndex=request.QueryString("btnIndex")
ref=request.QueryString("ref")

'--- Get Search Form parameters ---
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_key4=getUserInput(request("key4"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)
pshowimage=getUserInput(request("showimage"),0)
precords=form_resultCnt
pshowform=request.QueryString("ajaxSearch")


doIndex=request.QueryString("doIndex")
if doIndex="" then
    doIndex=0
else
    doIndex=-1
end if

if pshowimage="YES" then
	intPicsPerRow  = 4
	intPicsPerPage = precords
else
	intPicsPerRow  = 4
	intPicsPerPage = precords
end if

intPage = CInt(Request.QueryString("Page"))
If intPage = 0 Then
	intPage = 1
End If

CurFile = "ImageDir.asp"
PopFile = "showPicture.asp"


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check if Index exists
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
call opendb()
query="SELECT pcImgDir_Name,pcImgDir_DateIndexed from pcImageDirectory"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rs=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rs.EOF then
    doIndex=-1
else
    lastIndex = rs("pcImgDir_DateIndexed")
end if
set rs=nothing
call closedb()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check if Index exists
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>
<html>
<head>
<title>Display images</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function openGalleryWindow(url) {
	if (document.all)
		var xMax = screen.width, yMax = screen.height;
	else
		if (document.layers)
			var xMax = window.outerWidth, yMax = window.outerHeight;
		else
			var xMax = 800, yMax=600;
	var xOffset = (xMax - 200)/2, yOffset = (yMax - 200)/2;
	var xOffset = 100, yOffset = 100;

	popupWin = window.open(url,'new_page','width=700,height=535,screenX='+xOffset+',screenY='+yOffset+',top='+yOffset+',left='+xOffset+',scrollbars=auto,toolbars=no,menubar=no,resizable=yes')
}

function DoSubmit()
{
    document.frmIndex.btnIndex.style.visibility="hidden";
    document.frmIndex.btnWait.style.display="block";
	return(true);
    //document.frmIndex.submit();
}

function OnLoad()
{
    window.resizeTo(600,600);
    
    //center window
    //var x = screen.width/2, y = screen.height/2;
    //window.moveTo(x-300,y-300);
}

// done hiding -->
</script>
</head>
<body onload="Javascript:OnLoad();">
<div id="pcMain" style="background-color: #FFF; color: #000;">
	<table class="pcMainTable">
		<tr>
			<td>
			<h1 style="color: #000;">Locate an Image</h1>
		<%

  	
    IF doIndex=-1 and imgIndex="" then
%>    
	    <form action="imageDir.asp" method="get" name="frmIndex" class="pcForms">
	        <input type="hidden" name="fid" value="<%=fid%>">
	        <input type="hidden" name="ffid" value="<%=ffid%>">
	        <input type="hidden" name="ref" value="<%=ref%>">
        	
	        <table class="pcShowContent" style="color: #000;">
		        <tr>
							<td>
								Searching for an image requires that you first index your image directory.<br />Click on the following button to create a searchable index of all the images contained in the &quot;catalog&quot; folder. Please note that this may take some time if you have a large amount of images in that folder.</p>
							</td>
		        </tr>
		        <tr>
			        <td>
			            <input type="submit" name="btnIndex" class="submit2" value="Index"  onclick="Javascript:DoSubmit();">
			            <input type="submit" name="btnWait" class="submit2" value="Please wait...." disabled style="display:none" />
			        </td>
		        </tr>
	        </table>
	    </form>
<%

   	elseIF imgIndex="Index" OR pshowform="submit"  THEN '1 - Check form submission
	
	    IF imgIndex="Index"  THEN '1 - Check form submission
            '=======================================
            ' START Index Files
            '=======================================
            strPhysicalPath = Server.MapPath(".\catalog")
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFSO.GetFolder(strPhysicalPath)
            Set objFolderContents = objFolder.Files

            Dim pTodayDate
            pTodayDate=Date()
            if SQL_Format="1" then
                pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
            else
                pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
            end if
            pTodayDate=pTodayDate&" "&Time()
            call opendb()
            query="DELETE from pcImageDirectory"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            set rs=nothing
            
            msg=""
            For Each objFileItem in objFolderContents
	            If Ucase(Right(objFileItem.Name,4))=".GIF" OR Ucase(Right(objFileItem.Name,4))=".JPG" OR Ucase(Right(objFileItem.Name,4))=".JPE" OR Ucase(Right(objFileItem.Name,5))=".JPEG" OR Ucase(Right(objFileItem.Name,4))=".PNG" THEN
	            	IF inStr(objFileItem.Name,"'")=0 AND inStr(objFileItem.Name,"%")=0 AND inStr(objFileItem.Name,",")=0 AND inStr(objFileItem.Name,"#")=0 AND inStr(objFileItem.Name,"+")=0 THEN
	            	    query="INSERT INTO pcImageDirectory (pcImgDir_Name,pcImgDir_Type,pcImgDir_Size,pcImgDir_DateUploaded,pcImgDir_DateIndexed) "
						if SQL_Format="1" then
							pDateCreated=Day(objFileItem.DateCreated)&"/"&Month(objFileItem.DateCreated)&"/"&Year(objFileItem.DateCreated)
						else
							pDateCreated=Month(objFileItem.DateCreated)&"/"&Day(objFileItem.DateCreated)&"/"&Year(objFileItem.DateCreated)
						end if
						pDateCreated=pDateCreated&" "&Time()	
                	    if scDB="Access" then
                	        query=query&"VALUES('"&objFileItem.Name&"','"&objFileItem.Type&"',"&objFileItem.Size&",#"&pDateCreated&"#,#"&pTodayDate&"#)"
                	    else
                	        query=query&"VALUES('"&objFileItem.Name&"','"&objFileItem.Type&"',"&objFileItem.Size&",'"&pDateCreated&"','"&pTodayDate&"')"
                	    end if
                	    set rs=server.CreateObject("ADODB.RecordSet")
                	    set rs=conntemp.execute(query)
                	    if err.number<>0 then
                	        '//Logs error to the database
                	        call LogErrorToDatabase()
                	        '//clear any objects
                	        set rs=nothing
                	        '//close any connections
                	        call closedb()
                	        '//redirect to error page
                	        response.redirect "techErr.asp?err="&pcStrCustRefID
                	    end if
                	    set rs=nothing
					ELSE
						msg="One or more images in the 'pc/catalog' folder could not be indexed because the file names contained an apostrophe ('), percent sign (%), comma (,), number sign (#), or plus sign (+). Please rename these images and click on the 'Index Now' button again."
					END IF
	            End if
            Next

            set rs=nothing
            call closedb()
            
            lastIndex = pTodayDate
            doIndex=0 

            '=======================================
            ' END Index Files
            '=======================================
        end if
        
    end if
    
    IF doIndex=0  THEN 
%>
        <form action="imageDir.asp" method="get" name="frmReIndex" class="pcForms">
            <input type="hidden" name="fid" value="<%=fid%>">
            <input type="hidden" name="ffid" value="<%=ffid%>">
            <input type="hidden" name="ref" value="<%=ref%>">

            <table class="pcShowContent" style="color: #000;">
            	<%if msg<>"" then%>
            	<tr>
            		<td>
            			<p>&nbsp;</p>
						<div class="pcErrorMessage"><%=msg%></div>
						<p>&nbsp;</p>
            		</td>
            	</tr>
            	<%end if%>
                <tr>
                    <td>
                       Where are the images that I just uploaded? If you don't see images that you recently uploaded, re-index the image directory using the &quot;Index now&quot; button below.
                    </td>
                </tr>    
                <tr>
                    <td>Last Index: <%=lastIndex %></td>
                </tr>
                <tr>
                    <td>
                        <input type="submit" name="doIndex" class="submit2" value="Index now">
                    </td>
                </tr>
								<tr>
									<td class="pcSpacer">&nbsp;</td>
								</tr>
            </table>
        </form>
		<%
			src_FormTitle1=""
			src_FormTitle2=""
			src_FormTips1="Use the following filters to look for images in your store."
			src_FormTips2=""
			src_IncNormal=1
			src_IncBTO=0
			src_IncItem=0
			src_DisplayType=0
			src_ShowLinks=0
			src_FromPage="imagedir.asp?ajaxSearch=submit&fid="&fid&"&ffid="&ffid&" "
			src_ToPage="imagedir.asp"
			src_Button1=" Search "
			src_Button2=" Continue "
			src_Button3=" Back "
			src_PageSize=""
			UseSpecial=1
			session("srcprd_from")=""
			session("srcprd_where")=""
		%>
        <!--#INCLUDE FILE="inc_SrcImgs.asp"-->        
	<% 

	END IF %>
	
	</td>
	</tr>
	</table>
</div>
</body>
</html>