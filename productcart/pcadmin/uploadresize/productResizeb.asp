<%@ LANGUAGE="VBSCRIPT" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../../includes/ppdstatus.inc"-->
<!--#include file="../../includes/productcartFolder.asp"-->
<!--#include file="checkImgUplResizeObjs.asp"-->
<!--#include file="../../includes/pcSanitizeUpload.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <title>Upload Images</title>
    <link href="../pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<% 'Check if objects exists
If HaveImgUplResizeObjs=0 then %>
    <table class="pcCPcontent">
        <tr>
            <td>
            <div class="pcCPmessage">We are unable to find compatible Upload and/or Image Resize server components. Please consult the User Guide for detailed system requirements.</div>
            </td>
        </tr>
        <tr>
            <td align="center"><input type="button" name="Close" value=" Close window " onClick="javascript:window.close();" class="ibtnGrey"></td>
        </tr>
    </table>
	<% Response.End 'kill page
End If 
'if NO objects, kill the page %>

<%
Dim catalogpath, uploadpath, thumbfilename, generalfilename, detailfilename, thumbnailsize, generalsize, detailsize, sharpen, countfiles
Dim randomnum, FileName, BigBeforeWidth, BigBeforeHeight, BigAfterWidth, BigAfterHeight, imgcomp
Dim Image
Dim resizexy

Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

if PPD="1" then
	catalogpath=Server.Mappath ("\"&scPcFolder&"\pc\catalog\")
else
	catalogpath=Server.Mappath ("..\..\pc\catalog\")
end if
catalogpath = catalogpath & "\"

if PPD="1" then
	uploadpath=Server.Mappath ("\"&scPcFolder&"\includes\uploadresize\")
else
	uploadpath=Server.Mappath ("..\..\includes\uploadresize\")
end if
uploadpath = uploadpath & "\"

'on error resume next

Function UseSAFileUp()

	'--- Instantiate the FileUp object
	Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	
	Upload.Path = uploadpath
	
	thumbnailsize = Upload.Form("thumbnailsize")
	generalsize = Upload.Form("generalsize")
	detailsize = Upload.Form("detailsize")
	sharpen = Upload.Form("sharpen")	
	resizexy = Upload.Form("resizexy")
	
	If NOT IsUploadAllowed(Upload.Form("file1").UserFilename) or Upload.Form("file1").UserFilename = "" Then 	%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
                  <a href="#" onClick="history.back();">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
		<% Response.End
	End If
	%>	
	
	<%
	If right(Upload.Form("file1").UserFilename,3) <> "jpg" and right(Upload.Form("file1").UserFilename,3) <> "jpe" and right(Upload.Form("file1").UserFilename,4) <> "jpeg" Then
		If (pcv_ResizeObj=2) or ((pcv_ResizeObj=1) and (right(Upload.Form("file1").UserFilename,3)<>"gif")) then %>
            <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
            <tr> 
              <td> 
                <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
                  <tr> 
                    <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
                  </tr>
                  <tr> 
                    <td height="10"></td>
                  </tr>
                  <tr> 
                    <td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
                        <%If (pcv_ResizeObj=2) then%>
                            <b>This is not a JPG image.<br>GIF files are not supported with the ASPImage Component.</b>
                        <%Else%>
                            <b>This is not a JPG/GIF image.<br>Only JPG/GIF files are supported with the ASPJpeg Component.</b>
                        <%End if%>
                        <br><br>
                      <a href="#" onClick="history.back();">Click Here to go Back</a>
                      </font></td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td>&nbsp;</td>
            </tr>
            </table>
			<% Upload.Delete
			Set Upload = nothing
			Response.End
		End if
	End If

	Upload.Form("FILE1").Save

	FileName = Mid(Upload.Form("file1").UserFilename, InstrRev(Upload.Form("file1").UserFilename, "\") + 1)
	FileName = lcase(Filename)

	If (pcv_ResizeObj=1) then
		call UseASPJpeg(FileName,uploadpath & Filename)
	else
		call UseAspImage(FileName,uploadpath & Filename)
	end if

	Upload.Delete
	Set Upload = nothing

End Function

Function UseASPUpload()

	Set Upload = Server.CreateObject("Persits.Upload")
	
	If PPD="1" then
		Upload.SaveVirtual "\"&scPcFolder&"\includes\uploadresize\"
	else
		Upload.SaveVirtual "..\..\includes\uploadresize\"
	end if
	
	thumbnailsize = Upload.Form("thumbnailsize")
	generalsize = Upload.Form("generalsize")
	detailsize = Upload.Form("detailsize")
	sharpen = Upload.Form("sharpen")
	resizexy = Upload.Form("resizexy")
	
	countfiles = 0
	For Each File in Upload.Files
		countfiles = countfiles + 1
	Next
	
	'Count files in upload.  If none exist, exit script
	If countfiles = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
                  <a href="#" onClick="history.back();">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
        <% Response.End
	End If
	
	'Run the resizer 
	For Each File in Upload.Files
	
		If File.ImageType <> "JPG" then
			If NOT IsUploadAllowed(File.FileName) OR (pcv_ResizeObj=2) or ((pcv_ResizeObj=1) and (File.ImageType <> "GIF")) then%>
				<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
				<tr> 
				  <td> 
					<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
					  <tr> 
						<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
					  </tr>
					  <tr> 
						<td height="10"></td>
					  </tr>
					  <tr> 
						<td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
							<%If (pcv_ResizeObj=2) then%>
								<b>This is not a JPG image.<br>GIF files are not supported with the ASPImage Component.</b>
							<%Else%>
								<b>This is not a JPG/GIF image.<br>Only JPG/GIF files are supported with the ASPJpeg Component.</b>
							<%End if%>
							<br><br>
						  <a href="#" onClick="history.back();">Click Here to go Back</a>
						  </font></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
				</table>
				<%	
				'Delete source file & end script
				File.Delete
				Response.End
			End If
		End if
	
		FileName = lcase(File.FileName)
	
		If (pcv_ResizeObj=1) then
			call UseASPJpeg(FileName,File.Path)
		else
			call UseAspImage(FileName,File.Path)
		end if
	
		'Delete source file
		File.Delete
	
	Next

End Function

Function UseASPSmartUpload()

	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	
	mySmartUpload.Upload
	
	If PPD="1" then
		intCount = mySmartUpload.Save(uploadpath)
	else
		intCount = mySmartUpload.Save(uploadpath)
	end if
	
	thumbnailsize = mySmartUpload.Form("thumbnailsize")
	generalsize = mySmartUpload.Form("generalsize")
	detailsize = mySmartUpload.Form("detailsize")
	sharpen = mySmartUpload.Form("sharpen")
	resizexy = mySmartUpload.Form("resizexy")
	
	'Count files in mySmartUpload.  If none exist, exit script
	If intCount = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload any images.</b><br><br>
                  <a href="#" onClick="history.back();">Click Here to go Back</a>
                  </font></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        </table>
        <% Response.End
	End If
	
	'Run the resizer 
	For Each File in mySmartUpload.Files
	
		FileName = lcase(File.FileName)
	
		If NOT IsUploadAllowed(FileName) OR (right(FileName,3) <> "jpg" and right(FileName,3) <> "jpe" and right(FileName,4) <> "jpeg") Then
			If NOT IsUploadAllowed(FileName) OR (pcv_ResizeObj=2) or ((pcv_ResizeObj=1) and (right(FileName,3)<>"gif")) then  %>
                <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
                <tr> 
                  <td> 
                    <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
                      <tr> 
                        <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
                      </tr>
                      <tr> 
                        <td height="10"></td>
                      </tr>
                      <tr> 
                        <td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
                            <%If (pcv_ResizeObj=2) then%>
                                <b>This is not a JPG image.<br>GIF files are not supported with the ASPImage Component.</b>
                            <%Else%>
                                <b>This is not a JPG/GIF image.<br>Only JPG/GIF files are supported with the ASPJpeg Component.</b>
                            <%End if%>
                            <br><br>
                          <a href="#" onClick="history.back();">Click Here to go Back</a>
                          </font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                </tr>
                </table>
                <% 'Delete source file & end script
                Set fso=Server.CreateObject("Scripting.FileSystemObject")
                Set afi = fso.GetFile(uploadpath & FileName)
                afi.Delete
                Set afi=nothing
                Response.End
            End If
		End If
	
		If (pcv_ResizeObj=1) then
			call UseASPJpeg(FileName,uploadpath & FileName)
		else
			call UseAspImage(FileName,uploadpath & FileName)
		end if
		
		'Delete source file
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set afi = fso.GetFile(uploadpath & FileName)
		afi.Delete
		Set afi=nothing
	
	Next

End Function

Sub ResizeX(intXSize)
	Dim intYSize
	intYSize = round((intXSize / Image.MaxX) * Image.MaxY)
	err.number=0
	Image.ResizeR intXSize, intYSize
	if err.number<>0 then
		Image.Resize intXSize, intYSize
	end if
End sub
	
Sub ResizeY(intYSize)
	Dim intXSize
	intXSize = round((intYSize / Image.MaxY) * Image.MaxX)
	err.number=0
	Image.ResizeR intXSize, intYSize
	if err.number<>0 then
		Image.Resize intXSize, intYSize
	end if
End sub

Sub UseAspImage(FileName,SourceFile)
	'Generate random number to append to filename
	randomnum = RandomNumber(2353)
	
	'Generate new thumbnail image filename
	If right(FileName, 4) = ".jpg" Then
		thumbfilename = replace(FileName,".jpg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		thumbfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		thumbfilename = replace(FileName,".jpe","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		thumbfilename = replace(FileName,".gif","") & "_" & randomnum & "_thumb.gif"
	End If
	thumbfilename = replace(thumbfilename,"%20","")
	thumbfilename = replace(thumbfilename," ","")
	
	'Generate new general image filename
	If right(FileName, 4) = ".jpg" Then
		generalfilename = replace(FileName,".jpg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		generalfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		generalfilename = replace(FileName,".jpe","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		generalfilename = replace(FileName,".gif","") & "_" & randomnum & "_general.gif"	
	End If
	generalfilename = replace(generalfilename,"%20","")
	generalfilename = replace(generalfilename," ","")
	
	'Generate new detail image filename
	If right(FileName, 4) = ".jpg" Then
		detailfilename = replace(FileName,".jpg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		detailfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		detailfilename = replace(FileName,".jpe","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		detailfilename = replace(FileName,".gif","") & "_" & randomnum & "_detail.gif"
	End If
	detailfilename = replace(detailfilename,"%20","")
	detailfilename = replace(detailfilename," ","")

	'---- SAVE THUMBNAIL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)

	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	If resizexy = "Width" Then
		jpg_width = cint(thumbnailsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(thumbnailsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	Image.FileName = catalogpath & thumbfilename
	Image.SaveImage
	Set Image = nothing

	'---- SAVE GENERAL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)


	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	
	If resizexy = "Width" Then
		jpg_width = cint(generalsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(generalsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	Image.FileName = catalogpath & generalfilename
	Image.SaveImage
	Set Image = nothing

	'---- SAVE DETAIL IMAGE ----
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(SourceFile)

	BigBeforeWidth = Image.MaxX
	BigBeforeHeight = Image.MaxY
	
	If resizexy = "Width" Then
		jpg_width = cint(detailsize)
		BigAfterWidth = jpg_width
		BigAfterHeight = round ((jpg_width / Image.MaxX) * Image.MaxY)
		If cint(BigAfterWidth) >= cint(BigBeforeWidth) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeX(jpg_width)
		End If
	Else
		jpg_height = cint(detailsize)
		BigAfterHeight = jpg_height
		BigAfterWidth = round((jpg_height / Image.MaxY) * Image.MaxX)
		If cint(BigAfterHeight) >= cint(BigBeforeHeight) Then
			BigAfterWidth = Image.MaxX
			BigAfterHeight = Image.MaxY
		Else
			If sharpen = 1 Then
				Image.Sharpen 1
			End If
			Image.JPEGQuality = 85
			call ResizeY(jpg_height)
		End If	
	End If
	
	
	Image.FileName = catalogpath & detailfilename
	Image.SaveImage
	Set Image = nothing
End Sub

Sub UseASPJpeg(FileName,SourceFile)

	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	
	'Generate random number to append to filename
	randomnum = RandomNumber(2353)
	
	'Generate new thumbnail image filename
	If right(FileName, 4) = ".jpg" Then
		thumbfilename = replace(FileName,".jpg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		thumbfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		thumbfilename = replace(FileName,".jpe","") & "_" & randomnum & "_thumb.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		thumbfilename = replace(FileName,".gif","") & "_" & randomnum & "_thumb.gif"
	End If
	thumbfilename = replace(thumbfilename,"%20","")
	thumbfilename = replace(thumbfilename," ","")
	
	'Generate new general image filename
	If right(FileName, 4) = ".jpg" Then
		generalfilename = replace(FileName,".jpg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		generalfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		generalfilename = replace(FileName,".jpe","") & "_" & randomnum & "_general.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		generalfilename = replace(FileName,".gif","") & "_" & randomnum & "_general.gif"	
	End If
	generalfilename = replace(generalfilename,"%20","")
	generalfilename = replace(generalfilename," ","")
	
	'Generate new detail image filename
	If right(FileName, 4) = ".jpg" Then
		detailfilename = replace(FileName,".jpg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 5) = ".jpeg" Then
		detailfilename = replace(FileName,".jpeg","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".jpe" Then
		detailfilename = replace(FileName,".jpe","") & "_" & randomnum & "_detail.jpg"
	ElseIf right(FileName, 4) = ".gif" Then
		detailfilename = replace(FileName,".gif","") & "_" & randomnum & "_detail.gif"
	End If
	detailfilename = replace(detailfilename,"%20","")
	detailfilename = replace(detailfilename," ","")
	
	'---- SAVE THUMBNAIL IMAGE ----
	Jpeg.Open SourceFile
		
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight

	If resizexy = "Width" Then
		If cint(thumbnailsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = thumbnailsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(thumbnailsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = thumbnailsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If

	
	Jpeg.Save (catalogpath) & thumbfilename
	Jpeg.Close
	
	'---- SAVE GENERAL IMAGE ----
	Jpeg.Open SourceFile
	
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight
	
	If resizexy = "Width" Then
		If cint(generalsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = generalsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(generalsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = generalsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If	
	
	
	Jpeg.Save (catalogpath) & generalfilename
	Jpeg.Close
	
	'---- SAVE DETAIL IMAGE ----
	Jpeg.Open SourceFile
	
	BigBeforeWidth = jpeg.OriginalWidth
	BigBeforeHeight = jpeg.OriginalHeight

	If resizexy = "Width" Then
		If cint(detailsize) >= cint(BigBeforeWidth) Then
		Else
			BigAfterWidth = detailsize
			BigAfterHeight = round((BigAfterWidth / jpeg.Width) * jpeg.Height)
			
			Jpeg.Width = BigAfterWidth
			jpeg.Height = BigAfterHeight
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85
			
		End If
	Else
		If cint(detailsize) >= cint(BigBeforeHeight) Then
		Else
			BigAfterHeight = detailsize
			BigAfterWidth = round((BigAfterHeight / jpeg.Height) * jpeg.Width)
	
			Jpeg.Height = BigAfterHeight
			jpeg.Width = BigAfterWidth
	
			If sharpen = 1 Then
				Jpeg.Sharpen .1, 101
			End If
			Jpeg.Interpolation = 2
			jpeg.Quality = 85			
		End If	
	End If	
	
	Jpeg.Save (catalogpath) & detailfilename
	Jpeg.Close

End Sub

SELECT CASE pcv_UploadObj
	Case 1: UseSAFileUp()
	Case 2: UseASPUpload()
	Case 3: UseASPSmartUpload()
END SELECT%>

<script>
	function fillparentform() {
	parent.opener.document.hForm.smallImageUrl.value = "<%= thumbfilename %>"
	parent.opener.document.hForm.imageUrl.value = "<%= generalfilename %>"
	parent.opener.document.hForm.largeImageUrl.value = "<%= detailfilename %>"
	}
	
	fillparentform();
	
	imagename='';
	function enlrge(imgnme) {
		lrgewin=window.open("about:blank","","height=200,width=200")
		imagename=imgnme;
		setTimeout('update()',500)
	}
	
	function update() {
	doc=lrgewin.document;
	doc.open('text/html');
	doc.write('<HTML><HEAD><TITLE>Enlarged Image<\/TITLE><\/HEAD><BODY bgcolor="white" onLoad="if  (self.resizeTo)self.resizeTo((document.images[0].width+10),(document.images[0].height+80))" topmargin="4" leftmargin="0" rightmargin="0" bottommargin="0"><table border="0" cellspacing="0" cellpadding="0"><tr><td>');
	doc.write('<IMG SRC="' + imagename + '"><\/td><\/tr><tr><td><form name="viewn"><input type="image" src="../../pc/images/close.gif" align="right" value="Close Window" onClick="self.close()"><\/td><\/tr><\/table>');
	doc.write('<\/form><\/BODY><\/HTML>');
	doc.close();
	}
</script>
  <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td> 
        <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
          <tr> 
            <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
          </tr>
          <tr> 
            <td height="10"></td>
          </tr>
          <tr> 
            <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <b>Image Upload & Resizing Completed Successfully!</b><br><br>
			  The filenames for the 3 images have been sent to the product window.<br><br><br>
			  <b>Thumbnail Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= thumbfilename %>')"><%= thumbfilename %></a><br><br>
			  <b>General Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= generalfilename %>')"><%= generalfilename %></a><br><br>
			  <b>Detail Image:</b><br><a href="javascript:enlrge('../../pc/catalog/<%= detailfilename %>')"><%= detailfilename %></a>
			  <br><br><br>
			  <a href="#" onClick="self.close()"><img src="../../pc/images/close.gif" alt="Close Window" border="0"></a>
			  </font></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</body>
</html>