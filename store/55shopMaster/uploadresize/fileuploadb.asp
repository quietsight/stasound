<%@ LANGUAGE="VBSCRIPT" %>
<%
' PRV41 Start
pageTitle="File Upload" 
%>
<% Section="products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../../includes/ppdstatus.inc"-->
<!--#include file="../../includes/productcartFolder.asp"-->
<!--#include file="checkImgUplResizeObjs.asp"-->
<!--#include file="../../includes/pcSanitizeUpload.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Upload File</title>
<link href="../pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<%If pcv_UploadObj=0 then%>
    <table class="pcCPcontent">
        <tr>
            <td>
            <div class="pcCPmessage">We are unable to find compatible Upload server components. Please consult the User Guide for detailed system requirements.</div>
            </td>
        </tr>
        <tr>
            <td align="center"><input type="button" name="Close" value=" Close window " onClick="javascript:window.close();" class="ibtnGrey"></td>
        </tr>
    </table>
	<% response.end
End If%>
<%
Dim catalogpath, uploadpath, normalfilename, largefilename, normalsize, largesize, sharpen, countfiles
Dim randomnum, FileName, BigBeforeWidth, BigBeforeHeight, BigAfterWidth, BigAfterHeight, imgcomp
Dim Image
Dim resizexy

Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

uploadpath=Server.Mappath ("..\..\pc\library\")
uploadpath = uploadpath & "\"

'on error resume next

Function UseSAFileUp()

	'--- Instantiate the FileUp object
	Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	
	Upload.Path = uploadpath
	
	If NOT IsUploadAllowed(Upload.Form("file1").UserFilename) or Upload.Form("file1").UserFilename = "" Then 	%>
		  <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
			<tr> 
			  <td> 
				<table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
				  <tr> 
					<td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload</font></b></font></td>
				  </tr>
				  <tr> 
					<td height="10"></td>
				  </tr>
				  <tr> 
					<td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
					  <b>You did not upload anything.</b><br><br>
					  <a href="fileuploada.asp">Click Here to go Back</a>
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
	
	<%	If LCase(right(Upload.Form("file1").UserFilename,3)) <> "txt" Then %>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
                    <b>This is not a TXT file. For security reasons, only files ending with an extension of '.TXT' are allowed.</b>
                    <br><br>
                  <a href="fileuploada.asp">Click Here to go Back</a>
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
	End If

	Upload.Form("FILE1").Save

	FileName = Mid(Upload.Form("file1").UserFilename, InstrRev(Upload.Form("file1").UserFilename, "\") + 1)
	FileName = lcase(Filename)

	Set Upload = nothing

End Function

Function UseASPUpload()

	Set Upload = Server.CreateObject("Persits.Upload")
	
	Upload.SaveVirtual "..\..\pc\library\"
	
	
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
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload anything.</b><br><br>
                  <a href="fileuploada.asp">Click Here to go Back</a>
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
		Response.End
	End If
	
	'Run the resizer 
	For Each File in Upload.Files
	
		If NOT IsUploadAllowed(File.FileName) OR LCase(Right(File.FileName,3)) <> "txt" Then %>
            <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
            <tr> 
              <td> 
                <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
                  <tr> 
                    <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload</font></b></font></td>
                  </tr>
                  <tr> 
                    <td height="10"></td>
                  </tr>
                  <tr> 
                    <td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
                        <b>This is not a TXT file. For security reasons, only files ending with an extension of '.TXT' are allowed.</b>
                        <br><br>
                      <a href="fileuploada.asp">Click Here to go Back</a>
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
		End if
		
		FileName = lcase(File.FileName)
	
	Next

End Function

Function UseASPSmartUpload()

	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	
	mySmartUpload.Upload
	
	intCount = mySmartUpload.Save(uploadpath)
	
	
	'Count files in mySmartUpload.  If none exist, exit script
	If intCount = 0 Then%>
        <table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
        <tr> 
          <td> 
            <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload</font></b></font></td>
              </tr>
              <tr> 
                <td height="10"></td>
              </tr>
              <tr> 
                <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
                  <b>You did not upload anything.</b><br><br>
                  <a href="fileuploada.asp">Click Here to go Back</a>
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
		Response.End
	End If
	
	'Run the resizer 
	For Each File in mySmartUpload.Files
	
		FileName = lcase(File.FileName)
		
		If NOT IsUploadAllowed(FileName) OR right(FileName,3) <> "txt" Then	%>
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
						<b>This is not a TXT file. For security reasons, only files ending with an extension of '.TXT' are allowed.</b>
						<br><br>
					  <a href="fileuploada.asp">Click Here to go Back</a>
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
			Set fso=Server.CreateObject("Scripting.FileSystemObject")
			Set afi = fso.GetFile(uploadpath & FileName)
			afi.Delete
			Set afi=nothing
			Response.End
		End If
	Next

End Function

SELECT CASE pcv_UploadObj
	Case 1: UseSAFileUp()
	Case 2: UseASPUpload()
	Case 3: UseASPSmartUpload()
END SELECT%>


<script>
function fillparentform() {
parent.opener.document.hForm.sendreviewremindertemplate.value = '<%= replace(filename,"'","\'") %>'
}

fillparentform();

</script>

<table width="100%" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
<tr> 
  <td> 
    <table width="90%" border="0" cellspacing="0" cellpadding="4" align="center">
      <tr> 
        <td bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;File Upload (.TXT file)</font></b></font></td>
      </tr>
      <tr> 
        <td height="10"></td>
      </tr>
      <tr> 
        <td align="center"><font face="Arial, Helvetica, sans-serif" size="2"> 
          <b>File Upload Completed Successfully!</b><br><br>
          
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
<% 'PRV41 end %>