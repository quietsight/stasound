<%
Sub Delay(DelaySeconds)
	SecCount = 0
	Sec2 = 0
	While SecCount < DelaySeconds + 1
		Sec1 = Second(Time())
		If Sec1 <> Sec2 Then
			Sec2 = Second(Time())
			SecCount = SecCount + 1
		End If
	Wend 
End Sub

Function FTPUpload(sSite, sUsername, sPassword, sLocalFile, sRemotePath)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1
  Const ForWriting = 2
  
  Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
  Set oFTPScriptShell = CreateObject("WScript.Shell")

  sRemotePath = Trim(sRemotePath)
  sLocalFile = Trim(sLocalFile)
  
  '----------Path Checks---------
  'Here we willcheck the path, if it contains
  'spaces then we need to add quotes to ensure
  'it parses correctly.
  If InStr(sRemotePath, " ") > 0 Then
    If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
      sRemotePath = """" & sRemotePath & """"
    End If
  End If
  
  If InStr(sLocalFile, " ") > 0 Then
    If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
      sLocalFile = """" & sLocalFile & """"
    End If
  End If

  'Check to ensure that a remote path was
  'passed. If it's blank then pass a "\"
  If Len(sRemotePath) = 0 Then
    'Please note that no premptive checking of the
    'remote path is done. If it does not exist for some
    'reason. Unexpected results may occur.
    sRemotePath = "\"
  End If
  
  'Check the local path and file to ensure
  'that either the a file that exists was
  'passed or a wildcard was passed.
  If InStr(sLocalFile, "*") Then
    If InStr(sLocalFile, " ") Then
      FTPUpload = "FTP Error: Wildcard uploads do not work if the path contains a " & _
      "space." & vbCRLF
      FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
      Exit Function
    End If
  ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
    'nothing to upload
    FTPUpload = "FTP Error: File Not Found."
    Exit Function
  End If
  '--------END Path Checks---------
  
  'build input file for ftp command
  sFTPScript = sFTPScript & "USER " & sUsername & vbCRLF
  sFTPScript = sFTPScript & sPassword & vbCRLF
  sFTPScript = sFTPScript & "cd " & sRemotePath & vbCRLF
  sFTPScript = sFTPScript & "binary" & vbCRLF
  sFTPScript = sFTPScript & "prompt n" & vbCRLF
  sFTPScript = sFTPScript & "put " & sLocalFile & vbCRLF
  sFTPScript = sFTPScript & "quit" & vbCRLF & "quit" & vbCRLF & "quit" & vbCRLF


  sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
  sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
  sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName

  'Write the input file for the ftp command
  'to a temporary file.
  Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
  fFTPScript.WriteLine(sFTPScript)
  fFTPScript.Close
  Set fFTPScript = Nothing  

  oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & _
  " > " & sFTPResults
  
  sResults=""
  tmpFileSize=-1
  While len(sResults)<=0 OR tmpFileSize<>sResults
  Delay(1)
  tmpFileSize=sResults
  'Check results of transfer.
  on error resume next
  Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, _
  FailIfNotExist, OpenAsDefault)
  if err.number <> 0 then
  	err.number=0
  	err.description=""
  else
	  sResults = fFTPResults.ReadAll
	  fFTPResults.Close
  end if
  Wend
  
  oFTPScriptFSO.DeleteFile(sFTPTempFile)
  oFTPScriptFSO.DeleteFile (sFTPResults)
  
  If InStr(sResults, "226 Transfer complete.") > 0 Then
    FTPUpload = "OK"
  ElseIf InStr(sResults, "File not found") > 0 Then
    FTPUpload = "FTP Error: File Not Found"
  ElseIf InStr(sResults, "cannot log in.") > 0 Then
    FTPUpload = "FTP Error: Login Failed."
  Else
    FTPUpload = "FTP Error: Unknown."
  End If

  Set oFTPScriptFSO = Nothing
  Set oFTPScriptShell = Nothing
End Function

%>
