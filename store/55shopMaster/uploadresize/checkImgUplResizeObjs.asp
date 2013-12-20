<%
Dim HaveImgUplResizeObjs
Dim pcv_UploadObj
Dim	pcv_ResizeObj

HaveImgUplResizeObjs=0
pcv_UploadObj=0
pcv_ResizeObj=0

Function IsObjInstalled(strClassString)
	On Error Resume Next
	' initialize default values
	IsObjInstalled = False
	Err = 0
	' testing code
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	' cleanup
	Set xTestObj = Nothing
	Err = 0
End Function

if IsObjInstalled("SoftArtisans.FileUp") then
	HaveImgUplResizeObjs=1
	pcv_UploadObj=1
else
	if IsObjInstalled("Persits.Upload") then
		HaveImgUplResizeObjs=1
		pcv_UploadObj=2
	else
		if IsObjInstalled("aspSmartUpload.SmartUpload") then
			HaveImgUplResizeObjs=1
			pcv_UploadObj=3
		end if
	end if
end if

if HaveImgUplResizeObjs=1 then
	If IsObjInstalled("Persits.Jpeg") then
		HaveImgUplResizeObjs=2
		pcv_ResizeObj=1
	else
		If IsObjInstalled("AspImage.Image") then
			HaveImgUplResizeObjs=2
			pcv_ResizeObj=2
		end if
	end if
end if

if HaveImgUplResizeObjs=2 then
	HaveImgUplResizeObjs=1
else
	HaveImgUplResizeObjs=0
end if
%>