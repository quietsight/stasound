<%Response.ContentType = "text/xml"%><?xml version="1.0" ?>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="inc_srcImgsQuery.asp"-->
<%totalrecords=0
Dim connTemp
call opendb()
Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText

rstemp.AbsolutePage=iPageCurrent
totalrecords=clng(rstemp.RecordCount)

Dim strCol, Count, DCnt, HTMLResult, intPicsPerRow, PopFile,ShowSub,ShowPic,pShowButton,pName
HTMLResult=""
Count = 0
DCnt=0
strCol = ""
intPicsPerRow=3
PopFile = "showPicture.asp"
ShowSub = "catalog"
pName=""
pShowButton = 1

src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_key4=getUserInput(request("key4"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
src_FromPage=getUserInput(request("src_FromPage"),0)
src_Button3=getUserInput(request("src_Button3"),0)
pshowimage=getUserInput(request("showimage"),0)
fid=getUserInput(request("fid"),0)
ffid=getUserInput(request("ffid"),0)
submit=request("Submit")
submit2=request("Submit2")


HTMLResult=HTMLResult & "<table class=""pcContent"">" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<td colspan=4>" & vbcrlf
HTMLResult=HTMLResult & "<h2>Locate an Image</h2>" & vbcrlf
HTMLResult=HTMLResult & "</td>" & vbcrlf
HTMLResult=HTMLResult & "</tr>" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<td colspan=4>" & vbcrlf
HTMLResult=HTMLResult & "Click on an image or image name to view its actual size. Images are automatically resized to better fit into this window, and therefore <b>may appear distorted</b>." & vbcrlf
HTMLResult=HTMLResult & "</td>" & vbcrlf
HTMLResult=HTMLResult & "</tr>" & vbcrlf
HTMLResult=HTMLResult & "<tr>" & vbcrlf
HTMLResult=HTMLResult & "<td colspan=4>" & vbcrlf
HTMLResult=HTMLResult & "<hr>" & vbcrlf
HTMLResult=HTMLResult & "</td>" & vbcrlf
HTMLResult=HTMLResult & "</tr>" & vbcrlf

do while (not rsTemp.eof) and (count < rsTemp.pageSize)
				
	If strCol <> "#FFFFFF" Then
		strCol = "#FFFFFF"
	Else 
		strCol = "#E1E1E1"
	End If

	HTMLResult=HTMLResult & "<tr bgcolor=""" & strCol & """>" & vbcrlf

    PictureNo = 0
    for x = 1 to intPicsPerRow
        DCnt=DCnt+1
	    count=count + 1
		pName=rstemp("pcImgDir_Name")
        ShowPic = Replace(pName, " ", "%20")
        
	    HTMLResult=HTMLResult & "<td valign=""bottom"" align=""center"">" & vbcrlf
	    HTMLResult=HTMLResult & "<FORM NAME=""inputForm"&DCnt&""" onSubmit=""return setForm"&DCnt&"();"" class=""pcForms"">" & vbcrlf
	    HTMLResult=HTMLResult & "<INPUT NAME=""inputField"&DCnt&""" TYPE=""hidden"" VALUE="""&ShowPic&""">" & vbcrlf
        
        if pshowimage="YES" then
	        HTMLResult=HTMLResult & "<div>" & vbcrlf
   	        HTMLResult=HTMLResult & "<a href=""Javascript:openGalleryWindow('"&PopFile&"?ShowPic="&ShowSub&"/"&ShowPic&"')""><img src='"&ShowSub&"/"&pName&"' width=""40"" vspace=""2""></a>" & vbcrlf
	        HTMLResult=HTMLResult & "</div>" & vbcrlf
	    end if

	    HTMLResult=HTMLResult & "<div class=""pcSmallText"" style=""padding-bottom:2px;"">" & vbcrlf
	    HTMLResult=HTMLResult & "<a href=""Javascript:openGalleryWindow('"&PopFile&"?ShowPic="&ShowSub&"/"&ShowPic&"')"">"&Mid(pName,1,Len(pName)-4)&"</a>" & vbcrlf
	    HTMLResult=HTMLResult & "</div>" & vbcrlf
	    HTMLResult=HTMLResult & "<div>" & vbcrlf
	    HTMLResult=HTMLResult & "<input type=""checkbox"" name=""check"&DCnt&""" value="&ShowSub&"/"&ShowPic&" onBlur=""javascript:if (this.checked==true) {document.image1.cimg"&DCnt&".value=this.value;} else {document.image1.cimg"&DCnt&".value='';}"" class=""clearBorder"">&nbsp;"
    	
	    if pShowButton <> 0 then 
	        HTMLResult=HTMLResult & "<INPUT TYPE=""SUBMIT"" name=""UPD"&DCnt&""" VALUE=""Select"" class=""submit2"">" & vbcrlf
	    end if 
    	
	    HTMLResult=HTMLResult & "</div>" & vbcrlf
	    HTMLResult=HTMLResult & "</FORM>" & vbcrlf
	    HTMLResult=HTMLResult & "</td>" & vbcrlf

        if count = rsTemp.pageSize OR count = totalrecords then exit do
        
        PictureNo = PictureNo + 1
        If PictureNo=intPicsPerRow Then
            HTMLResult=HTMLResult & "</tr><tr><td colspan=""3"" class=""pcSpacer""></td></tr>" & vbcrlf
            PictureNo = 0
        End if
        
        rsTemp.MoveNext

        if rsTemp.eof then
        	exit do
        end if
        
    next        
loop

If PictureNo<>intPicsPerRow Then
    HTMLResult=HTMLResult & "</tr><tr><td colspan=""3"" class=""pcSpacer""></td></tr>" & vbcrlf
    PictureNo = 0
End if

HTMLResult=HTMLResult & "</table>" & vbcrlf
HTMLResult=HTMLResult & "</div>" & vbcrlf

IF session("admin")=-1 then
    IF DCnt>0 THEN
        HTMLResult=HTMLResult & "<br /><br /><div align=""center""><form name='image1' method='post' class='pcForms'>"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='fid' value="&fid&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='ffid' value="&ffid&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='ajaxSearch' value='submit'>"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='imgIndex' value='submit'>"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='src_FromPage' value="&src_FromPage&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='src_Button3' value="&src_Button3&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='showimage' value="&pshowimage&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='resultCnt' value="&form_resultCnt&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='Page' value="&iPageCurrent&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='ShowSub' value="&ShowSub&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='key1' value="&form_key1&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='key2' value="&form_key2&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='key3' value="&form_key3&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='hidden' name='key4' value="&form_key4&">"&vbCrlf
            HTMLResult=HTMLResult & "<input type='submit' name='del' value='Delete Checked Images' onclick=""return(confirm('You are about to delete the images you have selected. This action cannot be undone. Would you like to continue?'));"" class='submit2'>"&vbCrlf
            For i=1 to DCnt
                HTMLResult=HTMLResult & "<input type='hidden' name=""cimg"&i&""" value=''>"&vbCrlf
            Next
            HTMLResult=HTMLResult & "<input type='hidden' name=""dCount"" value="&DCnt&">"&vbCrlf
        HTMLResult=HTMLResult & "</form></div>"&vbCrlf
      END IF
END IF

HTMLResult=HTMLResult & "<input type=hidden name=count value="&count&">" & vbcrlf

set rstemp=nothing
call closedb()

'*** Fixed FireFox issues
Dim tmpData,tmpData1
Dim tmp1,tmp2,i,Count1
tmpData=Server.HTMLEncode(HTMLResult)
Count1=0
tmpData1=""
tmp1=split(tmpData,"&lt;/tr&gt;")
For i=lbound(tmp1) to ubound(tmp1)
	if i>lbound(tmp1) then
		tmp2="&lt;/tr&gt;" & tmp1(i)
	else
		tmp2=tmp1(i)
	end if
	Count1=Count1+1
	tmpData1=tmpData1 & "<data" & Count1 & ">" & tmp2 & "</data" & Count1 & ">" & vbcrlf
Next
%><note>
<data0><%=Count1%></data0>
<%=tmpData1%>
</note>