<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<% response.Buffer=true 
Server.ScriptTimeout = 120 %>
<% PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% 
on error resume next
dim query, conntemp, rs, idproduct, pSmallImageUrl,pImageUrl, pLargeImageUrl, title, saction, message, maxOrder

idproduct=request.QueryString("idproduct")
idimg=request.QueryString("imgid")
message = ""
maxOrder = 1

if idimg<>"" then
    title = "Edit View"
    saction = "update"
    call opendb()

    err.Clear
    query="SELECT pcProdImage_Url,pcProdImage_LargeUrl FROM pcProductsImages WHERE pcProdImage_ID="&idimg&" "
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    if err.number <> 0 then
        set rs=nothing	
        call closedb()
        response.redirect "techErr.asp?error="& Server.Urlencode("Error in addImg_popup line 24: "&Err.Description) 
    end if

    if not rs.EOF then
        pImageUrl=rs("pcProdImage_Url")
        pLargeImageUrl=rs("pcProdImage_LargeUrl")
    end if

    set rs=nothing
    call closedb()
else
    title = "Add View"
    saction = "add"
end if

%>

<html>
<head>
<title><%=title%></title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:470px; background-image: none;">
<% 
if request("action")="update" THEN
    pImageUrl=request("imageUrl")
    pLargeImageUrl=request("largeImageUrl")

    '// Update Additional Product Images if there are any
    if pImageUrl<>"" OR pLargeImageUrl<>"" then
	    call opendb()
    
        query="UPDATE pcProductsImages SET pcProdImage_Url='"&pImageUrl&"',pcProdImage_LargeUrl='"&pLargeImageUrl&"' WHERE pcProdImage_ID="&idimg&" "
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        set rs=nothing
	
	    call closedb()
	    message = "View successfully updated."
			Dim pImgAdded
			pImgAdded=1
    end if
end if

if request("action")="add" THEN
    pImageUrl=request("imageUrl")
    pLargeImageUrl=request("largeImageUrl")
	pcv_intImageError = 0
	
	If len(pImageUrl)>49 Then
		message = "The general image URL is too long. Please shorten the image name."
		pcv_intImageError = 1
	End If
	
	If len(pLargeImageUrl)>49 Then
		message = "The large image URL is too long. Please shorten the image name."
		pcv_intImageError = 1
	End If

    '// Insert Additional Product Images if there are any
    if (pcv_intImageError = 0) AND (pImageUrl<>"" OR pLargeImageUrl<>"") then
	    call opendb()
        
        err.Clear
        query="SELECT max(pcProdImage_Order) as maxord from pcProductsImages where idProduct="&idProduct&" "
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            set rs=nothing	
	        call closedb()
	        response.redirect "techErr.asp?error="& Server.Urlencode("Error in addImg_popup line 75: "&Err.Description) 
        end if

        if not rs.EOF then
            maxOrder = cint(rs("maxord")) + 1
        end if
        set rs=nothing
        
        if maxOrder=0 or maxOrder="" then
            maxOrder = 1
        end if

        err.Clear
        query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl,pcProdImage_Order) VALUES("&idProduct&",'"&pImageUrl&"','"&pLargeImageUrl&"',"&maxOrder&")"
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        if err.number <> 0 then
            set rs=nothing	
	        call closedb()
	        response.redirect "techErr.asp?error="& Server.Urlencode("Error in addImg_popup line 93: "&Err.Description) 
        end if
        
        set rs=nothing	
	    call closedb()
	    message = "View successfully added."
			pImgAdded=1
    end if


end if
%>
	<form name="hForm" method="post" action="addImg_popup.asp?action=<%=saction %>&idproduct=<%=idProduct%>&imgid=<%=idimg%>" class="pcForms">
	    <table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
		    <tr>
			    <th colspan="2"><%=title%></th>
		    </tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
	        <tr>
		        <td colspan="2">Type in the file name, not the file path. All images must be located in the 'pc/catalog' folder. When you upload an image, it is automatically saved to that folder.
		        <!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
				<%If HaveImgUplResizeObjs=1 then%>
			        To upload and resize an image <a href="javascript:;" onClick="window.open('uploadresize/productResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
		        <% Else %>
			        To upload an image <a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
		        <% End If %>
		        </td>
	        </tr>
	        <tr>
		        <script language="JavaScript"><!--
			        function chgAddWin(file,window) {
    			        msgAddWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
				        if (msgAddWindow.opener == null) msgAddWindow.opener = self;
		            }
		        //--></script> 
		        <td width="20%" align="right" nowrap="nowrap">General Image:</td>
		        <td width="80%">  
		            <input type="text" name="imageUrl" value="<%response.write pImageUrl%>" size="30"><a href="javascript:;" onClick="chgAddWin('../pc/imageDir.asp?ffid=imageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>  
		            <input type="hidden" name="smallImageUrl" value="<%response.write pSmallImageUrl%>">  
		        </td>
	        </tr>
	        <tr> 
		        <td align="right" nowrap="nowrap">Detail View Image:</td>
		        <td> 
			        <input type="text" name="largeImageUrl" value="<%response.write pLargeImageUrl%>" size="30"><a href="javascript:;" onClick="chgAddWin('../pc/imageDir.asp?ffid=largeImageUrl&fid=hForm','window2')"><img src="images/search.gif" alt="Locate previously uploaded images" width="16" height="16" border=0 hspace="3"></a>
		        </td>
	        </tr>
	        <tr> 
		        <td colspan="2" align="center"> 
			        <font color=red><%response.write message%></font>
		        </td>
	        </tr>
			<tr>
				<td colspan="2" align="center">
					<% if pImgAdded<>1 then %>
				    <input type="submit" name="Submit" value="Save" class="submit2">
					<% end if %>
				    <input type="button" name="Back" value="Close" onClick="opener.document.hForm.submit(); self.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
