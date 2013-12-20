<%@ LANGUAGE = VBScript.Encode %>
<%Server.ScriptTimeout = 5400%>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Export Product XML File" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="incFTPUploadFunc.asp"-->
<!--#include file="../xml/commonVariables.asp"-->
<!--#include file="../xml/commonFunctions.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<%Dim connTemp,query,rs
Dim iXML,oXML,iRoot,oRoot,tmpNode,tmpNode1,attNode,subNode,ChildNodes,iXML1,iRoot1
Dim ErrorCode,ErrorDesc
Dim XMLStream,tmpStream,tmpStream1,tmpStream2,tmpFileName,tmpGenName

ErrorCode=0
ErrorDesc=""
XMLStream=""
tmpFileName=""

lResolve = 120 * 1000
lConnect = 120 * 1000
lSend = 120 * 1000
lReceive = 120 * 1000

Sub XMLCreateNode(parentNode,tmpNodeName,tmpValue)
Dim attNode
	Set attNode=oXML.createNode(1,tmpNodeName,"")
	if tmpValue<>"" then
		if (tmpValue=-1) and (tmpNodeName<>prdStock_name) then
			tmpValue=1
		end if
		attNode.Text=tmpValue
	end if
	parentNode.appendChild(attNode)
End Sub

Function ConvertToXMLDate(tmpDate)
Dim tmp1,tmp2,tmp3
	tmp1=CDate(tmpDate)
	tmp2=Year(tmp1)
	tmp3=Month(tmp1)
	if tmp3<10 then
		tmp3="0" & tmp3
	end if
	tmp2=tmp2 & "-" & tmp3
	tmp3=Day(tmp1)
	if tmp3<10 then
		tmp3="0" & tmp3
	end if
	tmp2=tmp2 & "-" & tmp3
	ConvertToXMLDate=tmp2
End Function

IF request("action")="newsrc" THEN
	tmp_idcategory=request("idcategory")
	if tmp_idcategory="" then
		tmp_idcategory=0
	end if
	tmp_customfield=request("customfield")
	if tmp_customfield="" then
		tmp_customfield=0
	end if
	tmp_SearchValues=request("SearchValues")
	if tmp_SearchValues<>"" then
		tmp_SearchValues=Server.HTMLEncode(tmp_SearchValues)
	end if
	tmp_priceFrom=request("priceFrom")
	if tmp_priceFrom="" then
		tmp_priceFrom=0
	end if
	tmp_priceUntil=request("priceUntil")
	if tmp_priceUntil="" then
		tmp_priceUntil=0
	end if
	tmp_withstock=request("withstock")
	if tmp_withstock="" then
		tmp_withstock=0
	end if
	tmp_sku=request("sku")
	if tmp_sku<>"" then
		tmp_sku=Server.HTMLEncode(tmp_sku)
	end if
	tmp_IDBrand=request("IDBrand")
	if tmp_IDBrand="" then
		tmp_IDBrand=0
	end if
	tmp_keyWord=request("keyWord")
	if tmp_keyWord<>"" then
		tmp_keyWord=Server.HTMLEncode(tmp_keyWord)
	end if
	tmp_exact=request("exact")
	if tmp_exact="" then
		tmp_exact=0
	end if
	tmp_IncNormal=request("src_IncNormal")
	if tmp_IncNormal="" then
		tmp_IncNormal=0
	end if
	tmp_IncBTO=request("src_IncBTO")
	if tmp_IncBTO="" then
		tmp_IncBTO=0
	end if
	tmp_IncItem=request("src_IncItem")
	if tmp_IncItem="" then
		tmp_IncItem=0
	end if
	if tmp_IncBTO=0 AND tmp_IncItem=0 then
		tmp_IncNormal=1
	end if
	tmp_pinactive=request("pinactive")
	if tmp_pinactive="" then
		tmp_pinactive=0
	end if
	tmp_pdeleted=request("pdeleted")
	if tmp_pdeleted="" then
		tmp_pdeleted=0
	end if
	tmp_pSpecial=request("pSpecial")
	if tmp_pSpecial="" then
		tmp_pSpecial=0
	end if
	tmp_pFeatured=request("pFeatured")
	if tmp_pFeatured="" then
		tmp_pFeatured=0
	end if
	tmp_pFromDate=request("pFromDate")
	if tmp_pFromDate<>"" then
		tmp_pFromDate=ConvertToXMLDate(tmp_pFromDate)
	end if
	tmp_pToDate=request("pToDate")
	if tmp_pToDate<>"" then
		tmp_pToDate=ConvertToXMLDate(tmp_pToDate)
	end if
	tmp_order=request("order")
	if tmp_order="" then
		tmp_order=0
	end if
	tmp_pHideExported=request("pHideExported")
	if tmp_pHideExported="" then
		tmp_pHideExported=0
	end if
	tmp_pFTPPartner=request("pFTPPartner")
	if tmp_pFTPPartner="" then
		tmp_pFTPPartner=0
	end if
	tmp_pRmvFile=request("pRmvFile")
	if tmp_pRmvFile="" then
		tmp_pRmvFile=0
	end if
	
	call opendb()
	query="SELECT pcXP_PartnerID,pcXP_Password,pcXP_Key FROM pcXMLPartners WHERE pcXP_ExportAdmin=1;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		tmp_PartnerID=rs("pcXP_PartnerID")
		tmp_PartnerPass=rs("pcXP_Password")
		if tmp_PartnerPass<>"" then
			tmp_PartnerPass=enDeCrypt(tmp_PartnerPass, scCrypPass)
		end if
		tmp_PartnerKey=rs("pcXP_Key")
	end if
	set rs=nothing
	call closedb()
	
	Set iXML=Server.CreateObject("MSXML2.DOMDocument")

	call InitResponseDocument(cm_SearchProductsRequest_name)
	
	Call XMLCreateNode(oRoot,cm_partnerID_name,tmp_PartnerID)
	Call XMLCreateNode(oRoot,cm_partnerPassword_name,tmp_PartnerPass)
	Call XMLCreateNode(oRoot,cm_partnerKey_name,tmp_PartnerKey)
	
	Set subNode = oXML.createNode(1,cm_filters_name,"")
	oRoot.appendChild(subNode)
	
	Call XMLCreateNode(subNode,srcCategoryID_name,tmp_idcategory)
	If tmp_customfield>"0" AND tmp_SearchValues<>"" then
		Call XMLCreateNode(subNode,srcCFieldID_name,tmp_customfield)
	end if
	If tmp_customfield>"0" AND tmp_SearchValues<>"" then
		Call XMLCreateNode(subNode,srcCFieldValue_name,tmp_SearchValues)
	end if
	Call XMLCreateNode(subNode,srcPriceFrom_name,tmp_priceFrom)
	Call XMLCreateNode(subNode,srcPriceTo_name,tmp_priceUntil)
	Call XMLCreateNode(subNode,srcInStock_name,tmp_withstock)
	if tmp_sku<>"" then
		Call XMLCreateNode(subNode,srcSKU_name,tmp_sku)
	end if
	Call XMLCreateNode(subNode,srcBrandID_name,tmp_IDBrand)
	if tmp_keyWord<>"" then
		Call XMLCreateNode(subNode,srcKeyword_name,tmp_keyWord)
	end if
	Call XMLCreateNode(subNode,srcExactPhrase_name,tmp_exact)
	Call XMLCreateNode(subNode,srcIncNormal_name,tmp_IncNormal)
	Call XMLCreateNode(subNode,srcIncBTO_name,tmp_IncBTO)
	Call XMLCreateNode(subNode,srcIncBTOItems_name,tmp_IncItem)
	Call XMLCreateNode(subNode,srcIncInactive_name,tmp_pinactive)
	Call XMLCreateNode(subNode,srcIncDeleted_name,tmp_pdeleted)
	Call XMLCreateNode(subNode,srcSpecial_name,tmp_pSpecial)
	Call XMLCreateNode(subNode,srcFeatured_name,tmp_pFeatured)
	Call XMLCreateNode(subNode,srcSort_name,tmp_order)
	if tmp_pFromDate<>"" then
		Call XMLCreateNode(subNode,srcFromDate_name,tmp_pFromDate)
	end if
	if tmp_pToDate<>"" then
		Call XMLCreateNode(subNode,srcToDate_name,tmp_pToDate)
	end if
	Call XMLCreateNode(subNode,srcHideExported_name,tmp_pHideExported)
	
	Set objXML = Server.CreateObject("MSXML2.serverXMLHTTP"&scXML)
	
	'Create Link
	strPathInfo=""
	strPath=Request.ServerVariables("PATH_INFO")
	iCnt=0
	do while iCnt<2
		if mid(strPath,len(strPath),1)="/" then
			iCnt=iCnt+1
		end if
		if iCnt<2 then
			strPath=mid(strPath,1,len(strPath)-1)
		end if
	loop

	strPathInfo=Request.ServerVariables("HTTP_HOST") & strPath
			
	if Right(strPathInfo,1)="/" then
	else
		strPathInfo=strPathInfo & "/"
	end if
			
	tmpHTTPs=Request.ServerVariables("HTTPS")
	if UCase(tmpHTTPs)="OFF" then
		tmpStoreURL="http://" & strPathInfo
	else
		tmpStoreURL="https://" & strPathInfo
	end if
	
	ProductCartXMLServer=tmpStoreURL & "/xml/gateway.asp"
	
	objXML.setTimeouts lResolve, lConnect, lSend, lReceive
	objXML.open "POST",ProductCartXMLServer, True
	objXML.setRequestHeader "XML-Agent", "ProductCart XML Partner"
	requestText=oXML.xml
	objXML.send(requestText)
	If pcf_IsResponseGood()=False Then
		objXML.Abort()
		Set objXML=nothing
		Set oRoot=nothing
		Set oXML=nothing
		Set iXML=nothing
		response.Redirect("XMLExportPrdFile.asp?msg=The server’s XML export component is not responding.  Please click the back button, wait 10 seconds, and try again. If you have already setup you’re XML Tools Partner per the instructions in the user guide, and you continue to receive this message, you should read our troubleshooting tips. If you have not setup your XML Tools Partner then you must download the XML Tools Guide and follow the setup instructions.")
		response.End()
	End If
	
	iXML.async=false
	iXML.load(objXML.responseXML)
	Set iRoot=iXML.documentElement
	
	ErrorCode=iRoot.selectSingleNode(cm_requestStatus_name).Text


	if ErrorCode="200" then
		ErrorCode=iRoot.selectSingleNode(cm_resultCount_name).Text
		if ErrorCode>"0" then
			ErrorCode=0
		else
			ErrorCode=300 '//Products not found
			ErrorDesc="Products not found"
		end if
	else
		Set subNode=iRoot.selectSingleNode(cm_errorList_name)
		ErrorCode=subNode.selectSingleNode(cm_errorCode_name).Text
		ErrorDesc=subNode.selectSingleNode(cm_errorDesc_name).Text
	end if
	
	if ErrorCode=0 then

		Set attNode=iRoot.selectSingleNode(cm_products)
		Set ChildNodes = attNode.childNodes		
		Set oRoot=nothing
		Set oXML=nothing
		pcv_CountCompleted=0
		pcv_CountTotal=0
		For Each subNode In ChildNodes
			tmpProductID=trim(subNode.Text)
			if tmpProductID<>"" then
				
				Call InitResponseDocument(cm_GetProductDetailsRequest_name)		
				Call XMLCreateNode(oRoot,cm_partnerID_name,tmp_PartnerID)
				Call XMLCreateNode(oRoot,cm_partnerPassword_name,tmp_PartnerPass)
				Call XMLCreateNode(oRoot,cm_partnerKey_name,tmp_PartnerKey)
				Call XMLCreateNode(oRoot,prdID_name,tmpProductID)
				Set tmpNode = oXML.createNode(1,cm_requests_name,"")
				oRoot.appendChild(tmpNode)
				Set tmpNode1 = oXML.createNode(1,cm_request_name,"")
				tmpNode.appendChild(tmpNode1)
				tmpNode1.Text="All"				
				requestText=oXML.xml
				Set oRoot=nothing
				Set oXML=nothing

				Set objXML=Server.CreateObject("MSXML2.serverXMLHTTP"&scXML)
				objXML.setTimeouts lResolve, lConnect, lSend, lReceive
				'objXML.onreadystatechange=getRef("state_Change")
				objXML.open "POST",ProductCartXMLServer,True
				objXML.setRequestHeader "XML-Agent", "ProductCart XML Partner"
				objXML.send(requestText)								
				pcv_CountTotal=pcv_CountTotal+1	
				If pcf_IsResponseGood()=False Then		
					Call UpdateExportFlag(0, tmpProductID, 0)
					pcv_strSummary = pcv_strSummary & "Product No. " & tmpProductID & " export failed with errors." & Chr(10)		
				Else					
					Set iXML1=Server.CreateObject("MSXML2.DOMDocument"&scXML)
					iXML1.async=false
					iXML1.load(objXML.responseXML)
					If (iXML1.parseError.errorCode <> 0) Then				
						Call UpdateExportFlag(0, tmpProductID, 0)
						ErrorCode=""
						pcv_strSummary = pcv_strSummary & "Product No. " & tmpProductID & " cannot be exported. " & iXML1.parseError.reason & Chr(10)
					Else
						pcv_CountCompleted=pcv_CountCompleted+1
						Set iRoot1 = iXML1.documentElement	
						ErrorCode = iRoot1.selectSingleNode(cm_requestStatus_name).text		
						pcv_strSummary = pcv_strSummary & "Product No. " & tmpProductID & " export successful." & Chr(10)								
					End if
					If ErrorCode="200" Then	
						ErrorCode=0
						tmpStream=iXML1.xml
						tmpStream1=split(tmpStream,"<" & cm_product & ">")
						tmpStream2=split(tmpStream1(1),"</" & cm_product & ">")
						tmpStream=tmpStream2(0)
						XMLStream=XMLStream & chr(9) & "<" & cm_product & ">" & tmpStream & "</" & cm_product & ">" & vbcrlf
					ElseIf ErrorCode<>"" Then
						Call UpdateExportFlag(0, tmpProductID, 0)
						pcv_CountCompleted=pcv_CountCompleted-1
						Set subNode=iRoot.selectSingleNode(cm_errorList_name)
						ErrorCode=subNode.selectSingleNode(cm_errorCode_name).Text
						ErrorDesc=subNode.selectSingleNode(cm_errorDesc_name).Text
						exit for
					End If						
				End If
				Set objXML=nothing
				Set iRoot1=nothing
				Set iXML1=nothing	
			end if
		Next
	end if
	
	if ErrorCode="0" and XMLStream<>"" then
		tmpGenName=Month(Date()) & Day(Date()) & Year(Date()) & Hour(Now()) & Minute(Now()) & Second(Now())
		tmpFileName="Products-" & tmpGenName & ".xml"
		XMLStream="<?xml version=""1.0""?>" & vbcrlf & "<" & cm_products & ">" & vbcrlf & XMLStream & "</" & cm_products & ">"
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		if PPD="1" then
			Set afi=fso.CreateTextFile(server.MapPath("/"&scPcFolder& "/xml/export/" & tmpFileName),True)
		else
			Set afi=fso.CreateTextFile(server.MapPath("..") & "\xml\export\" & tmpFileName,True)
		end if
		afi.Write(XMLStream)
		afi.Close
		Set afi=nothing
		Set fso=nothing
		if tmp_pFTPPartner>"0" then
			if PPD="1" then
				pPathXMLFile=server.MapPath("/"&scPcFolder& "/xml/export/" & tmpFileName)
			else
				pPathXMLFile=server.MapPath("..") & "\xml\export\" & tmpFileName
			end if
			call opendb()
			query="SELECT pcXP_FTPHost,pcXP_FTPDirectory,pcXP_FTPUsername,pcXP_FTPPassword FROM pcXMLPartners WHERE pcXP_ID=" & tmp_pFTPPartner & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				tmpFTPHost=rs("pcXP_FTPHost")
				tmpFTPDirectory=rs("pcXP_FTPDirectory")
				tmpFTPUsername=rs("pcXP_FTPUsername")
				tmpFTPPassword=rs("pcXP_FTPPassword")
				if tmpFTPPassword<>"" then
					tmpFTPPassword=enDeCrypt(tmpFTPPassword, scCrypPass)
				end if
			end if
			set rs=nothing
			call closedb()
			ErrorDesc=FTPUpload(tmpFTPHost, tmpFTPUsername, tmpFTPPassword, pPathXMLFile, tmpFTPDirectory)
			if ErrorDesc="OK" then
				ErrorCode=0
				ErrorDesc=""
			else
				ErrorCode=301 'Have FTP Errors
				If ErrorDesc="" Then
					ErrorDesc = "FTP Permission Denied: Script needs access to WScript.Shell."
				End If
			end if
			if ErrorCode=0 then
				if tmp_pRmvFile="1" then
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					Set afi = fso.GetFile(pPathXMLFile)
					afi.Delete
					Set afi=nothing
					Set fso=nothing
				end if
			end if
		end if
	end if
END IF
err.number=0
err.description=""
%><!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent" width="100%">
<tr>
	<td>
		
		<%if tmpFileName="" and ErrorCode>"0" then%>
        <div class="pcCPmessage">
			Cannot export XML data to a file!<br>
			Error Code: <%=ErrorCode%><br>
			Error Description: <%=ErrorDesc%>
        </div>
		<%else
		if tmpFileName<>"" and ErrorCode>"0" then%>
        <div class="pcCPmessageSuccess">
			Exported data successfully to the file: "<a href="../xml/export/<%=tmpFileName%>"><%=scPcFolder%>/xml/export/<%=tmpFileName%></a>"<br>
			<i>(Right click on the link and choose "Save Target As" to download file)</i>
        </div>
        <div class="pcCPmessage">
			Cannot upload this file to partner FTP Server!<br>
			Error Code: <%=ErrorCode%><br>
			Error Description: <%=ErrorDesc%>
        </div>
		<%else
		if tmpFileName<>"" and ErrorCode="0" and tmp_pFTPPartner>"0" and tmp_pRmvFile="1" then%>
        <div class="pcCPmessageSuccess">
			Exported data successfully to the file: "<%=tmpFileName%>"<br>
			This file has been uploaded to Partner FTP Server and removed from "<%=scPcFolder%>/xml/export" folder successfully!
        </div>
		<%else
		if tmpFileName<>"" and ErrorCode="0" then%>
        <div class="pcCPmessageSuccess">
			Exported data successfully to the file: "<a href="../xml/export/<%=tmpFileName%>"><%=scPcFolder%>/xml/export/<%=tmpFileName%></a>"<br>
			<%if tmp_pFTPPartner>"0" then%>
				This file has been uploaded to Partner FTP Server successfully!
			<%end if%>
			<i>(Right click on this link and choose "Save Target As" to download file)</i>
        </div>
		<%end if
		end if
		end if
		end if
		%>
		<% 
		if Clng(pcv_CountCompleted)<Clng(pcv_CountTotal) then %>
			<div class="pcCPmessage">
			<%=pcv_CountTotal-pcv_CountCompleted%> Product(s) may not have exported in the allowed time. Please try your export again, and select "No" under the heading &quot;Include/Exclude Previously Exported Products&quot;. View the &quot;Export Summary&quot; below and the &quot;Partner Logs&quot; for additional error reports.
			</div>
		<% end if %>
	</td>
</tr>
<%if Clng(pcv_CountCompleted)<Clng(pcv_CountTotal) then%>
<tr>
	<td>
	<div class="pcCPsectionTitle">Export Summary:</div>
	<textarea cols="70" rows="13"><%=pcv_strSummary%></textarea>		
	</td>
</tr>
<tr>
	<td colspan="2" class="pcSpacer">&nbsp;</td>
</tr>
<%end if%>
<tr>
	<td><input type="button" name="Back" value="XML Tools Manager" onclick="location='XMLToolsManager.asp';" class="ibtnGrey">&nbsp;
	<input type="button" name="back" value="Back" onClick="javascript:history.back()" class="ibtnGrey"></td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->