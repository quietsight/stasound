<%'Start SDBA - Send Low Inventory Notification
	strPath=Request.ServerVariables("PATH_INFO")
	dim iCnt, strPath,strPathInfo
	iCnt=0
	do while iCnt<2
		if mid(strPath,len(strPath),1)="/" then
			iCnt=iCnt+1
		end if
		if iCnt<2 then
			strPath=mid(strPath,1,len(strPath)-1)
		end if
	loop
	
	strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
				
	if Right(strPathInfo,1)="/" then
	else
		strPathInfo=strPathInfo & "/"
	end if
		
	query="SELECT DISTINCT idproduct,sku,serviceSpec,pcProd_ReorderLevel,description FROM Products WHERE removed=0 AND active=-1 AND configOnly=0 AND pcProd_NotifyStock=1 AND stock<pcProd_ReorderLevel AND pcProd_SentNotice=0;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=ConnTemp.execute(query)
	
	IF NOT rs.eof THEN
		pcArray=rs.GetRows()
		intCount=ubound(pcArray,2)
		set rs=nothing
	
		For i=0 to intCount
			pcv_idproduct=pcArray(0,i)
			pcv_sku=pcArray(1,i)
			pcv_BTO=pcArray(2,i)
			pcv_ReorderLevel=pcArray(3,i)
			pcv_prdName=pcArray(4,i)
		
			strURL=""
			strURL=strPathInfo & scAdminFolderName & "/login_1.asp?RedirectURL=" & Server.URLEnCode("FindProductType.asp?id=" & pcv_idproduct)
			strMsgBody=""
			strMsgBody=strMsgBody & pcv_prdName & " (product ID#" & pcv_idproduct & " - SKU #" & pcv_sku & ") needs to be restocked. The inventory level has dropped below the notification level of " & pcv_ReorderLevel & " units." & vbcrlf & vbcrlf
			strMsgBody=strMsgBody & "Use the link below to update inventory for this product when available:" & vbcrlf & vbcrlf
			strMsgBody=strMsgBody & strURL & VBCrlf & VBCrlf
			strMsgBody=strMsgBody & scCompanyName
				
			strMsgBodyMain=""&VBCrlf&VBCrlf&strMsgBody

			call sendmail(scCompanyName,scEmail,scFrmEmail,scCompanyName & " - " & "Low Inventory Notification: " & pcv_prdName & "(SKU: " & pcv_sku & ")",strMsgBodyMain)
			if err.number <> 0 then
				err.number=0
				err.description=""
			end if
	
			query="UPDATE products SET pcProd_SentNotice=1 WHERE idproduct=" & pcv_idproduct
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			set rs=nothing
		Next
	END IF
	set rs=nothing
'End SDBA - Send Low Inventory Notification%>
