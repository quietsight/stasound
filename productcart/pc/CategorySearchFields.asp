<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
Dim pcv_strCSFilters, pcv_strCSFieldQuery

'//////////////////////////////////////////////////////////////////////////////////////////////
'// START: CATEGORY SEARCH FIELDS
'//////////////////////////////////////////////////////////////////////////////////////////////
IF (lcase(pcStrPageName) = "showsearchresults.asp" AND SRCH_CSFRON="1") OR (lcase(pcStrPageName) = "viewcategories.asp" AND SRCH_CSFON="1")  THEN
	%>
    <link type="text/css" rel="stylesheet" href="pcSearchFields.css" />
    <link type="text/css" rel="stylesheet" href="../includes/spry/spryCollapsiblePanel-CSF.css" />
    <script type="text/javascript" src="../includes/spry/SpryDOMUtils.js"></script>
    <script type="text/javascript" src="../includes/spry/SpryCollapsiblePanel.js"></script>
    <%
    '//////////////////////////////////////////////////////////////////////////////////////////
    '// START: Get Widget Query Parameters
    '//////////////////////////////////////////////////////////////////////////////////////////
	Dim pcPageStyleCSF
    pcPageStyleCSF = LCase(getUserInput(Request("pageStyle"),1))

	'// Check querystring saved to session by 404.asp
	if pcPageStyleCSF = "" then
		strSeoQueryString=lcase(session("strSeoQueryString"))
		if strSeoQueryString<>"" then
			if InStr(strSeoQueryString,"pagestyle")>0 then
				pcPageStyleCSF=left(replace(strSeoQueryString,"pagestyle=",""),1)
			end if
		end if
	end if
	
    if pcPageStyleCSF = "" then
        pcPageStyleCSF = pStrPageStyle
    end if
    if isNULL(pcPageStyleCSF) OR trim(pcPageStyleCSF) = "" then
        pcPageStyleCSF = LCase(bType)
    end if
    if pcPageStyleCSF <> "h" and pcPageStyleCSF <> "l" and pcPageStyleCSF <> "m" and pcPageStyleCSF <> "p" then
        pcPageStyleCSF = LCase(bType)
    end if	

	'// SEO-START
	pcv_strCSFCatID=session("idCategoryRedirectSF")
	session("idCategoryRedirectSF")=""
	if pcv_strCSFCatID = "" then
		pcv_strCSFCatID=getUserInput(request("idCategory"),10)
	end if
	'// SEO-END
	
    if not validNum(pcv_strCSFCatID) then
        pcv_strCSFCatID=""
    end if
    pcv_strPage = getUserInput(Request("page"),10)
    if not validNum(pcv_strPage) then
        pcv_strPage=0
    end if
    
	If pcStrPageName = "showSearchResults.asp" Then
	
		SearchValues=getUserInput(Request("SearchValues"),0)
		pIdSupplier=getUserInput(request.querystring("idSupplier"),4)
		pPriceFrom=getUserInput(request.querystring("priceFrom"),20)
		pPriceUntil=getUserInput(request.querystring("priceUntil"),20)
		pSearchSKU=getUserInput(request.querystring("SKU"),150)
		IDBrand=getUserInput(request.querystring("IDBrand"),20)
		pKeywords=getUserInput(request.querystring("keyWord"),100)
		pcustomfield=getUserInput(request.querystring("customfield"),0)
		iPageSize=getUserInput(request("resultCnt"),10)
		strPrdOrd=getUserInput(request.querystring("order"),4)
		iPageCurrent=getUserInput(request.querystring("iPageCurrent"),4)
		tKeywords=pKeywords
		tIncludeSKU=getUserInput(request.querystring("includeSKU"),10)
		if tIncludeSKU = "" then
			tIncludeSKU = "true"
		end if
		
		if Instr(pPriceFrom,",")>Instr(pPriceFrom,".") then
			pPriceFrom=replace(pPriceFrom,",",".")
		end if
		if NOT isNumeric(pPriceFrom) then
			pPriceFrom=0
		end if
		if Instr(pPriceUntil,",")>Instr(pPriceUntil,".") then
			pPriceUntil=replace(pPriceUntil,",",".")
		end if
		if NOT isNumeric(pPriceUntil) then
			pPriceUntil=999999999
		end if
		if NOT validNum(pIdSupplier) or trim(pIdSupplier)="" then
			pIdSupplier=0
		end if
		pWithStock=getUserInput(request.querystring("withStock"),2)
		if NOT validNum(IDBrand) or trim(IDBrand)="" then
			IDBrand=0
		end if
		
		if NOT validNum(strPrdOrd) or trim(strPrdOrd)="" then
			strPrdOrd=3
		end if
		Select Case strPrdOrd
			Case "1": strORD1="A.idproduct ASC"
			Case "2": strORD1="A.sku ASC, A.idproduct DESC"
			Case "3": strORD1="A.description ASC"
			Case "4":
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) ASC"
					else
						strORD1 = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) ASC"
					end if
				else
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) ASC"
					else
						strORD1 = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) ASC"
					end if
				End if
			Case "5": 
				If Session("customerType")=1 then
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN (CASE A.bToBPrice WHEN 0 THEN A.Price ELSE A.bToBPrice END) ELSE (CASE (CASE WHEN A.pcProd_BTODefaultWPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultWPrice END) WHEN 0 THEN A.pcProd_BTODefaultPrice ELSE A.pcProd_BTODefaultWPrice END) END) DESC"
					else
						strORD1 = "(iif(iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),iif(IsNull(A.pcProd_BTODefaultPrice),0,A.pcProd_BTODefaultPrice),A.pcProd_BTODefaultWPrice)=0,iif(A.btoBPrice=0,A.Price,A.btoBPrice),iif((A.pcProd_BTODefaultWPrice=0) OR (IsNull(A.pcProd_BTODefaultWPrice)),A.pcProd_BTODefaultPrice,A.pcProd_BTODefaultWPrice))) DESC"
					end if
				else
					if Ucase(scDB)="SQL" then
						strORD1 = "(CASE (CASE WHEN A.pcProd_BTODefaultPrice IS NULL THEN 0 ELSE A.pcProd_BTODefaultPrice END) WHEN 0 THEN A.Price ELSE A.pcProd_BTODefaultPrice END) DESC"
					else
						strORD1 = "(iif((A.pcProd_BTODefaultPrice=0) OR (IsNull(A.pcProd_BTODefaultPrice)),A.Price,A.pcProd_BTODefaultPrice)) DESC"
					end if
				End if
		End Select
		strORD=strPrdOrd
		
		intExact=getUserInput(request.querystring("exact"),4)
		if NOT validNum(intExact) or trim(intExact)="" then
			intExact=0
		end if
		
	End If '// If pcStrPageName = "search.asp" Then
    
    pcs_CSFSetVariables()
    '//////////////////////////////////////////////////////////////////////////////////////////
    '// END: Get Widget Query Parameters
    '//////////////////////////////////////////////////////////////////////////////////////////
    
    
    
    '//////////////////////////////////////////////////////////////////////////////////////////
    '// START: Disply Widget 
    '//////////////////////////////////////////////////////////////////////////////////////////
    If pcv_strCSFCatID<>"" Then  
	
		'// Get the list of products that are currently available
		'// Only run this block if the include is in the header
		if len(pcv_strCValues)>0 AND len(pcv_strCSFilters)=0 then
		
			tmpStrEx3=""
			pcv_HavingCount2 = 0
			tmpSValues3=split(pcv_strCValues,"||")
			For k=lbound(tmpSValues3) to ubound(tmpSValues3)	
				if tmpSValues3(k)<>"" then
					if pcv_HavingCount2=0 then
						tmpStrEx3 = tmpStrEx3 & ""& tmpSValues3(k)
					else
						tmpStrEx3 = tmpStrEx3 & ","& tmpSValues3(k)
					end if 					
					pcv_HavingCount2 = pcv_HavingCount2 + 1										
				end if
			Next
			
			queryCSF = "SELECT pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "FROM pcSearchFields_Products "
			queryCSF = queryCSF & "INNER JOIN products ON products.idProduct=pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "INNER JOIN categories_products ON products.idProduct=categories_products.idProduct "
			queryCSF = queryCSF & "WHERE pcSearchFields_Products.idSearchData in (" & tmpStrEx3 & ") "
			queryCSF = queryCSF & "AND categories_products.idCategory="& pcv_strCSFCatID &" AND active=-1 AND configOnly=0 AND removed=0 "
			queryCSF = queryCSF & "GROUP BY pcSearchFields_Products.idProduct "
			queryCSF = queryCSF & "HAVING COUNT(DISTINCT pcSearchFields_Products.idSearchData) = " & pcv_HavingCount2

			set rsCSF=Server.CreateObject("ADODB.Recordset")  
			set rsCSF=connTemp.execute(queryCSF)
			if NOT rsCSF.eof then
				ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
				CartProductIdString = Join(ProductIdArray,",")
				pcv_strCSFilters = " AND (products.idProduct In ("& CartProductIdString &"))"
			else 
				pcv_strCSFilters = " AND (products.idProduct In (0))"
			end if
			set rsCSF = nothing

		end if

		pcv_strTmpCatID = pcv_strCSFCatID
		
		TmpCatList=""
		If pcStrPageName = "showSearchResults.asp" Then
		  if pIdCategory<>"0" then
			  if (schideCategory = "1") OR (SRCH_SUBS = "1") then	
			  	  TmpCatList=""				
				  call pcs_GetSubCats(pIdCategory) '// get sub cats
				  TmpCatList=pIdCategory&TmpCatList
				  pcv_strTmpCatID = TmpCatList
			  end if
		  end if
		End If
		
		query="SELECT DISTINCT idSearchData from pcSearchFields_Categories where idCategory IN (" & pcv_strTmpCatID & ") "
        set rs=server.CreateObject("ADODB.RecordSet")
        set rs=conlayout.execute(query)
        pcv_strSearchDataIDs = ""
        do while not rs.eof
            pcv_strSearchDataIDs=pcv_strSearchDataIDs & rs("idSearchData") & ","
            rs.movenext
        loop 
        set rs=nothing
        
        If pcv_strSearchDataIDs<>"" Then
        %>    
        <input type="hidden" name="customfield" id="customfield" value="0">  
        <input type="hidden" name="SearchValues" id="SearchValues" value="">   
        <%
        pcv_strTmpPageName = request.ServerVariables("SCRIPT_NAME")
        If NOT len(pcv_strTmpPageName)>0 OR instr(pcv_strTmpPageName,"404.asp")<>0  Then ' // SEO EDIT
            pcv_strTmpPageName="viewCategories.asp"
        End If
        %>
        <form name="CSF" id="CSF" action="<%=pcv_strTmpPageName%>" method="get"> 
        <div id="pcCSF">		
            <h3><%=dictLanguage.Item(Session("language")&"_categorysearchfield_1")%></h3>
            <span id="notice" name="notice" class="pcCSFNotice" style="display:none"><%=dictLanguage.Item(Session("language")&"_categorysearchfield_4")%></span>
            <span id="stable" name="stable" class="pcCSFStable"></span> 
            
            <% If pcStrPageName = "showSearchResults.asp" Then %>
                <%' Required Search Parameters %>  
                <input type="hidden" name="SKU" id="SKU" value="<%=pSearchSKU%>"> 
                <input type="hidden" name="keyWord" id="keyWord" value="<%=pKeywords%>"> 
                <input type="hidden" name="SearchValues" id="SearchValues" value="<%=SearchValues%>"> 
                <input type="hidden" name="includeSKU" id="includeSKU" value="<%=tIncludeSKU%>"> 
                <input type="hidden" name="priceFrom" id="priceFrom" value="<%=pPriceFrom%>">
                <input type="hidden" name="priceUntil" id="priceUntil" value="<%=pPriceUntil%>">
                <input type="hidden" name="idSupplier" id="idSupplier" value="<%=pIdSupplier%>">
                <input type="hidden" name="withStock" id="withStock" value="<%=pWithStock%>">
                <input type="hidden" name="customfield" id="customfield" value="<%=pcustomfield%>">
                <input type="hidden" name="IDBrand" id="IDBrand" value="<%=IDBrand%>">
                <input type="hidden" name="order" id="order" value="<%=strPrdOrd%>">
                <input type="hidden" name="exact" id="exact" value="<%=intExact%>">
                <input type="hidden" name="resultCnt" id="resultCnt" value="<%=iPageSize%>">
                <input type="hidden" name="iPageSize" id="iPageSize" value="<%=iPageSize%>">
                <input type="hidden" name="iPageCurrent" id="iPageCurrent" value="<%=iPageCurrent%>">
    		<% end if %>
            
            <%' Required Category Parameters %>
            <input type="hidden" name="SFID" id="SFID" value="<%=pcv_intSFID%>">
            <input type="hidden" name="SFNAME" id="SFNAME" value="<%=pcv_strSFNAME%>">
            <input type="hidden" name="SFVID" id="SFVID" value="<%=pcv_strCValues%>">
            <input type="hidden" name="SFVALUE" id="SFVALUE" value="<%=pcv_strCValues%>">
            <input type="hidden" name="SFCount" id="SFCount" value="<%=pcv_intSFCount%>">         
            <input type="hidden" name="page" id="page" value="<%=pcv_strPage%>">
            <input type="hidden" name="pageStyle" id="pageStyle" value="<%=pcPageStyleCSF%>"> 
            <input type="hidden" name="idcategory" id="idcategory" value="<%=pcv_strCSFCatID%>"> 
            <%
            pcv_strSearchDataIDs=left(pcv_strSearchDataIDs,len(pcv_strSearchDataIDs)-1)
			query="SELECT DISTINCT idSearchField, pcSearchFieldName, pcSearchFieldShow, pcSearchFieldOrder "
			query=query&"FROM pcSearchFields "
			query=query&"WHERE idSearchField IN ("& pcv_strSearchDataIDs &") "
			query=query&"ORDER BY pcSearchFieldOrder ASC, pcSearchFieldName ASC;"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=conlayout.execute(query)
			if not rs.eof then			
                pcArray=rs.getRows()
                intCount=ubound(pcArray,2)
                set rs=nothing	
                pcv_ClearAll = 0
				pcv_strSpryVars = 0
				pcv_strSpryVarArray = ""	
				pcv_intSpryCount = 0
                For i=0 to intCount
                    pcv_strCatSearchName = pcArray(1,i)  
                    pcv_strCatSearchID = pcArray(0,i)
                    pcv_strFullView=Request("VS"&pcv_strCatSearchID) ' note: getUserInput cannot be used here
                    If NOT validNum(pcv_strFullView) Then
                        pcv_strFullView=0
                    End If
					If len(CartProductIdString)>0 Then
						query="SELECT DISTINCT pcSearchData.idSearchData, pcSearchData.pcSearchDataName, pcSearchData.idSearchField, pcSearchData.pcSearchDataOrder "
						query=query&"FROM pcSearchData "
						query=query&"INNER JOIN pcSearchFields_Products ON pcSearchFields_Products.idSearchData = pcSearchData.idSearchData "
						query=query&"WHERE pcSearchFields_Products.idProduct IN ("& CartProductIdString &") "
						query=query&"AND pcSearchData.idSearchField=" & pcArray(0,i) & " "
						query=query&"ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
					Else
						query="SELECT DISTINCT A.idSearchData, A.pcSearchDataName, A.idSearchField, A.pcSearchDataOrder "
						query=query&"FROM pcSearchData A "
						query=query&"LEFT JOIN pcSearchFields_Products B on B.idSearchData=A.idSearchData "
						query=query&"LEFT JOIN  "
						query=query&"	( "
						query=query&"		SELECT D.idProduct, D.idCategory "
						query=query&"		FROM categories_products D "
						query=query&"		LEFT JOIN products E ON E.idProduct=D.idProduct "
						query=query&"		WHERE E.active=-1 AND E.configOnly=0 and E.removed=0 "
						query=query&"	) C on C.idProduct=B.idProduct "
						query=query&"WHERE C.idCategory in ("& pcv_strTmpCatID &") AND idSearchField=" & pcArray(0,i) & " "
						query=query&"ORDER BY A.pcSearchDataOrder ASC, A.pcSearchDataName ASC;"
					End If
                    set rs=Server.CreateObject("ADODB.Recordset")
                    set rs=conlayout.execute(query)
                    if not rs.eof then
                        tmpArr=rs.getRows()
                        LCount=ubound(tmpArr,2) 
                        set rs=nothing				
                        pcv_strGroupValues = pcf_ColumnToArray(tmpArr,0)
                        pcv_strGroupValues = Join(pcv_strGroupValues,"||") 
						if LCount>=0 then 
						pcv_intSpryCount=pcv_intSpryCount+1
                        %>
                        <div id="CollapsiblePanel<%=pcv_intSpryCount%>" class="CollapsiblePanel">      
                            <div class="CollapsiblePanelTab">
                                <div class="pcCSFTitle"><%=pcv_strCatSearchName%></div>
                                <input type="hidden" name="VS<%=pcv_strCatSearchID%>" id="VS<%=pcv_strCatSearchID%>" value="<%=pcv_strFullView%>">
                            </div>
                            <%
                            pcv_strFullViewClass = ""														
                            if pcv_strFullView = 1 then
                                pcv_strFullViewClass = "style=' height: 200px; overflow: auto'"
                            end if
                            %>
                            <div class="CollapsiblePanelContent">
                                <div id="FullContainer" <%=pcv_strFullViewClass%>> 
                                    <%
                                    Dim pcv_strNotInUse
                                    pcv_strInUse = 0
                                    pcv_CountGroup = 0
                                    For j=0 to LCount
                                        pcv_strInUse = 0
                                        pcv_strDataID = tmpArr(0,j)
                                        pcv_strSFID = tmpArr(2,j)
                                        pcv_strCatSearchDataName = tmpArr(1,j)
										pcv_strInUse = pcf_InUse(pcv_strCValues,pcv_strDataID)	
										If pcv_strInUse=0 Then                                        
                                            if pcStrPageName = "showSearchResults.asp" then
												pcv_intPrdCount = pcf_CountSearchResults(pcv_strDataID, pcv_strGroupValues)
											else
												pcv_intPrdCount = pcf_CountResults(pcv_strDataID, pcv_strGroupValues)										
											end if											
                                            If pcv_intPrdCount>0 Then
                                                pcv_strSearchValue = pcv_strCValues & pcv_strDataID & "||"
                                                pcv_strCatSearchName = replace(pcv_strCatSearchName,"""","&quot;")
                                                pcv_strCatSearchNameTmp = pcf_SanitizeJava(pcv_strCatSearchName)											
                                                pcv_strCatSearchDataName = replace(pcv_strCatSearchDataName,"""","&quot;")
                                                pcv_strCatSearchDataNameTmp = pcf_SanitizeJava(pcv_strCatSearchDataName)											
                                                %>
                                                <div class="pcCSFItem">
                                                    <a href="javascript:AddSF('<%=pcArray(0,i)%>','<%=pcv_strCatSearchNameTmp%>','<%=pcv_strDataID%>','<%=pcv_strCatSearchDataNameTmp%>',0);">
                                                        <img src="../pc/images/plus.jpg" alt="Add" border="0"> <%=pcv_strCatSearchDataName%>
                                                    </a>
                                                    <span class="pcCSFCount">(<%=pcv_intPrdCount%>)</span>
                                                </div>
                                                <%										
                                                pcv_ClearAll = pcv_ClearAll + 1 '// Count total filters
                                                pcv_CountGroup = pcv_CountGroup + 1 '// Count filters in group
                                            End If
                                        End If
                                        if pcv_strFullView = 0 then
                                            if pcv_CountGroup > 5 then
                                                %>
                                                <div class="pcCSFItem">
                                                    <a href="javascript:ShowMore(<%=pcv_strCatSearchID%>)"><strong><%=dictLanguage.Item(Session("language")&"_categorysearchfield_2")%></strong></a>
                                                </div>
                                                <%
                                                exit for
                                            end if
                                        end if
                                    Next
                                    if pcv_CountGroup = 0 then
                                        %>
                                        <script type="text/javascript">
                                           document.getElementById("CollapsiblePanel<%=pcv_intSpryCount%>").style.display = 'none'; 
                                        </script>
                                        <%
                                    end if
                                    %>
                                </div>
                                <%
                                if pcv_strFullView = 1 then
                                    if pcv_CountGroup > 5 then
                                        %>
                                        <div class="pcCSFItem">
                                            <a href="javascript:ShowLess(<%=pcv_strCatSearchID%>)"><strong><%=dictLanguage.Item(Session("language")&"_categorysearchfield_3")%></strong></a>
                                        </div>
                                        <%
                                    end if
                                end if
                                %>                            
                            </div>
          </div>
                        <%
						'// Create JavaScript Strings to Initialize the Panels
						pcv_strSpryVars = pcv_strSpryVars + 1
						pcv_strSpryVarArray = pcv_strSpryVarArray & "CollapsiblePanel" & pcv_intSpryCount & ".isOpen(),"						

                        end if '// if LCount>0 then
    
                    end if '// if not rs.eof then
    
                Next '// For i=0 to intCount
                
            end if
            if pcv_ClearAll = 0 AND pcv_strSFNAME="" then
                %>
                <script type="text/javascript">
                    document.getElementById("pcCSF").style.display = 'none'; 
                </script>
                <%
            end if
            %>
        </div>
        </form>
        <!-- search custom fields if any are defined -->
        <%
        tmpJSStr=""
        if pcv_intSFID<>"" then
            tmpJSStr=tmpJSStr & "var SFID=new Array(" & pcf_converToArray(pcv_intSFID) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
        end if
        if pcv_strSFNAME<>"" then
            tmpJSStr=tmpJSStr & "var SFNAME=new Array(" & pcf_converToArray(pcf_SanitizeJava(pcv_strSFNAME)) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf	
        end if
        if pcv_strCValues<>"" then
            tmpJSStr=tmpJSStr & "var SFVID=new Array(" & pcf_converToArray(pcv_strCValues) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
        end if
        if pcv_strSFVALUE<>"" then
            tmpJSStr=tmpJSStr & "var SFVALUE=new Array(" & pcf_converToArray(pcf_SanitizeJava(pcv_strSFVALUE)) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
        end if
        if SFVORDER<>"" then
            tmpJSStr=tmpJSStr & "var SFVORDER=new Array(" & pcf_converToArray(SFVORDER) & ")" & vbcrlf
        else
            tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
        end if
        if pcv_intSFCount<>"" then
            intCount=pcv_intSFCount
        else
            intCount=-1
        end if	        
        tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf
        %>
        <script type="text/javascript">
            <%=tmpJSStr%>
            function CreateTable(tmpRun)
            {
                var tmp1="";
                var tmp2="";
                var tmp3="";
                var i=0;
                var found=0;			
                tmp1='<table style="width: 100%; margin: 5px 0 5px 0;">';
                for (var i=0;i<=SFCount;i++)
                {
                    found=1;
                    tmp1=tmp1 + '<tr><td style="text-align: right;"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="" border="0"></a></td><td style="text-align: left; width: 100%;">'+SFNAME[i]+': '+SFVALUE[i]+'</td></tr>';
                    if (tmp2=="") tmp2=tmp2 + "||";
                    tmp2=tmp2 + SFID[i] + "||";
                    if (tmp3=="") tmp3=tmp3 + "||";
                    tmp3=tmp3 + SFVID[i] + "||";
                }
                tmp1=tmp1+'</table>';
                if (found==0) tmp1="";
                document.getElementById("stable").innerHTML=tmp1;
                if (tmp2=="") tmp2=0;
                document.getElementById("customfield").value=tmp2;
                document.getElementById("SearchValues").value=tmp3;
                if (tmp2==0)
                {
                    document.getElementById("customfield").value=0;
                    document.getElementById("SearchValues").value='';
                }
                if (tmpRun!=1) document.CSF.submit();
                
            }
    
            function ClearSF(tmpSFID)
            {
                var i=0;
                for (var i=0;i<=SFCount;i++)
                {
                    if (SFID[i]==tmpSFID)
                    {
                        removedArr = SFID.splice(i,1);
                        removedArr = SFNAME.splice(i,1);
                        removedArr = SFVID.splice(i,1);
                        removedArr = SFVALUE.splice(i,1);
                        removedArr = SFVORDER.splice(i,1);
                        SFCount--;
                        break;
                    }
                }
                document.getElementById("SFID").value=SFID.join("||");
                document.getElementById("SFNAME").value=SFNAME.join("||");
                document.getElementById("SFVID").value=SFVID.join("||");
                document.getElementById("SFVALUE").value=SFVALUE.join("||");
                document.getElementById("SFCount").value=SFCount;
                document.getElementById("notice").style.display = ''; 
                CreateTable(0);
            }
    
            function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
            {
                if ((tmpSVID!="") && (tmpSFID!="") && (tmpSVID!="0") && (tmpSFID!="0"))
                {
                    var i=0;
                    var found=0;
                    for (var i=0;i<=SFCount;i++)
                    {
                        if (SFID[i]==tmpSFID)
                        {
                            SFVID[i]=tmpSVID;
                            SFVALUE[i]=tmpSValue;
                            SFVORDER[i]=tmpSOrder;
                            found=1;
                            break;
                        }
                    }
                    if (found==0)
                    {
                        SFCount++;
                        SFID[SFCount]=tmpSFID;
                        SFNAME[SFCount]=tmpSFName;
                        SFVID[SFCount]=tmpSVID;
                        SFVALUE[SFCount]=tmpSValue;
                        SFVORDER[SFCount]=tmpSOrder;		
    
                    }
                    document.getElementById("SFID").value=SFID.join("||");
                    document.getElementById("SFNAME").value=SFNAME.join("||");
                    document.getElementById("SFVID").value=SFVID.join("||");
                    document.getElementById("SFVALUE").value=SFVALUE.join("||");
                    document.getElementById("SFCount").value=SFCount;
                    document.getElementById("notice").style.display = ''; 
                    CreateTable(0);
                }
            }  
    
            function ShowMore(VSID)
            {
                document.getElementById("VS"+VSID).value=1;
				document.getElementById("SFID").value=SFID.join("||");
				document.getElementById("SFNAME").value=SFNAME.join("||");
				document.getElementById("SFVID").value=SFVID.join("||");
				document.getElementById("SFVALUE").value=SFVALUE.join("||");
				document.getElementById("SFCount").value=SFCount;
				document.getElementById("notice").style.display = ''; 
				CreateTable(0);
            } 
    
            function ShowLess(VSID)
            {
                document.getElementById("VS"+VSID).value=0;
				document.getElementById("SFID").value=SFID.join("||");
				document.getElementById("SFNAME").value=SFNAME.join("||");
				document.getElementById("SFVID").value=SFVID.join("||");
				document.getElementById("SFVALUE").value=SFVALUE.join("||");
				document.getElementById("SFCount").value=SFCount;
				document.getElementById("notice").style.display = ''; 
				CreateTable(0);
            } 
    
            CreateTable(1);   
    
        </script>
        <%
        End If '// If pcv_strSearchDataIDs<>"" Then
        
    End If
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: Disply Widget 
	'//////////////////////////////////////////////////////////////////////////////////////////





	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcf_converToArray
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Converts the query string into a JavaScript array
	'// Params:		A search array separated by pipes
	'// Returns:	A Javascript array separated by commas and encapsulated with (')  
	Function pcf_converToArray(pipes)
		On Error Resume Next
		pcf_converToArray = pipes
		pcf_converToArray=replace(pcf_converToArray,"||","','")
		pcf_converToArray="'"&pcf_converToArray&"'"
	End Function
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_converToArray
	'//////////////////////////////////////////////////////////////////////////////////////////
	
	
	
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcf_CountResults & pcf_CountSearchResults
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Counts the number of products that will be returned when the filter is added
	'// Params:		SearchValueId - a string of search value ids separated by two pipes
	'// Returns:	Boolean  
	Function pcf_CountResults(SearchValueId,GroupValues)
		On Error Resume Next
		
		tmpStrEx = SearchValueId '// ""
		pcv_HavingCount = 1
		if len(pcv_strCValues)>0 then		
			tmpSValues=split(pcv_strCValues,"||")
			For k=0 to ubound(tmpSValues)		
				if tmpSValues(k)<>"" AND pcf_InUse(GroupValues,tmpSValues(k))=0 then
					tmpStrEx = tmpStrEx & ","& tmpSValues(k)
					pcv_HavingCount = pcv_HavingCount + 1	
				end if
			Next	
		end if	

		queryCSF = "SELECT pcSearchFields_Products.idProduct "
		queryCSF = queryCSF & "FROM pcSearchFields_Products "
		queryCSF = queryCSF & "INNER JOIN products ON products.idProduct=pcSearchFields_Products.idProduct "
		queryCSF = queryCSF & "INNER JOIN categories_products ON products.idProduct=categories_products.idProduct "
		queryCSF = queryCSF & "WHERE pcSearchFields_Products.idSearchData in (" & tmpStrEx & ") "
		queryCSF = queryCSF & "AND categories_products.idCategory="& pcv_strCSFCatID &" AND active=-1 AND configOnly=0 AND removed=0 "
		queryCSF = queryCSF & "GROUP BY pcSearchFields_Products.idProduct "
		queryCSF = queryCSF & "HAVING COUNT(DISTINCT pcSearchFields_Products.idSearchData) = " & pcv_HavingCount

		set rsCSF=Server.CreateObject("ADODB.Recordset")  
		set rsCSF=conlayout.execute(queryCSF)
		if NOT rsCSF.eof then
			pcarray_RowCount = rsCSF.GetRows
			pcf_CountResults = UBound(pcarray_RowCount, 2) + 1
		else 
			pcf_CountResults = 0
		end if 	
		set rsCSF=nothing

	End Function
	
	Function pcf_CountSearchResults(SearchValueId,GroupValues)
		On Error Resume Next				
		tmpStrEx = ""
		if len(pcv_strCValues)>0 then
			queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData>0 "			
			tmpSValues=split(pcv_strCValues,"||")
			For k=0 to ubound(tmpSValues)		
				if tmpSValues(k)<>"" AND pcf_InUse(GroupValues,tmpSValues(k))=0 then					
					SubQuery = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData = " & tmpSValues(k) & ""
					set rsSubQuery=Server.CreateObject("ADODB.Recordset")
					set rsSubQuery=conlayout.execute(SubQuery)
					If NOT rsSubQuery.eof Then
						ProductIdArray = pcf_ColumnToArray(rsSubQuery.getRows(),0)
						ProductIdString = Join(ProductIdArray,",")
						tmpStrEx=tmpStrEx & " AND pcSearchFields_Products.idProduct IN "
						tmpStrEx=tmpStrEx & "(" & ProductIdString & ")"	
					End If
					set rsSubQuery = nothing		
				end if
			Next	
			tmpStrEx = tmpStrEx & " AND pcSearchFields_Products.idSearchData=" & SearchValueId & ""
			queryCSF = queryCSF & tmpStrEx
		else
			queryCSF = "SELECT pcSearchFields_Products.idProduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData="& SearchValueId &" "
		end if		
		set rsCSF=Server.CreateObject("ADODB.Recordset")  
		set rsCSF=conlayout.execute(queryCSF)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsCSF=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if NOT rsCSF.eof then
			ProductIdArray = pcf_ColumnToArray(rsCSF.getRows(),0)
			ProductIdString = Join(ProductIdArray,",")
			pcv_strCSFilters = " AND (A.idProduct In ("& ProductIdString &"))"
		else 
			pcv_strCSFilters = " AND (A.idProduct In (0))"
		end if 	
		set rsCSF=nothing

		'*******************************
		' Create Search Query
		'*******************************
		tmp_StrQuery=""
		if session("customerCategory")="" or session("customerCategory")=0 then
			If session("customerType")=1 then
				tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultWPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultWPrice<=" &pPriceUntil&")"
			else
				tmp_StrQuery="(A.serviceSpec<>0 AND A.pcProd_BTODefaultPrice>="&pPriceFrom&" AND A.pcProd_BTODefaultPrice<=" &pPriceUntil&")"
			end if
		else
			tmp_StrQuery="(A.serviceSpec<>0 AND A.idproduct IN (SELECT DISTINCT idproduct FROM pcBTODefaultPriceCats WHERE pcBTODefaultPriceCats.idCustomerCategory=" & session("customerCategory") & " AND pcBTODefaultPriceCats.pcBDPC_Price>="&pPriceFrom&" AND pcBTODefaultPriceCats.pcBDPC_Price<=" &pPriceUntil&"))"
		end if
		
		if scDB="Access" then
			zSQL="A.sDesc"
		else
			zSQL="cast(A.sDesc as varchar(8000)) sDesc"
		end if
		
		pcv_strMaxResults=SRCH_MAX
		If pcv_strMaxResults>"0" Then
			pcv_strLimitPhrase="TOP " & pcv_strMaxResults
		Else
			pcv_strLimitPhrase=""
		End If
		
		strSQL= "SELECT "& pcv_strLimitPhrase &" A.idProduct, A.sku, A.description, A.price, A.listHidden, A.listPrice, A.serviceSpec, A.bToBPrice, A.smallImageUrl, A.noprices, A.stock, A.noStock, A.pcprod_HideBTOPrice, A.pcProd_BackOrder, A.FormQuantity, A.pcProd_BackOrder, A.pcProd_BTODefaultPrice, "& zSQL &" " 
		strSQL=strSQL& "FROM products A "
		strSQL=strSQL& " WHERE (A.active=-1 AND A.removed=0 AND A.idProduct IN (" 

			'// START: Category Sub-Query
			strSQL=strSQL& "SELECT B.idProduct FROM categories_products B INNER JOIN categories C ON "
			strSQL=strSQL & "C.idCategory=B.idCategory WHERE C.iBTOhide=0 "
			if pIdCategory<>"0" then
				if (schideCategory = "1") OR (SRCH_SUBS = "1") then					
					TmpCatList=""
					call pcs_GetSubCats(pIdCategory) '// get sub cats
					TmpCatList = pIdCategory&TmpCatList
					if len(TmpCatList)>0 then
						strSQL=strSQL & " AND B.idCategory IN ("& TmpCatList &")" '// include sub cats
					else
						strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
					end if
				else
					strSQL=strSQL & " AND B.idCategory=" &pIdCategory	
				end if
			end if
			if session("CustomerType")<>"1" then
				strSQL=strSQL & " AND C.pccats_RetailHide=0"
			end if
			'// END: Category Sub-Query
		
		strSQL=strSQL& ") AND (" & tmp_StrQuery & " OR (A.serviceSpec=0 AND A.configOnly=0 AND A.price>="&pPriceFrom&" AND A.price<=" &pPriceUntil&")) " 
		
		if len(pSearchSKU)>0 then
			strSQL=strSQL & " AND A.sku like '%"&pSearchSKU&"%'"
		end if
		
		if pIdSupplier<>"0" then
			strSQL=strSQL & " AND A.idSupplier=" &pIdSupplier
		end if
		
		if pWithStock="-1" then
			strSQL=strSQL & " AND (A.stock>0 OR A.noStock<>0)" 
		end if
		
		if (IDBrand&""<>"") and (IDBrand&""<>"0") then
			strSQL=strSQL & " AND A.IDBrand=" & IDBrand
		end if
		
		TestWord=""
		if intExact<>"1" then
			if Instr(pKeywords," AND ")>0 then
				keywordArray=split(pKeywords," AND ")
				TestWord=" AND "
			else
				if Instr(pKeywords," and ")>0 then
					keywordArray=split(pKeywords," and ")
					TestWord=" AND "
				else
					if Instr(pKeywords,",")>0 then
						keywordArray=split(pKeywords,",")
						TestWord=" OR "
					else
						if (Instr(pKeywords," OR ")>0) then
							keywordArray=split(pKeywords," OR ")
							TestWord=" OR "
						else
							if (Instr(pKeywords," or ")>0) then
								keywordArray=split(pKeywords," or ")
								TestWord=" OR "
							else
								if (Instr(pKeywords," ")>0) then
									keywordArray=split(pKeywords," ")
									TestWord=" AND "
								else
									keywordArray=split(pKeywords,"***")	
									TestWord=" OR "
								end if
							end if
						end if
					end if
				end if
			end if
		else
			pKeywords=trim(pKeywords)
			if pKeywords<>"" then
				if scDB="SQL" then
					pKeywords="'" & pKeywords & "'***'%[^a-zA-z0-9]" & pKeywords & "[^a-zA-z0-9]%'***'" & pKeywords & "[^a-zA-z0-9]%'***'%[^a-zA-z0-9]" & pKeywords & "'"
				else
					pKeywords="'" & pKeywords & "'***'%[!a-zA-z0-9]" & pKeywords & "[!a-zA-z0-9]%'***'" & pKeywords & "[!a-zA-z0-9]%'***'%[!a-zA-z0-9]" & pKeywords & "'"
				end if
			end if
			keywordArray=split(pKeywords,"***")	
			TestWord=" OR "
		end if
		
		tmpStrEx=""
		if pCValues<>"" AND pCValues<>"0" then
			tmpSValues=split(pCValues,"||")
			For k=lbound(tmpSValues) to ubound(tmpSValues)
				if tmpSValues(k)<>"" then	
					sfquery=""
					sfquery = "SELECT pcSearchFields_Products.idproduct FROM pcSearchFields_Products WHERE pcSearchFields_Products.idSearchData=" & tmpSValues(k)
					set rsSearchFields=Server.CreateObject("ADODB.Recordset")
					set rsSearchFields=connTemp.execute(sfquery)
					If NOT rsSearchFields.eof Then
						SearchFieldArray = pcf_ColumnToArray(rsSearchFields.getRows(),0)
						SearchFieldString = Join(SearchFieldArray,",")		
						If len(SearchFieldString)>0 Then
							tmpStrEx=tmpStrEx & " AND A.idproduct IN ("& SearchFieldString &")"
						End If
					End If
					set rsSearchFields = nothing
				end if
			Next
		end if
		
		'// Category Seach Fields 
		tmpStrEx = tmpStrEx & pcv_strCSFilters
		
		IF intExact<>"1" THEN
		
			if pKeywords<>"" then
			
				strSQl=strSql & " AND ("
				
				tmpSQL="(A.details LIKE "
				tmpSQL2="(A.description LIKE "
				tmpSQL3="(A.sDesc LIKE "
				if tIncludeSKU="true" then
					tmpSQL4="(A.SKU LIKE "
				end if
				Dim Pos
				Pos=0
				For L=LBound(keywordArray) to UBound(keywordArray)
					if trim(keywordArray(L))<>"" then
					Pos=Pos+1
					if Pos>1 Then
						tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
						tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
						tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
						if tIncludeSKU="true" then
							tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
						end if
					end if
						tmpSQL=tmpSQL  & "'%" & trim(keywordArray(L)) & "%'"
						tmpSQL2=tmpSQL2 & "'%" & trim(keywordArray(L)) & "%'"
						tmpSQL3=tmpSQL3 & "'%" & trim(keywordArray(L)) & "%'"
						if tIncludeSKU="true" then
							tmpSQL4=tmpSQL4 & "'%" & trim(keywordArray(L)) & "%'"
						end if
					end if
				Next
				tmpSQL=tmpSQL & ")"
				tmpSQL2=tmpSQL2 & ")"
				tmpSQL3=tmpSQL3 & ")"
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & ")"
				end if
				
				strSQL=strSQL & tmpSQL
				strSQL=strSQL & " OR " & tmpSQL2
				if tIncludeSKU="true" then
					strSQL=strSQL & " OR " & tmpSQL3
					strSQL=strSQL & " OR " & tmpSQL4 & ")"
				else	
					strSQL=strSQL & " OR " & tmpSQL3 & ")"
				end if
				strSQL=strSQL& ")" & tmpStrEx
				query=strSQL & " ORDER BY " & strORD1
			else
				strSQL=strSQL& ")" & tmpStrEx
				query=strSQL & " ORDER BY " & strORD1
			end if
		
		ELSE 'Exact=1
		
			if pKeywords<>"" then
			
				strSQl=strSql & " AND ("
				
				tmpSQL="(A.details LIKE "
				tmpSQL2="(A.description LIKE "
				tmpSQL3="(A.sDesc LIKE "
				if tIncludeSKU="true" then
					tmpSQL4="(A.SKU LIKE "
				end if
				Pos=0
				For L=LBound(keywordArray) to UBound(keywordArray)
					if trim(keywordArray(L))<>"" then
					Pos=Pos+1
					if Pos>1 Then
						tmpSQL=tmpSQL  & TestWord & " A.details LIKE "
						tmpSQL2=tmpSQL2 & TestWord & " A.description LIKE "
						tmpSQL3=tmpSQL3 & TestWord & " A.sDesc LIKE "
						if tIncludeSKU="true" then
							tmpSQL4=tmpSQL4 & TestWord & " A.SKU LIKE "
						end if
					end if
						tmpSQL=tmpSQL & trim(keywordArray(L))
						tmpSQL2=tmpSQL2 & trim(keywordArray(L))
						tmpSQL3=tmpSQL3 & trim(keywordArray(L))
						if tIncludeSKU="true" then
							tmpSQL4=tmpSQL4 & trim(keywordArray(L))
						end if
					end if
				Next
				tmpSQL=tmpSQL & ")"
				tmpSQL2=tmpSQL2 & ")"
				tmpSQL3=tmpSQL3 & ")"
				if tIncludeSKU="true" then
					tmpSQL4=tmpSQL4 & ")"
				end if
				
				strSQL=strSQL & tmpSQL
				strSQL=strSQL & " OR " & tmpSQL2
				if tIncludeSKU="true" then
					strSQL=strSQL & " OR " & tmpSQL3
					strSQL=strSQL & " OR " & tmpSQL4 & ")"
				else	
					strSQL=strSQL & " OR " & tmpSQL3 & ")"
				end if
				strSQL=strSQL& ")" & tmpStrEx
				query=strSQL & " ORDER BY " & strORD1
			else
				strSQL=strSQL& ")" & tmpStrEx
				query=strSQL & " ORDER BY " & strORD1
			end if
		END IF 'Exact

		queryCSF=query
		set rsCSF=Server.CreateObject("ADODB.Recordset")
		rsCSF.Open queryCSF, conlayout, adOpenStatic, adLockReadOnly, adCmdText
		if not rsCSF.eof then
			pcf_CountSearchResults = rsCSF.recordcount
		else
			pcf_CountSearchResults = 0
		end if
		set rsCSF=nothing
		
		tmp_StrQuery=""
		zSQL=""
		strSQL=""
		TestWord=""
		keywordArray=""
		pKeywords=""
		tmpStrEx=""
		tmpSQL=""
		tmpSQL2=""
		tmpSQL3=""
	End Function
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_CountResults & pcf_CountSearchResults
	'//////////////////////////////////////////////////////////////////////////////////////////



	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcf_InUse
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Determines a specific value is contained within an array
	'// Params:		A string of value separated by pipes
	'// Returns:	Boolean 
	Function pcf_InUse(theString,theValue)
		Dim r, tmpSValuesInUse
		pcf_InUse = 0		
		If instr(theString,"||")>0 Then			
			tmpSValuesInUse=split(theString,"||")
			For r=lbound(tmpSValuesInUse) to ubound(tmpSValuesInUse)
				if tmpSValuesInUse(r)<>"" then
					if cdbl(tmpSValuesInUse(r)) = cdbl(theValue) then
						pcf_InUse = 1
						Exit For
					end if
				end if
			Next
		Else
			if theString<>"" AND theValue<>"" then
				if cdbl(theString) = cdbl(theValue) then
					pcf_InUse = 1
				end if	
			end if
		End If
	End Function
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_InUse
	'//////////////////////////////////////////////////////////////////////////////////////////
	
	
	
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcf_SanitizeJava
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Escape Apostrophees in JavaScript
	'// Params:		A string
	'// Returns:	A safe string 
	Function pcf_SanitizeJava(theSting)
		pcf_SanitizeJava = replace(theSting,"'","\'")	
	End Function
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_SanitizeJava
	'//////////////////////////////////////////////////////////////////////////////////////////
	
	
	
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcf_CSFieldQuery
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Creates all the parameters needed for the query string
	'// Params:		Request variables for query string
	'// Returns:	A new query string 
	Function pcf_CSFieldQuery()
		pcf_CSFieldQuery = "&SFID="& pcv_intSFID &"&SFNAME="& Server.URLEncode(pcv_strSFNAME) &"&SFVID="& pcv_strCValues &"&SFVALUE="& Server.URLEncode(pcv_strSFVALUE) &"&SFCount="& pcv_intSFCount
	End Function
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_SanitizeJava
	'//////////////////////////////////////////////////////////////////////////////////////////
	
	
	
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// START: pcs_CSFSetVariables
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// Summary: 	Sets all the needed variables from the query string
	'// Params:		Request variables from query string
	'// Returns:	All the needed variables 
	Dim pcv_intSFID, pcv_strSFNAME, pcv_strCValues, pcv_strSFVALUE, pcv_intSFCount
	Sub pcs_CSFSetVariables()
		pcv_intSFID = getUserInput(Request("SFID"),0)
		pcv_strSFNAME = getUserInput(Request("SFNAME"),0)
		pcv_strSFNAME = replace(pcv_strSFNAME,"''","'")
		pcv_strCValues = getUserInput(Request("SFVID"),0)
		pcv_strSFVALUE = getUserInput(Request("SFVALUE"),0)
		pcv_strSFVALUE = replace(pcv_strSFVALUE,"''","'")	
		pcv_intSFCount = getUserInput(Request("SFCount"),10)
		if not validNum(pcv_intSFCount) then
			pcv_intSFCount=-1
		end if
	End Sub
	'//////////////////////////////////////////////////////////////////////////////////////////
	'// END: pcf_SanitizeJava
	'//////////////////////////////////////////////////////////////////////////////////////////
	
    If len(pcv_strSpryVarArray)>0 Then
		%>
		<script language="JavaScript" type="text/javascript">
		// Declare Variables
		var Spry;
		if (!Spry.Utils) Spry.Utils = {};
		
		// Declare vars outside the function so u can use then collap panel everywhere
		<% for xSpryCount = 1 to pcv_strSpryVars %>
			var CollapsiblePanel<%=xSpryCount%>;
		<% next %>

		// This function will fire anything when u leave the current page.		
		Spry.Utils.addUnLoadListener = function(handler)
		{
		if (typeof window.addEventListener != 'undefined')
			window.addEventListener('unload', handler, false);
		else if (typeof document.addEventListener != 'undefined')
			document.addEventListener('unload', handler, false);
		else if (typeof window.attachEvent != 'undefined')
			window.attachEvent('onunload', handler);
		};
		
		// This function will be used by cookie to check if the value is allready in the cookie, if so it returns it position
		Spry.Utils.CheckArray = function(a, s){
			for (i=0; i<a.length; i++){if (a[i] == s)return i;}return null;
		};		
		Spry.Utils.Cookie = function(type,string,options){
			if(type == 'create'){
				var expires='';
				if(options.days != null){
					var date = new Date();
					var UTCString;
					date.setTime(date.getTime()+(days*24*60*60*1000));
					expires = "; expires="+date.toUTCString();
				}
				var thePath = '; path=/';
				if(options.path != null){
					thePath = '; path='+options.path;
				}
				document.cookie = options.name+'='+escape(string)+expires+thePath;
			}else if(type == 'get'){
				var nameEQ = options.name + '=';
				var ca = document.cookie.split(';');
				for (var i=0; i<ca.length; i++){
					var c = ca[i];
					while (c.charAt(0)==' '){
					c = c.substring(1,c.length);
				}if (c.indexOf(nameEQ)==0) return unescape(c.substring(nameEQ.length,c.length)).split(",");
				}return null;
			}else if(type == 'destory'){
				Spry.Utils.Cookie('create','',{
					name: options.name
				});
			}else if(type == 'add'){
				var c = Spry.Utils.Cookie('get','',{name:options.name});
				if (typeof string == 'object') {
					for (i = 0, str; str = string[i], i < string.length; i++) {
						if (Spry.Utils.CheckArray(c, str) == null)c.push(str);
					}
				}else{
					if (Spry.Utils.CheckArray(c, string) == null) c.push(string)
				}
				Spry.Utils.Cookie('create',c,{name:options.name});		
			}
		};
		
		// Check if we have a panel to set from cookie value
		// We are using spry dom we can also include the tab construction within this function.
		Spry.Utils.addLoadListener(function(){
			// Cookie GET returns a array we have to add a var before it to store the data
			var col = Spry.Utils.Cookie('get','',{name:'col_history<%=pcv_strCSFCatID%>'});
			if(col)//checks if acc contains data
				  {
					<% for xSpryCount = 1 to pcv_strSpryVars %>
					
						if(col[<%=xSpryCount%>] == 'false')CollapsiblePanel<%=xSpryCount%> = new Spry.Widget.CollapsiblePanel("CollapsiblePanel<%=xSpryCount%>",{contentIsOpen:false});
						else CollapsiblePanel<%=xSpryCount%> = new Spry.Widget.CollapsiblePanel("CollapsiblePanel<%=xSpryCount%>");
					
					<% next %>		
					
				  }else{
			
			<% for xSpryCount = 1 to pcv_strSpryVars %>
			
				CollapsiblePanel<%=xSpryCount%> = new Spry.Widget.CollapsiblePanel("CollapsiblePanel<%=xSpryCount%>");
				//CollapsiblePanel<%=xSpryCount%> = new Spry.Widget.CollapsiblePanel("CollapsiblePanel<%=xSpryCount%>",{contentIsOpen:false});
			
			<% next %>
			}
		 });
		
		Spry.Utils.addUnLoadListener(function(){
			// Create a array where all panel statuses are stored
			<%
			'// Clean last comma from end of string
			if len(pcv_strSpryVarArray)>0 then
				pcv_strSpryVarArray = Left(pcv_strSpryVarArray,Len(pcv_strSpryVarArray)-1) 
			end if 
			%>
			var panel = new Array(<%=pcv_strSpryVarArray%>);
			// Destroy old cookie (I want a clean cookie)
			Spry.Utils.Cookie('destory','',{name:'col_history<%=pcv_strCSFCatID%>'});
			// Setting the new cookie value
			Spry.Utils.Cookie('create',panel,{name:'col_history<%=pcv_strCSFCatID%>'});
		});
		</script>
		<%
	End If '// If len(pcv_strSpryVarArray)>0 Then
	
END IF
'//////////////////////////////////////////////////////////////////////////////////////////////
'// END: CATEGORY SEARCH FIELDS
'//////////////////////////////////////////////////////////////////////////////////////////////
%>