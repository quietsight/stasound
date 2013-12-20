<% if scATCEnabled="1" then %>

    <%
    'Show shopping cart total
    'dim vcCartArr, vcCartIndex, v, scantidadCart7, vcTotal, vcItems
    
    vcCartArr=Session("pcCartSession")
    vcCartIndex=Session("pcCartIndex")
    vcTotal=Cint(0) 'calculates the cart total
    vcItems=Cint(0)	'counts items in your cart
    vcPrice=Cint(0) 'calculates the cart cross sell product
    vcpPrice=Cint(0) 'calculates the cart cross sell parent product
	
	Dim sscProList2(100,5)
    
    for v=1 to vcCartIndex
		sscProList2(v,0)=vcCartArr(v,0)
		sscProList2(v,1)=vcCartArr(v,10)
		sscProList2(v,3)=vcCartArr(v,2)
		sscProList2(v,4)=0
        if InStr(Cstr(10/3),",")>0 then
            if Instr(vcCartArr(v,17),".")>0 then
                if IsNumeric(vcCartArr(v,17)) then
                    vcCartArr(v,17)=replace(vcCartArr(v,17),".",",")
                end if
            end if
        else
            if Instr(vcCartArr(v,17),",")>0 then
                if IsNumeric(vcCartArr(v,17)) then
                    vcCartArr(v,17)=replace(vcCartArr(v,17),",",".")
                end if
            end if
        end if
    
        if vcCartArr(v,10)=0 then
            vcItems=vcItems + vcCartArr(v,2)
            vcPrice=Cint(0)
            vcPrice=vcPrice+(vcCartArr(v,2)*vcCartArr(v,17))
            vcPrice=vcPrice+(vcCartArr(v,2)*vcCartArr(v,5))
			if vcCartArr(v,30)<>"" then
				vcPrice=vcPrice-vcCartArr(v,30)
			end if
			if vcCartArr(v,31)<>"" then
				vcPrice=vcPrice+vcCartArr(v,31)
			end if
            vcPrice=vcPrice-vcCartArr(v,15)
            
            if trim(vcCartArr(v,27))="" then
                vcCartArr(v,27)=0
            end if
            if trim(vcCartArr(v,28))="" then
                vcCartArr(v,28)=0
            end if
            
            if (vcCartArr(v,27)>0) AND (vcCartArr(v,28)>0) then 
                vp = cint(vcCartArr(v,27))
                vcpPrice=Cint(0)
                vcpPrice=vcpPrice+(vcCartArr(vp,2)*vcCartArr(vp,17))
                vcpPrice=vcpPrice+(vcCartArr(vp,2)*vcCartArr(vp,5))
                vcpPrice=vcpPrice-vcCartArr(vp,30)
                vcpPrice=vcpPrice+vcCartArr(vp,31)
                vcpPrice=vcpPrice-vcCartArr(vp,15)
                vcPrice = (cdbl(vcPrice)+cdbl(vcpPrice)) - ((cdbl(vcCartArr(v,28)) + cdbl(vcCartArr(vp,28)))*vcCartArr(v,2))
            end if	
	
			sscProList2(v,2)=vcPrice
		
			'// Don't Add to total if parent of a Bundle Cross Sell Product
			pcv_HaveBundles=0
			if vcCartArr(v,27)=-1 then
				for mc=1 to vcCartIndex
					if (vcCartArr(mc,27)<>"") AND (vcCartArr(mc,12)<>"") then
						if cint(vcCartArr(mc,27))=v AND cint(vcCartArr(mc,12))="0" then
							pcv_HaveBundles=1
							exit for
						end if
					end if
				next
			end if
			if (vcCartArr(v,27)>-1) OR (pcv_HaveBundles=0) then
				vcTotal=vcTotal+vcPrice
			end if		
		end if
	
	next


	' ------------------------------------------------------
	' START - Calculate category-based quantity discounts
	' ------------------------------------------------------
	Set conTempSC=Server.CreateObject("ADODB.Connection")
	conTempSC.Open scDSN
	
	CatDiscTotal=0
	
	query="SELECT pcCD_idCategory as IDCat FROM pcCatDiscounts group by pcCD_idCategory"
	set rsSSCCatDis=server.CreateObject("ADODB.RecordSet")
	set rsSSCCatDis=conTempSC.execute(query)
	Do While not rsSSCCatDis.eof
		CatSubQty=0
		CatSubTotal=0
		CatSubDiscount=0
	
		for v=1 to vcCartIndex
			if (sscProList2(v,1)=0) and (sscProList2(v,4)=0) then
				if (vcCartArr(v,32)<>"") then
					pcv_tmpPPrd=split(vcCartArr(v,32),"$$")
					pcv_tmpID=pcv_tmpPPrd(ubound(pcv_tmpPPrd))
				else
					pcv_tmpID=sscProList2(v,0)
				end if
				query="select idproduct from categories_products where idcategory=" & rsSSCCatDis("IDCat") & " and idproduct=" & pcv_tmpID
				set rsSSCProd=server.CreateObject("ADODB.RecordSet")
				set rsSSCProd=conTempSC.execute(query)
				if not rsSSCProd.eof then
					CatSubQty=CatSubQty+sscProList2(v,3)
					CatSubTotal=CatSubTotal+sscProList2(v,2)
					sscProList2(v,4)=1
				end if
				set rsSSCProd=nothing
			end if
		Next
	
		if CatSubQty>0 then
			query="SELECT pcCD_discountPerUnit,pcCD_discountPerWUnit,pcCD_percentage,pcCD_baseproductonly FROM pcCatDiscounts WHERE pcCD_idCategory=" & rsSSCCatDis("IDCat") & " AND pcCD_quantityFrom<=" &CatSubQty& " AND pcCD_quantityUntil>=" &CatSubQty
			set rsSSCDiscount=server.CreateObject("ADODB.RecordSet")
			set rsSSCDiscount=conTempSC.execute(query)
			if not rsSSCDiscount.eof then
	
				' there are quantity discounts defined for that quantity 
				pDiscountPerUnit=rsSSCDiscount("pcCD_discountPerUnit")
				pDiscountPerWUnit=rsSSCDiscount("pcCD_discountPerWUnit")
				pPercentage=rsSSCDiscount("pcCD_percentage")
				pbaseproductonly=rsSSCDiscount("pcCD_baseproductonly")
				set rsSSCDiscount=nothing
				
				if session("customerType")<>1 then  'customer is a normal user
					if pPercentage="0" then 
						CatSubDiscount=pDiscountPerUnit*CatSubQty
					else
						CatSubDiscount=(pDiscountPerUnit/100) * CatSubTotal
					end if
				else  'customer is a wholesale customer
					if pPercentage="0" then 
						CatSubDiscount=pDiscountPerWUnit*CatSubQty
					else
						CatSubDiscount=(pDiscountPerWUnit/100) * CatSubTotal
					end if
				end if
			end if
		end if
	
		CatDiscTotal=CatDiscTotal+CatSubDiscount
		rsSSCCatDis.MoveNext
	loop
	set rsSSCCatDis=nothing	
	
	'// Round the Category Discount to two decimals
	if CatDiscTotal<>"" and isNumeric(CatDiscTotal) then
		CatDiscTotal = Round(CatDiscTotal,2)
		vcTotal=vcTotal-CatDiscTotal
	end if
	
	'// Display Applied Product Promotions (if any)
	if Session("pcPromoIndex")<>"" and Session("pcPromoIndex")>"0" then
		TotalPromotions=pcf_GetPromoTotal(Session("pcPromoSession"),Session("pcPromoIndex"))
		vcTotal=vcTotal-TotalPromotions
	end if	
	conTempSC.Close
	Set conTempSC=nothing
	' ------------------------------------------------------
	' END - Calculate category-based quantity discounts
	' ------------------------------------------------------


	if vcItems > 0 then	
		vcHaveGcsTest=0
		dim pcv_counterbc
		for pcv_counterbc=1 to vcCartIndex
			if vcCartArr(pcv_counterbc,10)=0 then
				Set conTempSC=Server.CreateObject("ADODB.Connection")
				conTempSC.Open scDSN
				query="select pcprod_Gc from Products where idproduct=" & vcCartArr(pcv_counterbc,0) & " AND pcprod_Gc=1"
				set rsGcVcObj=conTempSC.execute(query)
				if not rsGcVcObj.eof then
					vcHaveGcsTest=1
					exit for
				end if
				conTempSC.Close
				Set conTempSC=nothing
			end if
		next	
	end if
    
    '// Get sDesc
    dim atc_Description,atc_Sku,atc_idproduct,rsATCObj
    atc_idProduct = getuserinput(request("idproduct"),10)
    'Validate the input is a numeric value only
    if NOT ValidNum(atc_idProduct) then
        'Invalid entry - force product to 0
        atc_idProduct=0
    end if
	if atc_idProduct<>0 then
		query="SELECT description, sku FROM products WHERE idProduct=" & atc_idproduct & ";"
		Set conTempSC=Server.CreateObject("ADODB.Connection")
		conTempSC.Open scDSN
		
		set rsATCObj=server.CreateObject("ADODB.RecordSet")
		set rsATCObj=conTempSC.execute(query)
		atc_Description=replace(rsATCObj("description"),"&quot;",chr(34))
		atc_Sku=rsATCObj("Sku")
		set rsATCObj=nothing
		conTempSC.Close
		Set conTempSC=nothing
	end if
	%>

	<div id="OverlayMsgDialog" title="<%response.write dictLanguage.Item(Session("language")&"_addedtocart_1")%>" style="display:none; text-align: left;">

        <%
        Dim strCCSLCheck
        call opendb()
        strCCSLcheck = checkCartStockLevels(vcCartArr, vcCartIndex, aryBadItems)
        call closedb()
        If Len(Trim(strCCSLCheck))>0 Then
           response.write "<div class=""pcErrorMessage"">"
           response.write dictLanguage.Item(Session("language")&"_alert_19") & strCCSLcheck
           response.write "</div>"
        End If  
        %>	

    	<div class="ui-main">    
            <div style="margin: 5px 0 5px 0;">
            <% if atc_FlagM = 1 then %><%response.write dictLanguage.Item(Session("language")&"_addedtocart_2")%><% else %><% if len(atc_idProduct)> 0 and atc_Description<>"" then %><%response.write dictLanguage.Item(Session("language")&"_addedtocart_3")%>: <%Response.Write(atc_Description)%> <span class="pcSmallText">(<%Response.Write(atc_Sku)%>)</span><% else %><%response.write dictLanguage.Item(Session("language")&"_addedtocart_2")%><% end if %><% end if %>
            </div>
            <%response.write dictLanguage.Item(Session("language")&"_addedtocart_5")%><%response.write vcItems%><%response.write dictLanguage.Item(Session("language")&"_addedtocart_4")%><%response.write scCurSign & money(vcTotal)%>
    	</div>    
	</div>
	<script>
        /* check if ordered quantity is more than available - begin*/
        <%If Len(Trim(strCCSLCheck))< 1 Then%>
		    $(document).ready(function()
		    {
			    $("#OverlayMsgDialog").dialog({
					    bgiframe: true,
					    autoOpen: false,
					    resizable: false,
					    width: 400,
					    height: 190,
					    modal: true,
					    buttons: {

						    ' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_65"))%> ': function() {
								    location='viewcart.asp?cs=1';
								    $(this).dialog('close');						
						    },
						    ' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_78"))%> ': function() {
								    $(this).dialog('close');						
						    }
					    }
			    });
        <%else%>			
		    $(document).ready(function()
		    {
			    $("#OverlayMsgDialog").dialog({
					    bgiframe: true,
					    autoOpen: false,
					    resizable: false,
					    width: 680,
					    height: 400,
					    modal: true,
					    buttons: {

						    ' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_65"))%> ': function() {
								    location='viewcart.asp?cs=1';
								    $(this).dialog('close');						
						    },
						    ' <%=FixLang(dictLanguage.Item(Session("language")&"_opc_js_78"))%> ': function() {
								    $(this).dialog('close');						
						    }
					    }
			    });	

        <%end if%>			
        /* check if ordered quantity is more than available - end*/
    
			function getURLParam(strParamName){
				var strReturn = "";
				var strHref = window.location.href;
				if ( strHref.indexOf("?") > -1 ){
					var strQueryString = strHref.substr(strHref.indexOf("?")).toLowerCase();
					var aQueryString = strQueryString.split("&");
					for ( var iParam = 0; iParam < aQueryString.length; iParam++ ){
						if ( 
							aQueryString[iParam].indexOf(strParamName.toLowerCase() + "=") > -1 ){
							var aParam = aQueryString[iParam].split("=");
							strReturn = aParam[1];
							break;
						}
					}
				}
				return unescape(strReturn);
			} 
			
			function TurnOnUpdateDiv() {  
				var obj = document.getElementById("overlay");  
				if(getURLParam("atc") == 1) {  
					$("#OverlayMsgDialog").dialog('open');
				}  
				else  
				{  
					//
				}  
			}
		
			function addOnLoadEvent(func)
			{
				var oldonload = window.onload ;
				if (typeof window.onload  != 'function') {
				    window.onload  = func;
				}
				else
				{
					window.onload  = function() {
						oldonload();
						func();
					}
				}
			}

			addOnLoadEvent(TurnOnUpdateDiv);
			
		});
	</script>  
    
    
    
<% end if %>