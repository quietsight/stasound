<% if scATCEnabled="1" then
    Dim originalurl, rooturl, homepageurl
    
    pcv_URLPrefix = scStoreURL&"/"&scPcFolder
    pcv_URLPrefix = replace(pcv_URLPrefix,"//","/")
    pcv_URLPrefix = replace(pcv_URLPrefix,"http:/","http://")
    pcv_URLPrefix = replace(pcv_URLPrefix,"https:/","https://")
    
    rooturl = pcv_URLPrefix&"/pc/"
    if scURLredirect = "" then
        homepageurl = pcv_URLPrefix&"/pc/home.asp"
    else
        homepageurl = scURLredirect
    end if
    
    atc_Debug = 0	' 1 = ON
    
    ' --------------------------------------------------------------------
     
    originalurl 	= lcase(Request.ServerVariables("HTTP_REFERER"))
	If len(originalurl)=0 Then
		if scSeoURLs = 0 then
			originalurl	= rooturl & "viewPrd.asp"
		else
			originalurl	= rooturl & "viewcart.asp"
		end if
	End If
    atc_idProduct 	= getuserinput(request("idproduct"),0)
    
	' --------------------------------------------------------------------
	' // Debugging
	' response.write "homepageurl=" & homepageurl & "<br>"
	' response.write "rooturl=" & rooturl & "<br>"
	' response.write "originalurl=" & originalurl & "<br>"
	' response.End()
    ' --------------------------------------------------------------------
    
    if originalurl = rooturl then
        originalurl = homepageurl & "?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
            if atc_debug = 1 then originalurl = originalurl & "&home=1"			
        end if
        response.redirect originalurl
        response.end
    
    elseif originalurl = homepageurl then
        originalurl = homepageurl & "?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
			if atc_debug = 1 then originalurl = originalurl & "&home=2"			
        end if
        response.redirect originalurl
        response.end
    
    elseif InStr(originalurl,"showbestsellers.asp") <> 0  then
        originalurl = rooturl2 & "showbestsellers.asp?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
            if atc_debug = 1 then originalurl = originalurl & "&bestsellers"			
        end if
        response.redirect originalurl
        response.end
    
    elseif InStr(originalurl,"showfeatured.asp") <> 0  then
        originalurl = rooturl2 & "showfeatured.asp?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
            if atc_debug = 1 then originalurl = originalurl & "&featured"			
        end if
        response.redirect originalurl
        response.end
    
    elseif InStr(originalurl,"shownewarrivals.asp") <> 0  then
        originalurl = rooturl2 & "shownewarrivals.asp?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
            if atc_debug = 1 then originalurl = originalurl & "&new"			
        end if
        response.redirect originalurl
        response.end
    
    elseif InStr(originalurl,"showspecials.asp") <> 0  then
        originalurl = rooturl2 & "showspecials.asp?idproduct=" & atc_idproduct 
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
            if atc_debug = 1 then originalurl = originalurl & "&specials"			
        end if
        response.redirect originalurl
        response.end
    
    elseif InStr(originalurl,"showsearchresults.asp") <> 0  then
    
        if InStr(originalurl,"atc=")= 0 then
            ' on the first pass, all we have to do is add the flag and the product id
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
			else
            	originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
			end if
            if atc_debug = 1 then originalurl = originalurl & "&searchresults=1"			
            response.redirect originalurl
            response.end
        else
            ' on subsequent passes, first we have to strip the flag and product id, then add them back
        
            startpos 	  		= InStr(originalurl,"atc")			' where atc begins
            searchurl			= left(originalurl,startpos-1)
            originalurl = searchurl 
            originalurl = originalurl & "&atc=1"
            originalurl = originalurl & "&idproduct=" & atc_idproduct
            if atc_debug = 1 then originalurl = originalurl & "&searchresults=2"			 
            response.redirect originalurl
            response.end
        end if
    
    elseif InStr(originalurl,"viewcategories.asp") <> 0  then
    
		if InStr(originalurl,"idcategory")=0 then
			originalurl=originalurl & "?idcategory=1"
		end if
		
        lenourl	  			= len(originalurl)							' length of referrer URL string
        startpos 	  		= InStr(originalurl,"idcategory")			' where idcategory begins
        midstring 			= mid(originalurl, startpos, lenourl-1)
        eqpos 				= instr(midstring, "=")						' where the "=" following idcategory is located
        lenvalue 			= lenourl - eqpos
        if lenvalue > 0 then 
            beyondeq		= mid(midstring, eqpos+1, lenvalue)
        end if
        ampersandpos	  		= instr(1,beyondeq,"&") ' is there another variable beyond the category?
        if ampersandpos > 0 then
            category		  	= left(beyondeq,ampersandpos-1) 
        else
            category		  	= beyondeq
        end if 
  
        if lcase(InStr(originalurl,"sfid")) = 0  then ' normal category page call        
        	originalurl = rooturl 
			if category>1 then
			originalurl = originalurl & "viewCategories.asp?idcategory=" & category 
			originalurl = originalurl & "&idproduct=" & atc_idproduct 
			originalurl = originalurl & "&atc=1"
			else
				originalurl = originalurl & "viewCategories.asp?atc=1"
			end if
        else ' custom search fields
	        if InStr(originalurl,"atc=1") = 0  then
		        originalurl = originalurl & "&atc=1"
	        end if 
        end if
    
    ElseIf InStr(originalurl,"viewprd.asp") <> 0  Then
    
        If InStr(originalurl,"atc=")= 0 Then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
        	if atc_debug = 1 then originalurl = originalurl & "&viewprd=1"			 
        End If
        Response.redirect originalurl
        Response.End
    
    elseif request("pCnt") > 0  then
        ' this code handles format "M" - multiple add-to-cart
        '	response.redirect "viewCart.asp"
        '	response.end
    
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if
        end if
        response.redirect originalurl
        response.end
    end if
    
    
    
    if InStr(originalurl,".htm")<> 0 then
    
        If InStr(originalurl,"atc=")= 0 Then
            ' on the first pass, all we have to do is add the flag and the product id
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
			else
            	originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
			end if
            if atc_debug = 1 then originalurl = originalurl & "&alternate=1"			 
            Response.redirect originalurl
            Response.End
    
        Else
            ' on subsequent passes, first we have to strip the flag and product id, then add them back
            startpos    = InStr(originalurl,"?atc")			' where atc begins
            searchurl   = Left(originalurl,startpos - 1)
            originalurl = searchurl
            originalurl = originalurl & "?atc=1"
            originalurl = originalurl & "&idproduct=" & atc_idproduct
            if atc_debug = 1 then originalurl = originalurl & "&alternate=2"			 
    
            If InStr(originalurl,"atc=")= 0 Then
                ' on the first pass, all we have to do is add the flag and the product id
				if InStr(originalurl,"?") then
					originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
				else
					originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
				end if
                Response.redirect originalurl
                Response.End
            Else
                ' on subsequent passes, first we have to strip the flag and product id, then add them back
                startpos    = InStr(originalurl,"?atc")			' where atc begins
                searchurl   = Left(originalurl,startpos - 1)
                originalurl = searchurl
                originalurl = originalurl & "?atc=1"
                originalurl = originalurl & "&idproduct=" & atc_idproduct
                if atc_debug = 1 then originalurl = originalurl & "&alternate=3"		
                Response.redirect originalurl
                Response.End
            End If
    
        End If
    
        If InStr(originalurl,"idproduct") = 0 Then
            originalurl = originalurl & "&idproduct=" & atc_idproduct
        End If
    
        Response.redirect originalurl
        Response.End
    End If

    If InStr(originalurl,"idproduct") = 0 Then
        if InStr(originalurl,"atc=")= 0 then
			if InStr(originalurl,"?") then
            	originalurl = originalurl & "&atc=1"
			else
            	originalurl = originalurl & "?atc=1"			
			end if			
        end if
		originalurl = originalurl & "&idproduct=" & atc_idproduct
    End If
    
    Response.redirect originalurl
End if %>