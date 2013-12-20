<% 'PRV41 start %>
<%IF pRSActive AND pShowAvgRating THEN

    ' Assign pIDProduct to pcv_IDProduct - pcv_IDProduct is used in prv_incfunctions.asp
    pcv_IDProduct = pIDProduct

	IF pcv_RatingType="0" then

        query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if

		pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

		query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pIDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0"
		set rs=connTemp.execute(query)

		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		intCount=0
		if not rs.eof then
			intCount = clng(rs("ct"))
		end if

		set rs=Nothing
		
		%>
		<% If pcv_tmpRating>"0" Then %>
		<p style="white-space: nowrap;"><img src="catalog/<%=pcv_Img1%>" align="absbottom"><%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%> (<%=intCount%>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)</p>
		<% End If %>
	<%
	ELSE
		if pcv_CalMain="1" Then
		
			query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

            query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pIDProduct & " AND pcRev_Active=1 AND pcRev_MainDRate>0"
            set rs=connTemp.execute(query)

            if err.number<>0 then
                call LogErrorToDatabase()
                set rs=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
			if not rs.eof then
	            intCount = clng(rs("ct"))
			end if
			set rs=nothing
			
			if pcv_tmpRating>"0" then%>
			<p style="white-space: nowrap;"><%=dictLanguage.Item(Session("language")&"_prv_39")%>
			<% Call WriteStar(pcv_tmpRating,1) %></p>
			<%end if%>
		<% else
		    Call CreateList()
			pcv_SaveRating=CalRating()
			pcv_tmpRating=pcv_SaveRating
			if pcv_tmpRating>"0" then%>
			<p style="white-space: nowrap;"><% Call WriteStar(pcv_tmpRating,1) %></p>
			<%end if%>
		<% end if
	END IF 'Main Rating

END IF%>
<% 'PRV41 end %>
