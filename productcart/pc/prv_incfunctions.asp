<% sub WriteStar(svalue,wmode)
	Rev_Rating=svalue
	if (svalue="0") or (svalue=null) or (svalue="") then
		' response.write dictLanguage.Item(Session("language")&"_prv_15") & "<br>"
	else
		Rev_SaveRate=Round(Rev_Rating,1)
		if pcv_MaxRating="5" AND cdbl(Rev_Rating)>5 then
			Rev_Rating=Rev_Rating/2
		end if

		tmp3=0
		if Round(Rev_Rating)-Rev_Rating>0 then
			tmp3=Round(Rev_Rating)-1
		else
			if Round(Rev_rating)-Rev_Rating<0 then
				tmp3=Round(Rev_Rating)
			else
				tmp3=Round(rev_Rating)
			end if
		end if%>
		<%For l=1 to tmp3%>
			<img src="catalog/<%=pcv_Img3%>" align="baseline">
		<%Next
		if (Round(Rev_rating)-Rev_Rating<0) or (Round(Rev_rating)-Rev_Rating=0.5) then%>
			<img src="catalog/<%=pcv_Img4%>" align="baseline">
			<%tmp3=tmp3+1
		else
			if Round(Rev_rating)-Rev_Rating>0 then%>
				<img src="catalog/<%=pcv_Img3%>" align="baseline">
				<%tmp3=tmp3+1
			end if
		end if
		For l=tmp3+1 to pcv_MaxRating	%>
			<img src="catalog/<%=pcv_Img5%>" align="baseline">
		<%Next%>
		<%
		if wmode="1" then
			' Uncomment the following line to show the text value (e.g. "3.5")
			' response.write Rev_SaveRate
		end if%>
        <br>
	<% end if
end sub

Dim Fi(100)
Dim FS(100)
Dim FRe(100)
Dim FName(100)
Dim FType(100)
Dim FValue(100)
Dim FRecord(100)
Dim FCount

sub CreateList()
	if not validNum(pcv_IDProduct) then pcv_IDProduct=pIDProduct
	query="SELECT pcRS_FieldList,pcRS_FieldOrder,pcRS_Required FROM pcReviewSpecials WHERE pcRS_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if not rs.eof then
		pcv_FieldList=split(rs("pcRS_FieldList"),",")
		pcv_FieldOrder=split(rs("pcRS_FieldOrder"),",")
		pcv_Required=split(rs("pcRS_Required"),",")
		set rs=nothing
		
		FCount=0
		For iRevPrdCount=0 to ubound(pcv_FieldList)
			if pcv_FieldList(iRevPrdCount)<>"" then
				Fi(FCount)=pcv_FieldList(iRevPrdCount)
				FS(FCount)=pcv_FieldOrder(iRevPrdCount)
				FRe(FCount)=pcv_Required(iRevPrdCount)
					
				query="SELECT pcRF_Type,pcRF_Name FROM pcRevFields WHERE pcRF_IDField=" & Fi(FCount)
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)

				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
				if not rs.eof then
					FType(FCount)=rs("pcRF_Type")
					FName(FCount)=rs("pcRF_Name")
				end if
				set rs=nothing
				
				FCount=FCount+1

			end if
		Next
		
		For iRevPrdCount=0 to FCount-1
			For j=iRevPrdCount+1 to FCount-1
				if FS(iRevPrdCount)>FS(j) then
					tmpC=FS(j)
					FS(j)=FS(iRevPrdCount)
					FS(iRevPrdCount)=tmpC
					
					tmpC=Fi(j)
					Fi(j)=Fi(iRevPrdCount)
					Fi(iRevPrdCount)=tmpC
					
					tmpC=FRe(j)
					FRe(j)=FRe(iRevPrdCount)
					FRe(iRevPrdCount)=tmpC
					
					tmpC=FType(j)
					FType(j)=FType(iRevPrdCount)
					FType(iRevPrdCount)=tmpC
					
					tmpC=FName(j)
					FName(j)=FName(iRevPrdCount)
					FName(iRevPrdCount)=tmpC
				end if
			Next
		Next
	else
		query="SELECT pcRF_IDField,pcRF_Name,pcRF_Type,pcRF_Required,pcRF_Order FROM pcRevFields WHERE pcRF_Active=1 order by pcRF_Order asc"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
	
		if not rs.eof then
			pcArray=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArray,2)
		
			FCount=0
		
			For iRevPrdCount=0 to intCount
				Fi(FCount)=pcArray(0,iRevPrdCount)
				FName(FCount)=pcArray(1,iRevPrdCount)
				FType(FCount)=pcArray(2,iRevPrdCount)
				FRe(FCount)=pcArray(3,iRevPrdCount)
				FS(FCount)=pcArray(4,iRevPrdCount)
				FCount=FCount+1
			Next
			
		end if
	
	end if
	
end sub

function CalRating()
	if not validNum(pcv_IDProduct) then pcv_IDProduct=pIDProduct
	Dim subtotal1,subtotal2
	subtotal1=0
	subtotal2=0
	if FCount>0 then
		query="SELECT pcRev_IDReview FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		
		pcv_tmpRating=0
	
		if not rs.eof then
			pcArray=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArray,2)
			For l=0 to intCount
				query="SELECT pcRD_IDField,pcRD_Rate,pcRD_Feel FROM pcReviewsData WHERE pcRD_IDReview=" & pcArray(0,l) & " and ((pcRD_Rate>0) or (pcRD_Feel>0))"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rs=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				if not rs.eof then
					pcArray1=rs.getRows()
					set rs=nothing
					intCount1=ubound(pcArray1,2)
					For n=0 to intCount1
						For p=0 to FCount-1
							if clng(Fi(p))=clng(pcArray1(0,n)) then
								IF FType(p)="4" THEN
									subtotal1=subtotal1+pcArray1(1,n)
									subtotal2=subtotal2+1
									if FValue(p)>"0" then
									else
										FValue(p)=0
									end if
									FValue(p)=FValue(p)+pcArray1(1,n)
									if FRecord(p)>"0" then
									else
										FRecord(p)=0
									end if
									FRecord(p)=FRecord(p)+1
								ELSE
									if FValue(p)>"0" then
									else
										FValue(p)=0
									end if
									if pcArray1(2,n)="2" then
									FValue(p)=FValue(p)+1
									end if
									if FRecord(p)>"0" then
									else
										FRecord(p)=0
									end if
									FRecord(p)=FRecord(p)+1
								END IF
							end if
						Next
					Next
				end if
			Next
			if (subtotal1>"0") and (subtotal2>"0") then
				subtotal1=subtotal1/subtotal2
			else
				subtotal1="0"
			end if
		end if
	end if
	CalRating=subtotal1
end function
%>