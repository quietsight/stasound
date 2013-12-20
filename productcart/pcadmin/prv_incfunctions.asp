<% 
sub WriteStar(svalue,wmode)
Rev_Rating=svalue
if (svalue="0") or (svalue=null) or (svalue="") then
	'response.write dictLanguage.Item(Session("language")&"_prv_15")
else
	Rev_SaveRate=Round(Rev_Rating,1)
	if pcv_MaxRating="10" then
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
	end if
	
	For l=1 to tmp3 %>
		<img src="../pc/catalog/<%=pcv_Img3%>" border="0" align="baseline">
	<%
	Next
		if (Round(Rev_rating)-Rev_Rating<0) or (Round(Rev_rating)-Rev_Rating=0.5) then
		%>
			<img src="../pc/catalog/<%=pcv_Img4%>" border="0" align="baseline">
		<%
		tmp3=tmp3+1
		else
			if Round(Rev_rating)-Rev_Rating>0 then
		%>
				<img src="../pc/catalog/<%=pcv_Img3%>" border="0" align="baseline">
				<%
				tmp3=tmp3+1
			end if
		end if
		For l=tmp3+1 to 5
		%>
			<img src="../pc/catalog/<%=pcv_Img5%>" border="0" align="baseline">
		<%
		Next
		
		if wmode="1" then
			response.write Rev_SaveRate
		end if
		%>
        <br>
	<%
	end if
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
	
	
	query="SELECT pcRS_FieldList,pcRS_FieldOrder,pcRS_Required FROM pcReviewSpecials WHERE pcRS_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		pcv_FieldList=split(rs("pcRS_FieldList"),",")
		pcv_FieldOrder=split(rs("pcRS_FieldOrder"),",")
		pcv_Required=split(rs("pcRS_Required"),",")
		
		set rs=nothing
		
		FCount=0
		For i=0 to ubound(pcv_FieldList)
			if pcv_FieldList(i)<>"" then
				Fi(FCount)=pcv_FieldList(i)
				FS(FCount)=pcv_FieldOrder(i)
				FRe(FCount)=pcv_Required(i)
					
				query="SELECT pcRF_Type,pcRF_Name FROM pcRevFields WHERE pcRF_IDField=" & Fi(FCount)
				set rstemp=server.CreateObject("ADODB.RecordSet")
				set rstemp=connTemp.execute(query)
				
				if not rstemp.eof then
				
				FType(FCount)=rstemp("pcRF_Type")
				FName(FCount)=rstemp("pcRF_Name")
				
				end if
				set rstemp=nothing
				FCount=FCount+1
			end if
		Next

		For i=0 to FCount-1
			For j=i+1 to FCount-1
				if FS(i)>FS(j) then
					tmpC=FS(j)
					FS(j)=FS(i)
					FS(i)=tmpC
					
					tmpC=Fi(j)
					Fi(j)=Fi(i)
					Fi(i)=tmpC
					
					tmpC=FRe(j)
					FRe(j)=FRe(i)
					FRe(i)=tmpC
					
					tmpC=FType(j)
					FType(j)=FType(i)
					FType(i)=tmpC
					
					tmpC=FName(j)
					FName(j)=FName(i)
					FName(i)=tmpC
				end if
			Next
		Next

	else

		query="SELECT pcRF_IDField,pcRF_Name,pcRF_Type,pcRF_Required,pcRF_Order FROM pcRevFields WHERE pcRF_Active=1 order by pcRF_Order asc"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)

		if not rs.eof then
			pcArray=rs.getRows()
			set rs=nothing
			intCount=ubound(pcArray,2)
		
			FCount=0
			
			For i=0 to intCount
				Fi(FCount)=pcArray(0,i)
				FName(FCount)=pcArray(1,i)
				FType(FCount)=pcArray(2,i)
				FRe(FCount)=pcArray(3,i)
				FS(FCount)=pcArray(4,i)
				FCount=FCount+1
			Next
		
		end if

	end if
	
	
end sub

function CalRating()

	Dim subtotal1,subtotal2
	subtotal1=0
	subtotal2=0
	
	if FCount>0 then
		query="SELECT pcRev_IDReview FROM pcReviews WHERE pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
	
		pcv_tmpRating=0
		
		if not rs.eof then
			pcArray=rs.getRows()
			set rs=nothing
			
			intCount=ubound(pcArray,2)
			For l=0 to intCount
				query="SELECT pcRD_IDField,pcRD_Rate,pcRD_Feel FROM pcReviewsData WHERE pcRD_IDReview=" & pcArray(0,l) & " and ((pcRD_Rate>0) or (pcRD_Feel>0))"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
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