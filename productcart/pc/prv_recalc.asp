<%
'PRV41 start
Function GetOverallProductRating(whichIDProduct)
   ' Dependencies: prv_getSettings.asp
   ' Purpose:      Passed a Product ID, will return the average rating (can be
   '               used on the fly or for updating the Products Table field
   '               called pcProd_AvgRating

	if pcv_RatingType="0" Then
       if scDB="Access" Then
	      query="SELECT (AVG(pcRev_MainRate-1)) * 100 AS AverageRating FROM pcReviews WHERE pcRev_IDProduct=" & whichIDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0;"
       else
	      query="SELECT (AVG(CONVERT(decimal(5,2), pcRev_MainRate-1))) * 100 AS AverageRating FROM pcReviews WHERE pcRev_IDProduct=" & whichIDProduct & " AND pcRev_Active=1 AND pcRev_MainRate>0;"
       End if
	Else
       if scDB="Access" Then
		  query="SELECT AVG(pcRev_MainDRate) AS AverageRating FROM pcReviews WHERE pcRev_IDProduct=" & whichIDProduct & " AND pcRev_Active=1 AND pcRev_MainDRate>0;"
       else
		  query="SELECT AVG(CONVERT(decimal(5,2),pcRev_MainDRate)) AS AverageRating FROM pcReviews WHERE pcRev_IDProduct=" & whichIDProduct & " AND pcRev_Active=1 AND pcRev_MainDRate>0;"
       End if
	End If

	set rs=connTemp.execute(query)
	If rs.eof Then
	   GetOverallProductRating = 0	   
	Else
	   If IsNull(rs("AverageRating")) Then
	      GetOverallProductRating = 0
	   Else
	      GetOverallProductRating = rs("AverageRating")
	   End if
	End If
	rs.close
	Set rs = Nothing

End Function
' PRV41 End
%>