<%
'**** Get General Settings
query = "SELECT TOP 1 pcRS_Active, pcRS_DisplayRatings,pcRS_RatingType,pcRS_MainRateTxt1,pcRS_MainRateTxt2,pcRS_MainRateTxt3,pcRS_SubRateTxt1,pcRS_SubRateTxt2,pcRS_MaxRating,pcRS_Img1,pcRS_Img2,pcRS_Img3,pcRS_Img4,pcRS_Img5,pcRS_Active,pcRS_ShowRatSum,pcRS_RevCount,pcRS_NeedCheck,pcRS_LockPost,pcRS_PostCount,pcRS_CalMain, pcRS_RewardForReviewMinLength FROM pcRevSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pShowAvgRating = False
pRSActive = False
if not rs.eof then
	pRSActive = CBool(CLng(rs("pcRS_Active")) <> 0)
	pShowAvgRating = CBool(CLng(rs("pcRS_DisplayRatings")) <> 0)
	pcv_RatingType=rs("pcRS_RatingType")
	pcv_MainRateTxt1=rs("pcRS_MainRateTxt1")
	pcv_MainRateTxt2=rs("pcRS_MainRateTxt2")
	pcv_MainRateTxt3=rs("pcRS_MainRateTxt3")
	pcv_SubRateTxt1=rs("pcRS_SubRateTxt1")
	pcv_SubRateTxt2=rs("pcRS_SubRateTxt2")
	pcv_MaxRating=rs("pcRS_MaxRating")
	pcv_Img1=rs("pcRS_Img1")
	pcv_Img2=rs("pcRS_Img2")
	pcv_Img3=rs("pcRS_Img3")
	pcv_Img4=rs("pcRS_Img4")
	pcv_Img5=rs("pcRS_Img5")
	pcv_Active=rs("pcRS_Active")
	if isNULL(pcv_Active) or pcv_Active="" then
		pcv_Active="0"
	end if
	pcv_ShowRatSum=rs("pcRS_ShowRatSum")
	pcv_RevCount=rs("pcRS_RevCount")
	pcv_NeedCheck=rs("pcRS_NeedCheck")
	pcv_LockPost=rs("pcRS_LockPost")
	pcv_PostCount=rs("pcRS_PostCount")
	pcv_CalMain=rs("pcRS_CalMain")
	pcv_RewardForReviewMinLength=rs("pcRS_RewardForReviewMinLength")
end if
set rs=nothing
%>