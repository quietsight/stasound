<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/languages.asp"-->
<% 
Dim pageTitle, Section
pageTitle = "Product Reviews - General Settings"
pageIcon = "pcv4_icon_reviews.png"
Section = "reviews" 
%>
<!--#include file="adminheader.asp"-->
<% 
Dim rs, connTemp, query

call openDb()	

if request("action")="update" then
	pcv_RatingType=request("RatingType")
	if pcv_RatingType="" then
		pcv_RatingType="0"
	end if
	pcv_MainRateTxt1=request("MainRateTxt1")
	pcv_MainRateTxt2=request("MainRateTxt2")
	pcv_MainRateTxt3=request("MainRateTxt3")
	pcv_SubRateTxt1=request("SubRateTxt1")
	pcv_SubRateTxt2=request("SubRateTxt2")
	if pcv_MainRateTxt1<>"" then
		pcv_MainRateTxt1=replace(pcv_MainRateTxt1,"'","''")
	end if
	if pcv_MainRateTxt2<>"" then
		pcv_MainRateTxt2=replace(pcv_MainRateTxt2,"'","''")
	end if
	if pcv_MainRateTxt3<>"" then
		pcv_MainRateTxt3=replace(pcv_MainRateTxt3,"'","''")
	end if
	if pcv_SubRateTxt1<>"" then
		pcv_SubRateTxt1=replace(pcv_SubRateTxt1,"'","''")
	end if
	if pcv_SubRateTxt2<>"" then
		pcv_SubRateTxt2=replace(pcv_SubRateTxt2,"'","''")
	end if
		pcv_MaxRating=request("MaxRating")
	if pcv_MaxRating="" then
		pcv_MaxRating="10"
	end if
	pcv_Img1=request("Img1")
	if pcv_Img1="" then
		pcv_Img1="smileygreen.gif"
	end if
	pcv_Img2=request("Img2")
	if pcv_Img2="" then
		pcv_Img2="smileyred.gif"
	end if
	pcv_Img3=request("Img3")
	if pcv_Img3="" then
		pcv_Img3="fullstar.gif"
	end if
	pcv_Img4=request("Img4")
	if pcv_Img4="" then
		pcv_Img4="halfstar.gif"
	end if
	pcv_Img5=request("Img5")
	if pcv_Img5="" then
		pcv_Img5="emptystar.gif"
	end if
	pcv_Active=request("Active")
	if pcv_Active="" then
		pcv_Active="1"
	end if
	pcv_ShowRatSum=request("ShowRatSum")
	if pcv_ShowRatSum="" then
		pcv_ShowRatSum="0"
	end if
	pcv_RevCount=request("RevCount")
	if (pcv_RevCount<"0") or (not (IsNumeric(pcv_RevCount))) then
		pcv_RevCount="1"
	end if
	pcv_NeedCheck=request("NeedCheck")
	if pcv_NeedCheck="" then
		pcv_NeedCheck="0"
	end if
	pcv_LockPost=request("LockPost")
	if pcv_LockPost="" then
		pcv_LockPost="0"
	end if
	pcv_PostCount=request("PostCount")
	if (pcv_PostCount<"0") or (not (IsNumeric(pcv_PostCount))) then
		pcv_PostCount="1"
	end if
	pcv_CalMain=request("calmain")
	if pcv_CalMain="" then
		pcv_CalMain="0"
	end If
	
' prv41 start
	pcv_SendReviewReminder=request("sendreviewreminder")
	if pcv_SendReviewReminder="" then
		pcv_SendReviewReminder="0"
	end If

	If IsNumeric(request("sendreviewreminderdays")) And IsNull(request("sendreviewreminderdays"))=False then
	   pcv_SendReviewReminderDays=CLng(request("sendreviewreminderdays"))
	else
		pcv_SendReviewReminderDays="0"
	end If
	pcv_SendReviewReminderType=request("sendreviewremindertype")
	if pcv_SendReviewReminderType="" then
		pcv_SendReviewReminderType="0"
	end If
	pcv_SendReviewReminderFormat=request("sendreviewreminderformat")
	if pcv_SendReviewReminderFormat="" then
		pcv_SendReviewReminderFormat="0"
	end If
	If request("sendreviewremindertemplateopt")=0 Then
	   pcv_SendReviewReminderTemplate = ""
	Else
	   If Len(Trim(request("sendreviewremindertemplate")&""))>0 then
	      pcv_SendReviewReminderTemplate = replace(Trim(request("sendreviewremindertemplate")), "'", "''")
	   Else
	      pcv_SendReviewReminderTemplate = ""
	   End if
	End If
	pcv_RewardForReview=request("RewardForReview")
	if pcv_RewardForReview="" then
		pcv_RewardForReview="0"
	end If
	pcv_RewardForReviewURL=request("RewardForReviewURL")
	If IsNumeric(request("rewardForReviewFirstPts")) And IsNull(request("rewardForReviewFirstPts"))=False then
	   pcv_rewardForReviewFirstPts=CLng(request("rewardForReviewFirstPts"))
	else
		pcv_rewardForReviewFirstPts="0"
	end If
	If IsNumeric(request("rewardForReviewAdditionalPts")) And IsNull(request("rewardForReviewAdditionalPts"))=False then
	   pcv_rewardForReviewAdditionalPts=CLng(request("rewardForReviewAdditionalPts"))
	else
		pcv_rewardForReviewAdditionalPts="0"
	end If
	If IsNumeric(request("rewardForReviewMinLength")) And IsNull(request("rewardForReviewMinLength"))=False then
	   pcv_rewardForReviewMinLength=CLng(request("rewardForReviewMinLength"))
	else
		pcv_rewardForReviewMinLength="0"
	end If
	If IsNumeric(request("rewardForReviewMaxPts")) And IsNull(request("rewardForReviewMaxPts"))=False then
	   pcv_rewardForReviewMaxPts=CLng(request("rewardForReviewMaxPts"))
	else
		pcv_rewardForReviewMaxPts="0"
	end If

	pcv_DisplayRatings=request("displayratings")
	if pcv_DisplayRatings="" then
		pcv_DisplayRatings="0"
	end If
On Error goto 0
	query="UPDATE pcRevSettings Set pcRS_RatingType=" & pcv_RatingType & ",pcRS_MainRateTxt1='" & pcv_MainRateTxt1 & "',pcRS_MainRateTxt2='" & pcv_MainRateTxt2 & "',pcRS_MainRateTxt3='" & pcv_MainRateTxt3 & "',pcRS_SubRateTxt1='" & pcv_SubRateTxt1 & "',pcRS_SubRateTxt2='" & pcv_SubRateTxt2 & "',pcRS_MaxRating=" & pcv_MaxRating & ",pcRS_Img1='" & pcv_Img1 & "',pcRS_Img2='" & pcv_Img2 & "',pcRS_Img3='" & pcv_Img3 & "',pcRS_Img4='" & pcv_Img4 & "',pcRS_Img5='" & pcv_Img5 & "',pcRS_Active=" & pcv_Active & ",pcRS_ShowRatSum=" & pcv_ShowRatSum & ",pcRS_RevCount=" & pcv_RevCount & ",pcRS_NeedCheck=" & pcv_NeedCheck & ",pcRS_LockPost=" & pcv_LockPost & ",pcRS_PostCount=" & pcv_PostCount & ",pcRS_CalMain=" & pcv_CalMain & ", pcRS_SendReviewReminder=" & pcv_SendReviewReminder & ",pcRS_SendReviewReminderDays=" & pcv_SendReviewReminderDays & ",pcRS_SendReviewReminderType=" & pcv_SendReviewReminderType & ",pcRS_SendReviewReminderFormat=" & pcv_SendReviewReminderFormat & ",pcRS_SendReviewReminderTemplate='" & pcv_SendReviewReminderTemplate & "',pcRS_RewardForReview=" & pcv_RewardForReview & ",pcRS_RewardForReviewURL='" & Replace(pcv_RewardForReviewURL,"'","''") & "',pcRS_RewardForReviewFirstPts=" & pcv_RewardForReviewFirstPts & ",pcRS_RewardForReviewAdditionalPts=" & pcv_RewardForReviewAdditionalPts & ",pcRS_RewardForReviewMinLength=" & pcv_RewardForReviewMinLength & ",pcRS_RewardForReviewMaxPts=" & pcv_RewardForReviewMaxPts & ",pcRS_DisplayRatings=" & pcv_DisplayRatings
' prv41 end
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	msg="Product Reviews Settings were updated successfully!"
	msgType=1
	set rs=nothing
	
	if request("saveratingtype")<>request("ratingtype") then
		query="UPDATE Products SET pcProd_AvgRating=0;"
		set rs=connTemp.execute(query)
	end if
	
	if request("save_maxrating")<>pcv_MaxRating AND request("save_maxrating")<>"" then
		if pcv_MaxRating="5" then
			query="UPDATE pcReviewsData SET pcRD_Rate=Round((pcRD_Rate/2)+0.06,0);"
			query1="UPDATE pcReviews SET pcRev_MainDRate=Round((pcRev_MainDRate/2)+0.06,0);"
		else
			query="UPDATE pcReviewsData SET pcRD_Rate=pcRD_Rate*2;"
			query1="UPDATE pcReviews SET pcRev_MainDRate=pcRev_MainDRate*2;"
		end if
		set rs=connTemp.execute(query)
		set rs=nothing
		set rs=connTemp.execute(query1)
		set rs=nothing
	end if
end if

' prv41 start
query = "SELECT pcRS_RatingType,pcRS_MainRateTxt1,pcRS_MainRateTxt2,pcRS_MainRateTxt3,pcRS_SubRateTxt1,pcRS_SubRateTxt2,pcRS_MaxRating,pcRS_Img1,pcRS_Img2,pcRS_Img3,pcRS_Img4,pcRS_Img5,pcRS_Active,pcRS_ShowRatSum,pcRS_RevCount,pcRS_NeedCheck,pcRS_LockPost,pcRS_PostCount,pcRS_CalMain, pcRS_SendReviewReminder, pcRS_SendReviewReminderDays, pcRS_SendReviewReminderType, pcRS_SendReviewReminderFormat, pcRS_SendReviewReminderTemplate, pcRS_RewardForReview, pcRS_RewardForReviewURL, pcRS_RewardForReviewFirstPts, pcRS_RewardForReviewAdditionalPts, pcRS_RewardForReviewMinLength, pcRS_RewardForReviewMaxPts, pcRS_DisplayRatings FROM pcRevSettings;"
' prv41 end
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

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
if isNull(pcv_Active) or pcv_Active="" then
	pcv_Active="0"
end if
pcv_ShowRatSum=rs("pcRS_ShowRatSum")
pcv_RevCount=rs("pcRS_RevCount")
pcv_NeedCheck=rs("pcRS_NeedCheck")
pcv_LockPost=rs("pcRS_LockPost")
pcv_PostCount=rs("pcRS_PostCount")
pcv_CalMain=rs("pcRS_CalMain")
' prv41 start
pcv_sendReviewReminder=rs("pcRS_SendReviewReminder")
pcv_sendReviewReminderDays=rs("pcRS_SendReviewReminderDays")
pcv_sendReviewReminderType=rs("pcRS_SendReviewReminderType")
pcv_sendReviewReminderFormat=rs("pcRS_SendReviewReminderFormat")
pcv_sendReviewReminderTemplate=rs("pcRS_SendReviewReminderTemplate")
pcv_RewardForReview = rs("pcRS_RewardForReview")
pcv_RewardForReviewURL = rs("pcRS_RewardForReviewURL")
pcv_RewardForReviewFirstPts = rs("pcRS_RewardForReviewFirstPts")
pcv_RewardForReviewAdditionalPts = rs("pcRS_RewardForReviewAdditionalPts")
pcv_RewardForReviewMinLength = rs("pcRS_RewardForReviewMinLength")
pcv_RewardForReviewMaxPts = rs("pcRS_RewardForReviewMaxPts")
pcv_DisplayRatings = rs("pcRS_DisplayRatings")
' prv41 end
Set rs=Nothing
call closedb()
%>

<link href="../includes/spry/SpryTabbedPanels-Settings.css" rel="stylesheet" type="text/css" />
<script src="../includes/spry/SpryTabbedPanels.js" type="text/javascript"></script>
<script src="../includes/spry/SpryURLUtils.js" type="text/javascript"></script>
<script type="text/javascript"> 
	var params = Spry.Utils.getLocationParamsAsObject(); 
</script>


<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<script>
	function chgWin(file,window) {
		msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
	//-->
</script>
<form name="hForm" method="post" action="PrvSettings.asp?action=update" class="pcForms">
	<table class="pcCPcontent">	
		<tr>
			<td valign="top">
            <div id="TabbedPanels1" class="TabbedPanels">
              <ul class="TabbedPanelsTabGroup">
                <li class="TabbedPanelsTab" tabindex="100">Status &amp; General Settings</li>
                <li class="TabbedPanelsTab" tabindex="200">Write a Review Reminder</li>
                <li class="TabbedPanelsTab" tabindex="300">Rewards for Reviews</li>
                <li class="TabbedPanelsTab" tabindex="400">Other Settings</li>
              </ul>
              
                    <div class="TabbedPanelsContentGroup">
                        <div class="TabbedPanelsContent">
                        <table class="pcCPcontent">
                            <tr>
                                <td class="pcCPspacer" colspan="2"></td>
                            </tr>
                            <tr> 
                                <td colspan="2">
                                <h2>Product Reviews System Status</h2>
                                <input type="radio" name="active" value="1" <%if pcv_Active="1" then%>checked<%end if%> class="clearBorder"> Turn Product Reviews On
                                <input type="radio" name="active" value="0" <%if (pcv_Active="0") or (pcv_Active="") then%>checked<%end if%> class="clearBorder"> Turn Product Reviews Off
								<input type="hidden" name="saveratingtype" value="<%=pcv_RatingType%>">
								</td>
                            </tr>
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td colspan="2">
                                <h2>Product Reviews Rating Settings <a href="http://wiki.earlyimpact.com/productcart/products_reviews" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information about this feature" border="0"></a></h2>
                               NOTE: changing this option after product reviews have been submitted <strong>could generate some issues</strong> (e.g. average rating not shown). We recommend that this option is not changed, or that you remove existing reviews if you decide to change it. <a href="http://wiki.earlyimpact.com/productcart/products_reviews" target="_blank">More information</a>.</td>
                            </tr>
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td style="border-bottom: 1px dashed #e1e1e1;">
                                <input type="radio" name="ratingtype" value="0" <%if pcv_RatingType="0" then%>checked<%end if%> class="clearBorder" onClick="document.getElementById('rateByFeelingTable').style.display=''; document.getElementById('rateByMarksTable').style.display='none'"><strong>Rate by 'Feeling'</strong>
                                <td class="pcCPnotes" style="border-bottom: 1px dashed #e1e1e1;">Only 2 choices: Thumbs Up/Thumbs Down - e.g: Like it/Don't like it, Good/Bad</td>
                            </tr>
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td colspan="2">
                                    <table id="rateByFeelingTable" style="display: <%if pcv_RatingType<>"0" then%>none<%end if%>;" class="pcCPcontent">
                                        <tr> 
                                            <td width="20%" align="right" nowrap>Display Rating Text:</td>
                                            <td width="80%"><input name="mainratetxt1" type="text" size="60" value="<%=pcv_MainRateTxt1%>"></td>
                                        </tr>
                                        <tr>
                                          <td align="right"><strong>Rating Names</strong>:</td>
                                          <td><span class="pcSmallText">Sub-rating: when rating product properties (e.g. comfort, style, etc.)</span></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Overall &quot;Thumbs Up&quot;:</td>
                                            <td><input name="mainratetxt2" type="text" size="60" value="<%=pcv_MainRateTxt2%>"></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Overall &quot;Thumbs Down&quot;:</td>
                                            <td><input name="mainratetxt3" type="text" size="60" value="<%=pcv_MainRateTxt3%>"></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Sub-Rating &quot;Thumbs Up&quot;:&nbsp;</td>
                                            <td><input name="subratetxt1" type="text" size="60" value="<%=pcv_SubRateTxt1%>"></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Sub-Rating &quot;Thumbs Down&quot;:</td>
                                            <td><input name="subratetxt2" type="text" size="60" value="<%=pcv_SubRateTxt2%>"></td>
                                        </tr>
                                        <tr> 
                                            <td align="right"><strong>Rating Images</strong>:</td>
                                            <td>&nbsp;</td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Overall &quot;Thumbs Up&quot; Image:</td>
                                            <td><img src="../pc/catalog/<%=pcv_Img1%>" border="0">&nbsp;<input name="img1" type="text" value="<%=pcv_Img1%>" size="50">&nbsp;<a href="javascript:chgWin('../pc/imageDir.asp?ffid=img1&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" border=0 hspace="3"></a></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">Overall &quot;Thumbs Down&quot; Image:</td>
                                            <td><img src="../pc/catalog/<%=pcv_Img2%>" border="0">&nbsp;<input name="img2" type="text" value="<%=pcv_Img2%>" size="50">&nbsp;<a href="javascript:chgWin('../pc/imageDir.asp?ffid=img2&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" border=0 hspace="3"></a></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="2" class="pcCPspacer"></td>
                                        </tr>
                                    </table>
                            	</td>
                            </tr>
                            <tr> 
                                <td style="border-bottom: 1px dashed #e1e1e1;"><input type="radio" name="ratingtype" value="1" <%if pcv_RatingType="1" then%>checked<%end if%> class="clearBorder" onClick="document.getElementById('rateByFeelingTable').style.display='none'; document.getElementById('rateByMarksTable').style.display=''"><strong>Rate by Marks</strong></td>
                                <td class="pcCPnotes" style="border-bottom: 1px dashed #e1e1e1;">E.g:  From 1 to 10, where 1 = worst, 10 = best</td>
                            </tr>
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td colspan="2">
                                    <table id="rateByMarksTable" style="display: <%if pcv_RatingType<>"1" then%>none<%end if%>;" class="pcCPcontent">
                                        <tr> 
                                            <td align="right" width="20%">Rating Marks:</td>
                                            <td width="80%">From 1 to <select size="1" name="maxrating">
                                            <option value="5" <%if pcv_MaxRating<>"10" then%>selected<%end if%>>5</option>
                                            <option value="10" <%if pcv_MaxRating="10" then%>selected<%end if%>>10</option>
                                            </select>
                                            <input type="hidden" name="save_maxrating" value="<%=pcv_MaxRating%>">
                                            </td>
                                        </tr>
                                        <tr> 
                                          <td align="right" valign="top">Overall Product Rating:</td>
                                          <td>

                                            <%
											call openDb()
											Dim pcIntDataEmpty, pcIntFieldCount
											query="SELECT pcRF_IDField,pcRF_Name,pcRF_Type,pcRF_Active,pcRF_Required,pcRF_Order FROM pcRevFields WHERE pcRF_Type=4 AND pcRF_Active=1 ORDER BY pcRF_Order asc,pcRF_IDField ASC"
											set rs=server.CreateObject("ADODB.RecordSet")
											set rs=connTemp.execute(query)
												
											if rs.eof then
												pcIntDataEmpty=1
											else
												pcIntDataEmpty=0
												pcArray=rs.getRows()
												pcIntFieldCount=(ubound(pcArray,2)+1)
											end if
											set rs=nothing
											call closeDb()
											%>
                                          
                                          
                                            <input type="radio" name="calmain" value="0"<%if pcv_CalMain<>"1" and pcIntDataEmpty=0 then%> checked<%end if%><%if pcIntDataEmpty="1" then%> disabled<%end if%> class="clearBorder"> Will be calculated automatically by averaging <a href="prv_FieldManager.asp" target="_blank">sub-ratings</a>
                                            <br>
                                            <span class="pcSmallText" style="padding-left: 24px;">E.g. quality, comfort, value for the money, etc. <a href="http://wiki.earlyimpact.com/productcart/products_reviews" target="_blank">Learn more</a> about how sub-ratings work.</span><br />
                                            <input type="radio" name="calmain" value="1" <%if pcv_CalMain="1" or pcIntDataEmpty=1 then%>checked<%end if%> class="clearBorder"> Will be set on the overall product
                                            <br>
                                            <span class="pcSmallText" style="padding-left: 24px;">E.g. <em>Rate this product</em>. Sub-ratings - if any - do not affect it</span><br />
                                            <% if pcIntDataEmpty = 1 then%>
                                            	<div class="pcCPmessageInfo">There are currently no active 'Mark Rating' <a href="prv_FieldManager.asp" target="_blank">fields</a>, so the overall product rating cannot be calculated by averaging sub-ratings (that option has been disabled). <a href="http://wiki.earlyimpact.com/productcart/products_reviews" target="_blank">Learn more</a> about how sub-ratings work.</div>                                         
                                            <% else %>
                                                <div class="pcCPmessageInfo">There are currently <%=pcIntFieldCount%> active 'Mark Rating' <a href="prv_FieldManager.asp" target="_blank">fields</a>, which can be used to calculate the overall product rating (1st option above).</div>
                                            <% end if %>
                                            </td>
                                        </tr>
                                        <tr> 
                                            <td align="right"><strong>Images</strong>: </td>
                                            <td>E.g. 3.5 out of 5 will be shown as: <img src="../pc/catalog/<%=pcv_Img3%>" border="0"><img src="../pc/catalog/<%=pcv_Img3%>" border="0"><img src="../pc/catalog/<%=pcv_Img3%>" border="0"><img src="../pc/catalog/<%=pcv_Img4%>" border="0"><img src="../pc/catalog/<%=pcv_Img5%>" border="0"></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">&quot;Full Mark&quot; Image:</td>
                                            <td> 
                                            <img src="../pc/catalog/<%=pcv_Img3%>" border="0">                                
                                            <input name="img3" type="text" value="<%=pcv_Img3%>" size="50">&nbsp;<a href="javascript:chgWin('../pc/imageDir.asp?ffid=img3&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">&quot;1/2 Mark&quot; Image:</td>
                                            <td> 
                                            <img src="../pc/catalog/<%=pcv_Img4%>" border="0">
                                            <input name="img4" type="text" value="<%=pcv_Img4%>" size="50">&nbsp;<a href="javascript:chgWin('../pc/imageDir.asp?ffid=img4&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a></td>
                                        </tr>
                                        <tr> 
                                            <td align="right">&quot;Empty Mark&quot; Image:</td>
                                            <td> 
                                            <img src="../pc/catalog/<%=pcv_Img5%>" border="0">
                                            <input name="img5" type="text" value="<%=pcv_Img5%>" size="50">&nbsp;<a href="javascript:chgWin('../pc/imageDir.asp?ffid=img5&fid=hForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="pcCPspacer" colspan="2"></td>
                            </tr>  
                        </table>							
                    </div>
                                                    
                    <div class="TabbedPanelsContent">
                        <table class="pcCPcontent">
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right" width="10%" nowrap>Send a &quot;Write a Review&quot; reminder:</td>
                                <td width="90%"><input type="radio" name="sendreviewreminder" value="1" <%if pcv_sendReviewReminder="1" then%>checked<%end if%> class="clearBorder"> Yes&nbsp;<input type="radio" name="sendreviewreminder" value="0" <%if pcv_sendReviewReminder="0" then%>checked<%end if%> class="clearBorder"> No &nbsp;<a href="http://wiki.earlyimpact.com/productcart/products_reviews#write_a_review_reminder" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information about this feature" border="0"></a></td>
                            </tr>
                            <tr>
                               <td align="right">When to send it:</td>
                               <td><input type="text" name="sendreviewreminderdays" value="<%=pcv_sendReviewReminderDays%>" size="4" style="text-align: right"> days after <input type="radio" name="sendreviewremindertype" value="0" <% If pcv_sendReviewReminderType=0 Then response.write "checked" %> class="clearBorder"> Order Processed&nbsp;&nbsp;<input type="radio" name="sendreviewremindertype" value="1" <% If pcv_sendReviewReminderType=1 Then response.write "checked" %> class="clearBorder"> Order Shipped</td>
                            </tr>
                            <tr>
                               <td align="right">Format:</td>
                               <td><input type="radio" name="sendreviewreminderformat" value="0" <% If pcv_sendReviewReminderFormat=0 Then response.write "checked" %> class="clearBorder"> Text&nbsp;&nbsp;<input type="radio" name="sendreviewreminderformat" value="1" <% If pcv_sendReviewReminderFormat=1 Then response.write "checked" %> class="clearBorder"> HTML</td>
                            </tr>
                            <tr valign="top">
                               <td align="right">Template:</td>
                               <td>
                                  <input type="radio" name="sendreviewremindertemplateopt" value="0" <% If Len(pcv_sendReviewReminderTemplate)=0 Then response.write "checked" %> class="clearBorder"> Use default template | <a href="javascript:chgWin('ReviewsEmailTest.asp','window2')">View &amp; Test</a><br />
                                  <input type="radio" name="sendreviewremindertemplateopt" value="1" <% If Len(pcv_sendReviewReminderTemplate)>0 Then response.write "checked" %> class="clearBorder"> Use custom template: 
                                  <input name="sendreviewremindertemplate" value="<% response.write pcv_sendReviewReminderTemplate %>" size="20" maxlength="255"> <a href="javascript:chgWin('ReviewsEmailTest.asp','window2')">View &amp; Test</a> | <a href="javascript:chgWin('uploadresize/fileuploada.asp?ffid=sendreviewremindertemplate&fid=hForm','window2')">Upload New</a>
                            </tr>
                            <tr>
                                <td class="pcCPspacer" colspan="2"></td>
                            </tr>  
                        </table>							
                    </div>
                                                    
                    <div class="TabbedPanelsContent">
                        <table class="pcCPcontent">
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">Reward customer for writing a review?</td>
                                <td><input type="radio" name="RewardForReview" value="1" <%if pcv_RewardForReview="1" then%>checked<%end if%> class="clearBorder"> Yes&nbsp;<input type="radio" name="RewardForReview" value="0" <%if pcv_RewardForReview="0" then%>checked<%end if%> class="clearBorder"> No</td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">URL to page that contains program details:</td>
                                <td><input type="text" name="RewardForReviewURL" size="50" maxlength="255" value="<% = pcv_RewardForReviewURL %>"><span id="urlerr" style="color: #f00;"></span><br /><span class="pcSmallText"><br />Note: This could be a "content page". Please enter a full url (e.g. http://www.earlyimpact.com/) </span></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">Points to be awarded on first review:</td>
                                <td><input type="text" name="RewardForReviewFirstPts" size="5" maxlength="5" value="<% = pcv_RewardForReviewFirstPts %>"></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">Points to be awarded on additional reviews:</td>
                                <td><input type="text" name="RewardForReviewAdditionalPts" size="5" maxlength="5" value="<% = pcv_RewardForReviewAdditionalPts %>"></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">Minimum length of review:</td>
                                <td><input type="text" name="RewardForReviewMinLength" size="5" maxlength="5" value="<% = pcv_RewardForReviewMinLength %>"></td>
                            </tr>
                            <tr valign="top"> 
                                <td align="right">Max Points to be awarded per customer:</td>
                                <td><input type="text" name="RewardForReviewMaxPts" size="5" maxlength="5" value="<% = pcv_RewardForReviewMaxPts %>"></td>
                            </tr>
                        <tr>
                            <td class="pcCPspacer" colspan="2"></td>
                        </tr>  
                        </table>							
                    </div>
                                                    
                    <div class="TabbedPanelsContent">
                        <table class="pcCPcontent">
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td colspan="2"><h2>Display Settings:</h2></td>
                            </tr>
                            <tr> 
                                <td align="right" valign="top"><input type="checkbox" name="showratsum" value="1" <%if pcv_ShowRatSum="1" then%>checked<%end if%> class="clearBorder"></td>
                                <td>Show Rating Summary at the top of the product details page.<br><span class="pcSmallText">Shown right below the Part Number (SKU), with a link to the Product Reviews section of the same page. If not shown, <a href="http://wiki.earlyimpact.com/productcart/products_reviews#troubleshooting" target="_blank">read this.</a></span></td>
                            </tr>
                            <tr> 
                                <td align="right" valign="top"><input type="checkbox" name="displayratings" value="1" <%if pcv_DisplayRatings="1" then%>checked<%end if%> class="clearBorder"></td>
                                <td>Display average rating on category &amp; search results pages.<br><span class="pcSmallText">Shown right below the product prices, wherever products are listed.</span></td>
                            </tr>
                            <tr> 
                                <td align="right" valign="top"><input type="text" name="revcount" value="<%=pcv_RevCount%>" size="4" style="text-align: right"></td>
                                <td>reviews will be displayed on the product details page. Other reviews will be displayed on a separate page.</td>
                            </tr>
                            <tr> 
                                <td align="right" valign="top"><input type="checkbox" name="needcheck" value="1" <%if pcv_NeedCheck="1" then%>checked<%end if%> class="clearBorder"></td>
                                <td>All reviews must be approved by the store manager before they are made public.</td>
                            </tr>
                            <tr> 
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                            <tr> 
                                <td colspan="2"><h2>Lock Posting Options:</h2></td>
                            </tr>
                            <tr> 
                                <td nowrap align="right">Max <input type="text" name="postcount" value="<%=pcv_PostCount%>" size="4" style="text-align: right"></td>
                                <td>reviews submitted by the same customer, for the same product.</td>
                            </tr>
                            <tr> 
                                <td align="right"><input type="radio" name="lockpost" value="0" <%if pcv_LockPost="0" then%>checked<%end if%> class="clearBorder"></td>
                                <td>Lock by using the customer's IP Address</td>
                            </tr>
                            <tr> 
                                <td align="right"><input type="radio" name="lockpost" value="1" <%if pcv_LockPost="1" then%>checked<%end if%> class="clearBorder"></td>
                                <td>Lock by using cookies</td>
                            </tr>
                            <tr> 
                                <td align="right"><input type="radio" name="lockpost" value="2" <%if pcv_LockPost="2" then%>checked<%end if%> class="clearBorder"></td>
                                <td>Lock using both methods</td>
                            </tr>
                        <tr>
                            <td class="pcCPspacer" colspan="2"></td>
                        </tr>  
                        </table>
                        </div>				
                    </div>	
                </div>		
				<script type="text/javascript">
					var TabbedPanels1 = new Spry.Widget.TabbedPanels("TabbedPanels1", {defaultTab: params.tab ? params.tab : 0});
                </script>
			</td>
		</tr>
		<tr>
            <td colspan="2"><hr></td>
		</tr>
		<tr> 
		<td align="center" colspan="2"> 
            <input type="submit" name="Submit" value="Update Settings" onclick="return(checkRatingType())" class="submit2">&nbsp;
            <input type="button" value="Manage Fields" onClick="location.href='prv_FieldManager.asp'">&nbsp;
            <input type="button" value="Bad Words" onClick="location.href='prv_ManageBadWords.asp'">&nbsp;
            <input type="button" value="Product-specific Settings" onClick="location.href='prv_SpecialPrd.asp'">&nbsp;
            <input type="button" value="Product Exclusions" onClick="location.href='prv_PrdExc.asp'">
			<script>
			function checkRatingType()
			{
				var tmpvalue=0;
				if (document.hForm.ratingtype[0].checked==true)
				{
					tmpvalue=document.hForm.ratingtype[0].value;
				}
				else
				{
					tmpvalue=document.hForm.ratingtype[1].value;
				}
				if (document.hForm.saveratingtype.value	!= tmpvalue)
				{
					return(confirm('Are you sure you want to change the main Rating Type? Average product ratings WILL BE REMOVED from the system when you switch this setting as they are calculated differently for each Rating Type. This action cannot be undone.'));
				}
				return(true);
			}
			</script>
        </td>
		</tr>
		<tr> 
            <td align="center" colspan="2"> 
                <input type="button" value="View Pending Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=1'">&nbsp;
                <input type="button" value="View Live Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=2'">&nbsp;                
            	<input type="button" name="back" value="Back" onClick="javascript:history.back()">
            </td>
		</tr>
	</table>
</form>
<!--#include file="adminfooter.asp"-->