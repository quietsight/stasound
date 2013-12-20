<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3
section="specials"
pageIcon="pcv4_icon_salesManager.png"
if request("e")="1" then
	pageTitle="View/Edit Sale Details"
else
	pageTitle="Sale Details"
end if
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/emailsettings.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<% 
Dim query, conntemp, rs

function ShowDateTimeFrmt(datestring)
Dim tmp1,tmp2
	tmp1=split(datestring," ")
	if scDateFrmt="DD/MM/YY" then
		tmp2=day(tmp1(0))&"/"&month(tmp1(0))&"/"&year(tmp1(0))
	else
		tmp2=month(tmp1(0))&"/"&day(tmp1(0))&"/"&year(tmp1(0))
	end if
	if instr(datestring," ") then
		tmp2=tmp2 & " " & tmp1(1) & tmp1(2)
	end if
	ShowDateTimeFrmt=tmp2
end function

call opendb()


	pcSCID=request("id")
	if pcSCID="" then
		call closedb()
		response.redirect "sm_sales.asp"
	else
		if Not (IsNumeric(pcSCID)) then
			call closedb()
			response.redirect "sm_sales.asp"
		end if
	end if
	
	msg=""
	
	if request("action")="update" then
	
		pcSaleID=request("pcSaleID")
		salename=replace(request("salename"),"'","''")
		saledesc=replace(request("saledesc"),"'","''")
		saleicon=replace(request("saleicon"),"'","''")
		query="UPDATE pcSales_Completed SET pcSC_SaveName='" & salename & "',pcSC_SaveDesc='" & saledesc & "',pcSC_SaveIcon='" & saleicon & "' WHERE pcSC_ID=" & pcSCID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		query="UPDATE pcSales SET pcSales_Name='" & salename & "',pcSales_Desc='" & saledesc & "',pcSales_ImgURL='" & saleicon & "' WHERE pcSales_ID=" & pcSaleID & ";"
		set rs=connTemp.execute(query)
		set rs=nothing
		
		msg="The Sale details have been updated successfully!"
	
	end if



%>
<!--#include file="AdminHeader.asp"-->
<%'// Initialize the Prototype.js files
Dim pcv_strPrototype
Response.Write(pcf_InitializePrototype())
%>
<%
	query="SELECT pcSC_Status,pcSC_StartedDate,pcSC_StoppedDate,pcSC_BUTotal,pcSC_SaveName,pcSC_SaveDesc,pcSC_SaveTech,pcSC_SaveIcon,pcSales.pcSales_ID,pcSC_Archived FROM pcSales INNER JOIN pcSales_Completed ON pcSales.pcSales_ID=pcSales_Completed.pcSales_ID WHERE pcSC_ID=" & pcSCID & ";"
	set rs=connTemp.execute(query)
	if rs.eof then
		set rs=nothing
		call closedb()
		response.redirect "sm_sales.asp"
	else
		tmpSaleStatus=rs("pcSC_Status")
		tmpStartedDate=rs("pcSC_StartedDate")
		tmpStoppedDate=rs("pcSC_StoppedDate")
		tmpUPrds=rs("pcSC_BUTotal")
		tmpSaleName=rs("pcSC_SaveName")
		tmpSaleDesc=rs("pcSC_SaveDesc")
		tmpSaleDetails=rs("pcSC_SaveTech")
		tmpSaleIcon=rs("pcSC_SaveIcon")
		pcSaleID=rs("pcSales_ID")
		tmpArchived=rs("pcSC_Archived")
		if tmpArchived="" then
			tmpArchived="0"
		end if
		set rs=nothing
		
	%>
	<%if request("e")="1" then%>
	<script language="JavaScript">
	<!--
	function Form1_Validator(theForm)
	{

	if (theForm.salename.value == "")
 	{
		    alert("Please enter 'Sale Name'");
			theForm.salename.focus();
		    return (false);
	}
	
 	if (theForm.saledesc.value == "")
 	{
		    alert("Please enter 'Sale Description'");
			theForm.saledesc.focus();
		    return (false);
	}
	if (theForm.saleicon.value == "")
 	{
		    alert("Please enter 'Sale Icon'");
			theForm.saleicon.focus();
		    return (false);
	}
	
	return (true);
	}

	function newWindow(file,window)
	{
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
		
	function chgWin1(file,window)
	{
		msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=500,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}

	//-->
	</script> 
	
	<form name="UpdateForm" id="UpdateForm" method="post" action="sm_saledetails.asp?action=update" onsubmit="return Form1_Validator(this)" class="pcForms">
	<input type=hidden name="e" value="1">
	<input type="hidden" name="id" value="<%=pcSCID%>">
	<input type="hidden" name="pcSaleID" value="<%=pcSaleID%>">
	<%end if%>
	<table class="pcCPcontent">
	<%if msg<>"" then%>
		<tr>
		<td colspan="2">
			<div class="pcCPmessageSuccess">
				<%=msg%>
			</div>
		</td>
		</tr>
	<%end if%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td width="15%" nowrap>Sale Name:<%if request("e")="1" then%><img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5"><%end if%></td>
		<td>
		<%if request("e")="1" then%>
			<input type="text" name="salename" id="salename" value="<%=tmpSaleName%>" size="50">
		<%else%>
			<b><%=tmpSaleName%></b>
		<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td nowrap>Sale Description:<%if request("e")="1" then%><img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5"><%end if%></td>
		<td>
		<%if request("e")="1" then%>
		<textarea name="saledesc" id="saledesc" cols="50" rows="10"><%=tmpSaleDesc%></textarea>
		<%else%>
		<%=tmpSaleDesc%>
		<%end if%>
		</td>
	</td>
	</tr>
	<%if request("e")="1" then%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
	<td>                          
		Sale Icon:<%if request("e")="1" then%><img src="images/pc_required.gif" alt="required field" width="9" height="9" hspace="5"><%end if%>
	</td>
	<td>
		<input type="text" name="saleicon" id="saleicon" value="<%=tmpSaleIcon%>" size="50"> <a href="javascript:;" onClick="chgWin1('../pc/imageDir.asp?ffid=saleicon&fid=UpdateForm','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>&nbsp;
		<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
		<a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=460')"><img src="images/sortasc_blue.gif" alt="Upload your images"></a>&nbsp;
		<%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%><a href="javascript:;" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=460')"><%=dictLanguageCP.Item(Session("language")&"_cpCommon_303")%></a>.
     </td>
	</tr>
	<% if trim(tmpSaleIcon)<>"" then %>
	<tr>
		<td>&nbsp;</td>
		<td>Currently using this icon: <img src="../pc/catalog/<%=tmpSaleIcon%>"></td>
	</tr>
	<% end if %>
	<%end if%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Sale Status:</td>
		<td><b>
			<%Select Case tmpSaleStatus
			Case "1": response.write "Started"
			Case "2": response.write "Live"
			Case "3": response.write "Stopped"
			Case "4": response.write "Completed"
			Case Else: response.write "N/A"
			End Select%>
			</b>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Products:</td>
		<td>
			<%if tmpArchived="1" then%>
				<i>NOTE: this sales has been archived. The system no longer has detailed information on the products that were included in the sale</i>
			<%else%>
				<a href="sm_addedit_S1.asp?c=new&a=rev&id=<%=pcSaleID%>&b=<%if request("e")="1" then%>5<%else%>7<%end if%>&scid=<%=pcSCID%>">Click here</a> to view list products included in this sale
			<%end if%>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr valign="top">
		<td>Details:</td>
		<td><%=tmpSaleDetails%></td>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2">The sale was started on <b><%=ShowDateTimeFrmt(tmpStartedDate)%></td>
	</tr>
	<tr>
		<td colspan="2">It affected <b><%=tmpUPrds%></b> product(s) in your store.</td>
	</tr>
	<%if tmpStoppedDate<>"" AND (not IsNull(tmpStoppedDate)) then%>
	<tr>
		<td colspan="2">The sale was stopped on <b><%=ShowDateTimeFrmt(tmpStoppedDate)%></b></td>
	</tr>
	<%end if%>
	<tr>
		<td colspan="2" class="pcCPspacer"><hr></td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<%if request("e")="1" then%><input type="submit" name="submit1" value=" Save " class="submit2">&nbsp;<%end if%><input type="button" name="Go" value=" View Current & Completed Sales " onclick="location='sm_sales.asp';" <%if request("e")="1" then%>class="ibtnGrey"<%else%>class="submit2"<%end if%>>
		</td>
	</tr>
	</table>
	<%if request("e")="1" then%>
	</form>
	<%end if%>	
	<%end if
%>
<% 
call closeDb()
set rs= nothing
%>
<!--#include file="AdminFooter.asp"-->