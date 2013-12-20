<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!-- #Include File="checkdate.asp" -->
<!-- #Include File="header.asp" -->
<div id="pcMain">
	<table class="pcMainTable">
		<tr>
			<td>
				<h1><%response.write dictLanguage.Item(Session("language")&"_viewPostings_3")%></h1>
			</td>
		</tr>
		<tr>
			<td>
			<%
			
			'Allow upload: change to "0" to disallow
			AllowUpload="1"
			
			Dim rstemp, connTemp, mySQL
			Dim lngIDOrder,lngIDFeedback,lngIDComment
			call openDB()
			
			lngIDOrder=Clng(getUserInput(request("IDOrder"),0))-clng(scpre)
			lngIDFeedback=getUserInput(request("IDFeedback"),0)
			lngIDComment=getUserInput(request("IDComment"),0)
			
			 mySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDUser=" & session("IDCustomer")
			 set rstemp=connTemp.execute(mySQL)
			
			 if rstemp.eof then
			 call closedb()
			 response.redirect "userviewfeedback.asp?IDOrder=" & clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback & "&r=1&msg=" & dictLanguage.Item(Session("language")&"_editFeedback_b")
			 end if
			
			Dim strFDetails,dtComDate,ACount,ACount1
			
			'Create new feedback
			if (request("action")="update") and (request("rewrite")="0") then
				strFDetails=getUserInput(request("Details"),0)
				
				dtComDate=CheckDateSQL(now())
				
				MySQL="Update pcComments Set pcComm_EditedDate='" & dtComDate & "', pcComm_Details='" & strFDetails & "' where pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback
			
				set rstemp=connTemp.execute(mySQL)
				
				MySQL="Update pcComments Set pcComm_EditedDate='" & dtComDate & "' where pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDFeedback=" & lngIDFeedback
				set rstemp=connTemp.execute(mySQL)
				
				if AllowUpload="1" then
				ACount=getUserInput(request("ACount"),0)
				if ACount<>"" then
				ACount1=clng(ACount)
				For k=1 to ACount1
				if request("AC" & k)="1" then
				MySQL="update pcUploadFiles set pcUpld_IDFeedback=" & lngIDComment & " where pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
				set rstemp4=connTemp.execute(mySQL)
				else
				
				MySQL="Select * from pcUploadFiles where pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=" & lngIDComment
				set rstemp5=connTemp.execute(mySQL)
				if not rstemp5.eof then
				 Filename=rstemp5("pcUpld_Filename")
				 if Filename<>"" then
					QfilePath="Library/" & Filename
						findit = Server.MapPath(QfilePath)
					findit1 = findit
					Set fso = server.CreateObject("Scripting.FileSystemObject")
					Set f = fso.GetFile(findit)
					f.Delete
					Set fso = nothing
					Set f = nothing
					Err.number=0
					Err.Description=""
				 end if
					end if
			
				MySQL="delete from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment & " and pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
				set rstemp4=connTemp.execute(mySQL)
				
				end if
				next
				end if
				end if
				
				%>
				<div class="pcErrorMessage">
					<%response.write dictLanguage.Item(Session("language")&"_editFeedback_a")%>
				</div>
				<%end if%>
				<script language="JavaScript">
				<!--
					
				function Form1_Validator(theForm)
				{
				
							if (theForm.Details.value == "")
					{
								alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_h")%>");
								theForm.Details.focus();
								return (false);
					}
					
				return (true);
				}
				
				function newWindow(file,window) {
						msgWindow=open(file,window,'resizable=no,width=400,height=500');
						if (msgWindow.opener == null) msgWindow.opener = self;
				}
				
				function newWindow1(file,window) {
				catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
				if (catWindow.opener == null) catWindow.opener = self;
				}
				
				//-->
				</script>
				<%
					mySQL="select * from pcComments where pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder
				 	set rstemp=connTemp.execute(mySQL)
				 	Details=rstemp("pcComm_Details")
				 %>
 				<h2><%response.write dictLanguage.Item(Session("language")&"_editFeedback_c")%></h2>
				<form name="hForm" method="post" action="usereditComment.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
				<input type="hidden" name="Priority" value="<%=Priority%>">
				<input type="hidden" name="FStatus" value="<%=FStatus%>">
				<input type="hidden" name="FType" value="<%=FType%>">
				<input type="hidden" name="IDOrder" value="<%=clng(lngIDOrder)+scpre%>">
				<input type="hidden" name="IDFeedback" value="<%=lngIDFeedback%>">
				<input type="hidden" name="IDComment" value="<%=lngIDComment%>">
				<table class="pcShowContent">
					<tr>
						<td width="25%" align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewPostings_b")%>
						</td>
						<td width="75%">
							<b><%=scpre+clng(lngIDOrder)%></b>
						</td>
					</tr>
					<tr>
						<td align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_j")%>
						</td>
						<td>
							<%
								MySQL="Select * from pcComments where pcComm_IDParent=0 and pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDFeedback=" & lngIDFeedback
					 			set rstemp1=connTemp.execute(mySQL)
					 			response.write(rstemp1("pcComm_Description"))
							%>
						</td>
					</tr>
					<tr>
						<td align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_k")%>
						<br><br>
						<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=Details','window2')"></td>
						<td>
							<textarea name="Details" cols="40" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=Details%><%end if%></textarea>
						</td>
					</tr>
					<%if AllowUpload="1" then%>
					<tr>
						<td nowrap valign="top" align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%>
					</td>
					<td valign="top">
				<%MySQL="Select * from pcUploadFiles where pcUpld_IDFeedback=" & lngIDComment
				set rstemp4=connTemp.execute(mySQL)
				if rstemp4.eof then%>
				<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_6")%>
				<br>
				<%else
				ACount=0
				do while not rstemp4.eof
				ACount=ACount+1
				%>
				<input type="hidden" name="AID<%=ACount%>" value="<%=rstemp4("pcUpld_IDFile")%>">
				<input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">&nbsp;<%
				Filename= rstemp4("pcUpld_FileName")
				FileName = mid(FileName,instr(Filename,"_")+1,len(FileName))%>
				<%=FileName%>
				<br>
				<%rstemp4.MoveNext
				loop%>
				<input type="hidden" name="ACount" value="<%=ACount%>">
				<%end if%>
				<br>
				<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_7")%> <a href="#" onclick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=<%=lngIDComment%>&ReLink=<%=Server.URLencode("usereditcomment.asp?IDComment=" & lngIDComment & "&IDOrder=" & scpre+clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback)%>','window2')"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>
				</td>
				</tr>
				<%end if%>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<tr>
						<td></td>
						<td>
							<input type="submit" name="Submit" value="Update" class="submit2" onclick="document.hForm.rewrite.value='0';">
							<input type="button" name="back" value="Back" onClick="location='userviewfeedback.asp?IDOrder=<%=scpre+clng(lngIDOrder)%>&IDFeedback=<%=lngIDFeedback%>'">
							<%if session("IDOrder")>0 then%>
							<input type="button" name="go" value="Order Messages" onClick="location='userviewallposts.asp?IDOrder=<%=session("IDOrder")%>';">
							<%end if%>
							<input type="hidden" name="uploaded" value="">
							<input type="hidden" name="rewrite" value="1">
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>
</div>
<%call closedb()%><!-- #Include File="footer.asp" -->