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
			
			Dim rs, connTemp, query
			call openDB()
			
			LngIdOrder=clng(getUserInput(request("IDOrder"),0))-clng(scpre)
			intIdFeedback=getUserInput(request("IDFeedback"),0)
			
			query="SELECT * FROM pcComments WHERE pcComm_IDFeedback=" & intIdFeedback & " and pcComm_IDParent=0 and pcComm_IDOrder=" & LngIdOrder & " and pcComm_IDUser=" & session("IDCustomer")
			set rs=connTemp.execute(query)
			 
			if rs.eof then
				call closedb()
				response.redirect "userviewfeedback.asp?IDOrder=" & LngIdOrder & "&IDFeedback=" & intIdFeedback & "&r=1&msg="&dictLanguage.Item(Session("language")&"_editFeedback_b")
			end if
			
			'Update feedback
			if (request("action")="update") and (request("rewrite")="0") then
				LngIdOrder=getUserInput(request("IDOrder"),0)
				strFDesc=getUserInput(request("Description"),0)
				strFDetails=getUserInput(request("Details"),0)
				intFStatus=getUserInput(request("FStatus"),0)
				intFType=getUserInput(request("FType"),0)
				intPriority=getUserInput(request("Priority"),0)
				
				dtComDate=CheckDateSQL(now())
				
				if scDB="SQL" then
					query="UPDATE pcComments SET pcComm_EditedDate='" & dtComDate & "',pcComm_FType=" & intFType & ",pcComm_FStatus=" & intFStatus & ",pcComm_Priority=" & intPriority & ",pcComm_Description='" & strFDesc & "',pcComm_Details='" & strFDetails & "' WHERE pcComm_IDOrder=" & (LngIdOrder)-scpre & " and pcComm_IDFeedback=" & intIdFeedback
				else
					query="UPDATE pcComments SET pcComm_EditedDate=#" & dtComDate & "#,pcComm_FType=" & intFType & ",pcComm_FStatus=" & intFStatus & ",pcComm_Priority=" & intPriority & ",pcComm_Description='" & strFDesc & "',pcComm_Details='" & strFDetails & "' WHERE pcComm_IDOrder=" & int(LngIdOrder)-scpre & " and pcComm_IDFeedback=" & intIdFeedback
				end if
			
				set rs=connTemp.execute(query)
			
				if AllowUpload="1" then
					ACount=getUserInput(request("ACount"),0)
					if ACount<>"" then
						ACount1=clng(ACount)
						For k=1 to ACount1
							if request("AC" & k)="1" then
								query="UPDATE pcUploadFiles SET pcUpld_IDFeedback=" & intIdFeedback & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
								set rs=connTemp.execute(query)
							else
								query="SELECT pcUpld_Filename FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=" & intIdFeedback
								set rs=connTemp.execute(query)
								if not rs.eof then
									strFileName=rs("pcUpld_Filename")
									if strFileName<>"" then
										QfilePath="Library/" & strFileName
										findit = Server.MapPath(QfilePath)
										Set fso = server.CreateObject("Scripting.FileSystemObject")
										Set f = fso.GetFile(findit)
										f.Delete
										Set fso = nothing
										Set f = nothing
										Err.number=0
										Err.Description=""
									end if
								end if
			
								query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & intIdFeedback & " and pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
								set rs=connTemp.execute(query)
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
				<%if session("UserType")=3 then%>
						if (theForm.FType.value == "")
					{
								alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_d")%>");
								theForm.FType.focus();
								return (false);
					}
							if (theForm.Priority.value == "")
					{
								alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_e")%>");
								theForm.Priority.focus();
								return (false);
					}
								if (theForm.FStatus.value == "")
					{
								alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_f")%>");
								theForm.FStatus.focus();
								return (false);
					}
				<%end if%>
				
							if (theForm.Description.value == "")
					{
								alert("<%response.write dictLanguage.Item(Session("language")&"_editFeedback_g")%>");
								theForm.Description.focus();
								return (false);
					}
					
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
				//-->
				</script>
				
				<% query="SELECT pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details FROM pcComments WHERE pcComm_IDFeedback=" & intIdFeedback & ";"
				set rs=connTemp.execute(query)
				intFType=rs("pcComm_FType")
				intFStatus=rs("pcComm_FStatus")
				intPriority=rs("pcComm_Priority")
				strDesc=rs("pcComm_Description")
				strDetails=rs("pcComm_Details")
				%>
				
				<form name="hForm" method="post" action="usereditFeedback.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
				<input type="hidden" name=IDOrder value="<%=scpre+clng(LngIdOrder)%>">
				<input type="hidden" name=IDFeedback value="<%=intIdFeedback%>">
				<h2><%response.write dictLanguage.Item(Session("language")&"_editFeedback_c")%></h2>
				<table class="pcShowContent">
				<tr>
					<td width="25%" align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_b")%>
					</td>
					<td width="75%">
						<b><%=scpre+clng(LngIdOrder)%></b>
					</td>
				</tr>
				<%if (session("UserType")=3) then%>  
					<tr>
						<td align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_h")%>
						</td>
						<td>
						<select name="FType">
							<option value=""></option>
							<% query="SELECT pcFType_IDType,pcFType_Name FROM pcFTypes"
							set rs=connTemp.execute(query)
							do while not rs.eof %>
								<option value="<%=rs("pcFType_idtype")%>" <% if rs("pcFType_idtype")=intFType then%>selected<%end if%> ><%=rs("pcFType_name")%></option>
								<%rs.MoveNext
							Loop%>
						</select>
						</td>
					</tr>
					<tr>
						<td align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_i")%>
						</td>
						<td>
						<select name="Priority">
							<option value=""></option>
							<% query="SELECT pcPri_idPri,pcPri_name FROM pcPri_Priority"
							set rs=connTemp.execute(query)
							do while not rs.eof %>
								<option value="<%=rs("pcPri_idPri")%>" <% if rs("pcPri_idPri")=intPriority then%>selected<%end if%>><%=rs("pcPri_name")%></option>
								<%rs.MoveNext
							Loop%>
						</select> 
						</td>
					</tr>
					<tr>
						<td align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_l")%>
						</td>
						<td>
						<select name="FStatus">
							<option value=""></option>
							<% query="SELECT pcFStat_idStatus,cFStat_name FROM pcFStatus"
							set rs=connTemp.execute(query)
							do while not rs.eof
								%>
								<option value="<%=rs("pcFStat_idStatus")%>" <% if rs("pcFStat_idStatus")=intFStatus then%>selected<%end if%>><%=rs("pcFStat_name")%></option>
								<%rs.MoveNext
							Loop%>
						</select> 
						</td>
					</tr>
				<%else 'Not Admins%>
					<tr>
						<td align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_h")%>
						</td>
						<td>
						<% query="SELECT pcFType_name FROM pcFTypes WHERE pcFType_IDType=" & intFType
						set rs=connTemp.execute(query)
						if not rs.eof then %>
							<%=rs("pcFType_name")%>
						<%else%>
							&nbsp;
						<% end if %>
						</td>
  				</tr>
					<tr>
						<td align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_i")%>
						</td>
						<td>
						<% query="SELECT pcPri_name FROM pcPriority WHERE pcPri_IDPri=" & intPriority
						set rs=connTemp.execute(query)
						if not rs.eof then %>
							<%=rs("pcPri_name")%>
						<%end if%>
						</td>
 					</tr>
  				<tr>
						<td align="right">
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_l")%>
						</td>
						<td>
						<% query="SELECT pcFStat_name FROM pcFStatus WHERE pcFStat_IDStatus=" & intFStatus
						set rs=connTemp.execute(query)
						if not rs.eof then %>
							<%=rs("pcFStat_name")%>
						<%else%>
							&nbsp;
						<% end if %>
						<input type="hidden" name="FType" value="<%=intFType%>">
						<input type="hidden" name="FStatus" value="<%=intFStatus%>">
						<input type="hidden" name="Priority" value="<%=intPriority%>">
						</td>
					</tr>
				<%end if 'Check Admins & Users%>
				<tr>
					<td align="right">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_j")%>
					</td>
					<td>
						<input name="Description" type="text" value="<%if request("Description")<>"" then%><%=request("Description")%><%else%><%=strDesc%><%end if%>" size="25" maxlength="100"> 
					</td>
				</tr>
				<tr>
					<td align="right" valign="top">
						<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_k")%>
						<br><br>
						<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=Details','window2')"></td>
					<td>
						<textarea name="Details" cols="40" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=strDetails%><%end if%></textarea>
					</td>
				</tr>
				<%if AllowUpload="1" then%>
					<tr>
						<td nowrap valign="top">
						<p align="right"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_s")%></p>
						</td>
						<td valign="top">
							<%query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & intIdFeedback
							set rs=connTemp.execute(query)
							if rs.eof then%>
								<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_6")%>
							<%else
								ACount=0
								do while not rs.eof
									ACount=ACount+1 %>
									<input type="hidden" name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>">
									<input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">
									<%
									strFileName= rs("pcUpld_FileName")
									strFileName = mid(strFileName,instr(strFileName,"_")+1,len(strFileName))%>
									<%=strFileName%><br>
									<%rs.MoveNext
								loop%>
								<input type="hidden" name="ACount" value="<%=ACount%>">
							<%end if%>
							<script language="JavaScript"><!--
								function newWindow1(file,window) {
								catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
								if (catWindow.opener == null) catWindow.opener = self;
								}
							//--></script>
							<br>
							<%response.write dictLanguage.Item(Session("language")&"_viewFeedback_7")%><a href="#" onclick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=<%=intIdFeedback%>&ReLink=<%=Server.URLencode("usereditfeedback.asp?IDOrder=" & scpre+clng(LngIdOrder) & "&IDFeedback=" & intIdFeedback)%>','window2')"><%response.write dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>
							</td>
						</tr>
						<%end if%>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
							<td align="right"></td>
							<td>
								<input type="submit" name="Submit" value="Update" class="submit2" onclick="document.hForm.rewrite.value='0';">
								<input type="button" name="back" value="Back" onClick="location='userviewfeedback.asp?IDOrder=<%=scpre+clng(LngIdOrder)%>&IDFeedback=<%=intIdFeedback%>'">
								<%if session("IDOrder")>0 then%>
								<input type="button" name="go" value="Order Messages" onClick="location='userviewallposts.asp?IDOrder=<%=session("IDOrder")%>';"><%end if%>
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