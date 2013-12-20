<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Seach Fields - Map Fields" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/shipFromSettings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="AdminHeader.asp"-->
<% Dim rstemp, connTemp, query, rs, rsQ

pcv_strExportType = request("export")
	
	if trim(pcv_strExportType)="" then
%>
    <table class="pcCPcontent">
      <tr>
        <td>Choose a data feed type to create the mappings:
          <ul class="pcListIcon">
            <li><a href="SearchFields_Export.asp?export=f">Google Shopping</a></li>
            <li><a href="SearchFields_Export.asp?export=c">Microsoft Bing Shopping</a></li>
          </ul>
          <br />
          <br />
        </td>
       </tr>
    </table>
<%
   else

			Select Case pcv_strExportType
				Case "f": pcv_strExportFile = "Google Shopping"
				Case "c": pcv_strExportFile = "Cashback"
			End Select
			
			call openDb()
			
			sub displayerror(msg)
				if msg<>"" then %>
					<div class="pcCPmessage"> 
						<img src="images/pcadmin_note.gif" width="20" height="20"> 
						<%=msg%>
					</div>
				<%end if 
			end sub
			%>
			<form method="post" action="SearchFields_Export2.asp?export=<%=pcv_strExportType%>" class="pcForms">
					<table class="pcCPcontent">
							<tr>
									<td valign="top" colspan="2">
											<table class="pcCPcontent">
													<tr>
															<td width="5%" align="center"><img border="0" src="images/step1a.gif"></td>
															<td width="95%"><strong>Map Fields</strong></td>
													</tr>
													<tr>
															<td align="center"><img border="0" src="images/step2.gif"></td>
															<td><font color="#A8A8A8">Confirm Mapping</font></td>
													</tr>
													<tr>
															<td align="center"><img border="0" src="images/step3.gif"></td>
															<td><font color="#A8A8A8">Return to Export Wizard</font></td>
													</tr>
											</table>
							
									<!--#include file="../includes/ppdstatus.inc"-->
									<%
									FileCSV = "CSF_"& pcv_strExportFile &"Fields.txt"
									findit = Server.MapPath(FileCSV)
									Set fso = server.CreateObject("Scripting.FileSystemObject")	
									Err.number=0
									Set f = fso.OpenTextFile(findit, 1)
									if Err.number>0 then
											session("importfilename")=""%>
											<script>
											location="msg.asp?message=31";
											</script><%
									end if
									
									Topline = f.Readline
									a=split(Topline,",")
			
									validfields=0
									for i=lbound(a) to ubound(a)
											if trim(a(i))<>"" then
													validfields=validfields+1
											end if
									next
			
									session("totalfields")=ubound(a)-lbound(a)+1
									if a(ubound(a))="" then
											session("totalfields")=session("totalfields")-1
									end if
									f.Close
									Set fso = nothing
									Set f = nothing
									
									msg=request.querystring("msg")
									if msg<>"" then 
											displayerror(msg)
											response.Write("<br />")
									end if 
									%>
									<div class="pcCPnotes">
											Use the drop-down menus below to map the export fields, located on the left side of the page under
											&quot;Export Field&quot; to custom search fields, which are located on the right side of the page under &quot;Custom Search Field&quot;.
									</div>
									</td>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<tr>
									<th width="20%"><b>Export Field:</b></th>
									<th width="80%"><b>Custom Search Field:</b></th>
							</tr>
							<tr>
									<td colspan="2" class="pcCPspacer"></td>
							</tr>
							<% 
							validfields=0
							For i=lbound(a) to ubound(a)
									if trim(a(i))<>"" then
											if left(a(i),1)=chr(34) then
													a(i)=mid(a(i),2,len(a(i)))
											end if
											if right(a(i),1)=chr(34) then
													a(i)=mid(a(i),1,len(a(i))-1)
											end if    	
											validfields=validfields+1
											%>
											<tr>
													<td>
															<strong><%=a(i)%></strong>
															<input type=hidden name="F<%=validfields%>" value="<%=a(i)%>" >
															<input type=hidden name="P<%=validfields%>" value="<%=i%>" >
													</td>
													<td>
													<%
								FieldName=""
								query="SELECT pcSearchFields_Mappings.idSearchField "
								query = query & "FROM pcSearchFields INNER JOIN pcSearchFields_Mappings "
								query = query & "ON pcSearchFields.idSearchField  = pcSearchFields_Mappings.idSearchField "
								query=query&"WHERE pcSearchFields_Mappings.pcSearchFieldsColumn='" & a(i) & "' "
								query=query&"AND pcSearchFields_Mappings.pcSearchFieldsFileID='" & pcv_strExportType & "';"
								set rs=server.CreateObject("ADODB.RecordSet")
								set rs=connTemp.execute(query)
								if not rs.eof then
									FieldName=rs("idSearchField")
								end if
								set rs=nothing
								%>
													<select size="1" name="T<%=validfields%>">
															<option value="0"></option>
															<%
															query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
															set rs=Server.CreateObject("ADODB.Recordset")
															set rs=conlayout.execute(query)
															if not rs.eof then
																	pcArray=rs.getRows()
																	intCount=ubound(pcArray,2)
																	set rs=nothing	
																	pcv_ClearAll = 0
																	For pcv_intSFCount=0 to intCount
																			pcv_strCatSearchName = pcArray(1,pcv_intSFCount)  
																			pcv_strCatSearchID = pcArray(0,pcv_intSFCount)
																			
																			if (request("T" & validfields)<>"") AND ( cint(request("T" & validfields))=cint(pcv_strCatSearchID) ) then
																					%>
																					<option value="<%=pcv_strCatSearchID%>" selected><%=pcv_strCatSearchName%></option>
																					<%
																			else
			
												
																					if FieldName=pcv_strCatSearchID then
																					%>
																							<option value="<%=pcv_strCatSearchID%>" selected><%=pcv_strCatSearchName%></option>
																					<% else %>
																							<option value="<%=pcv_strCatSearchID%>"><%=pcv_strCatSearchName%></option>
																					<%
																					end if
																			end if
																	Next '// For i=0 to intCount
															end if
															%>
													</select>
													</td>
											</tr>
											<%
									end if
							Next
							%>  
							<tr>
									<td class="pcCPspacer" colspan="2"><hr></td>
							</tr>                 
							<tr>
									<td colspan="2">
											<input type="hidden" name="validfields" value=<%=validfields%> >         
											<input type="submit" name="submit" value="Map Fields" class="submit2">
                                            &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
									</td>
							</tr>
					</table>
			</form>
			<% call closeDb()
			
			end if
			%>
<!--#include file="AdminFooter.asp"-->