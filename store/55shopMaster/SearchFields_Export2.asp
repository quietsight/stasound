<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Seach Fields - Confirm Mapping" %>
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
<% dim query, conntemp, rstemp

pcv_strExportType = request("export")
Select Case pcv_strExportType
	Case "f": pcv_strExportFile = "Google Shopping"
	Case "c": pcv_strExportFile = "Cashback"
End Select

call openDb()

validfields=request.form("validfields")

R1=0
mfilter=""
herror=false
mtemp=""

'/////////////////////////////////////////////////////////////////////
'// START: QUERY STRING
'/////////////////////////////////////////////////////////////////////
For i=1 to validfields
	if trim(ucase(request("T" & i)))<>"" then
		if instr(mfilter,"*" & ucase(request("T" & i)) & "*")>0 then
			herror=true
		else
			mfilter=mfilter & "*" & ucase(request("T" & i)) & "*"
		end if
	else
		R1=R1+1
	end if
	mtemp= mtemp & "&" & "T" & i & "=" & request("T" & i)
Next
'/////////////////////////////////////////////////////////////////////
'// END: QUERY STRING
'/////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////
'// START: VALIDATION
'/////////////////////////////////////////////////////////////////////
if R1=cint(validfields) then
	msg="Please make sure that at least one field is mapped."
	response.redirect "SearchFields_Export.asp?msg=" & msg & mtemp
end if
'if herror=true then
'	msg="Some of the mapping instructions are overlapping. Please make sure that the database fields are mapped uniquely."
'	response.redirect "SearchFields_Export.asp?msg=" & msg & mtemp
'end if
'/////////////////////////////////////////////////////////////////////
'// END: VALIDATION
'/////////////////////////////////////////////////////////////////////
%>
<!--#include file="AdminHeader.asp"-->
<form method="post" action="SearchFields_Export3.asp?export=<%=pcv_strExportType%>" class="pcForms">
	<table class="pcCPcontent">
		<tr>
			<td valign="top" colspan="2">
                <table class="pcCPcontent">
                    <tr>
                        <td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
                        <td width="95%"><font color="#A8A8A8">Map Fields</font></td>
                    </tr>
                    <tr>
                        <td align="center"><img border="0" src="images/step2a.gif"></td>
                        <td><strong>Confirm Mapping</strong></td>
                    </tr>
                    <tr>
                        <td align="center"><img border="0" src="images/step3.gif"></td>
                        <td><font color="#A8A8A8">Return to Export Wizard</font></td>
                    </tr>
                </table>
            	<div class="pcCPnotes">
       				Please make sure that the database fields are mapped correctly. If not, click &quot;Back to Step One&quot; to try again.
              </div>
             </td>
        </tr>
        <tr>
          <td class="pcCPspacer"></td>
        </tr>	
        <tr>
          <th><b>Export Field:</b></th>
          <th><b>Custom Search Field:</b></th>
        </tr>
        <tr>
          <td class="pcCPspacer"></td>
        </tr>	
        <% 
		validfields=0
        For i=1 to request("validfields")
            if trim(request("T" & i))<>"" then
            validfields=validfields+1
			%>
            <tr>
                <td width="20%">
                    <strong><%=request("F" & i)%></strong>
                    <input type=hidden name="P<%=validfields%>" value="<%=request("P" & i)%>">
                    <input type=hidden name="F<%=validfields%>" value="<%=request("F" & i)%>">                
                </td>
                <td width="80%">
                	<%
					pcv_intSearchField = request("T" & i)
                    query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE idSearchField="& pcv_intSearchField &" ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
                    set rs=Server.CreateObject("ADODB.Recordset")
                    set rs=conlayout.execute(query)
                    if not rs.eof then
						pcv_intSearchFieldName = rs("pcSearchFieldName")
					else
						pcv_intSearchFieldName = ""					
					end if
					%>
                    <%=pcv_intSearchFieldName%>
                    <input type=hidden name="T<%=validfields%>" value="<%=pcv_intSearchField%>">                
                </td>
            </tr>
            <%
			end if
        Next
		%>    
        <tr>
          <td colspan="2" class="pcCPspacer"><hr></td>
        </tr>	               
        <tr>
            <td colspan="2">
                <input type="hidden" name="validfields" value=<%=validfields%> >         
                <input type="button" name="backstep" value="<< Back to Step One" onClick="location='SearchFields_Export.asp?a=1<%=mtemp%>&export=<%=pcv_strExportType%>';" class="ibtnGrey">&nbsp; 
                <input type="submit" name="submit" value="Go to Step Three >>" class="submit2">            
        	</td>
        </tr>
    </table>
</form>
<% call closeDb() %>
<!--#include file="AdminFooter.asp"-->