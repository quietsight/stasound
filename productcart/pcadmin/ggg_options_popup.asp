<% response.Buffer=true %>
<% PmAdmin=7%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/productcartinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include FILE="../includes/pcProductOptionsCode.asp"--> 
<%
dim f, query, conntemp, rstemp
Dim pcv_strOptionGroupDesc, pcv_intOptionGroupCount, pcv_strOptionGroupCount, pcv_strOptionGroupID, pcv_strOptionRequired
Dim xOptionsCnt, pcv_strNumberValidations
'--> open database connection
call openDB()
%>
<!--#INCLUDE FILE="../pc/viewPrdCode.asp"-->
<%
pidproduct=Request("idproduct")
poption=Request("option")
gIDEvent=request("IDEvent")
epID=request("epID")

if request("action")="upd" then
	'--> New Product Options
	pcv_intOptionGroupCount = getUserInput(request.Form("OptionGroupCount"),0)
	if IsNull(pcv_intOptionGroupCount) OR pcv_intOptionGroupCount="" then
		pcv_intOptionGroupCount = 0
	end if
	pcv_intOptionGroupCount = cint(pcv_intOptionGroupCount)
	
	xOptionGroupCount = 0
	pcv_strSelectedOptions = ""
	do until xOptionGroupCount = pcv_intOptionGroupCount	
		xOptionGroupCount = xOptionGroupCount + 1
		pcvstrTmpOptionGroup = request.Form("idOption"&xOptionGroupCount)
		if pcvstrTmpOptionGroup <> "" then			
			pcv_strSelectedOptions = pcv_strSelectedOptions & pcvstrTmpOptionGroup & chr(124)	
		end if	
	loop
	' trim the last pipe if there is one
	xStringLength = len(pcv_strSelectedOptions)
	if xStringLength>0 then
		pcv_strSelectedOptions = left(pcv_strSelectedOptions,(xStringLength-1))
	end if
	
	query ="Update pcEvProducts SET pcEP_OptionsArray='" & pcv_strSelectedOptions & "' WHERE pcEP_ID=" & epID & " AND pcEP_IDProduct=" & pidproduct & " AND pcEP_IDEvent=" & gIDEvent & ";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	msg="This product has been updated!"
	btn="2"

end if
%>

<html>
<head>
<title>Modify Product Options</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<%if msg<>"" then%>
	<table class="pcCPcontent">
	<tr> 
		<td align="center"> 
			<div class="pcCPmessage">
				<%=msg%>
				<br>
				<br>
				<br>
				<br>
			</div>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();">
		</td>
	</tr>
</table>
<%else%>
<form name="form1" method="post" action="ggg_options_popup.asp?action=upd" class="pcForms">
<table class="pcCPcontent">
		
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
		<th colspan="2">Available Product Option</th>
	</tr>
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	<tr>
	<tr>
		<td>
		<p>
		This product currently has the following options available. <br>
	    When you click 'Save', the <u>options on this page will replace ALL existing options selections</u> on this line item.
		</p>
		</td>
	</tr>
	<tr>
		<td>		
			<table width="98%" border="0" align="center" cellpadding="4" cellspacing="0">
				<tr class="normal">
				<td><strong>Available Option(s)</strong></td>
				</tr>
				<% 
				pidproduct=request("idproduct")
				pidProductOrdered=request("idProductOrdered") 
				%>
				
				<tr class="normal">
					<td>		
					<% 
					pcs_OptionsN
					%>			
					<input type="hidden" name="idproduct" value="<%=pidproduct%>">
					<input type="hidden" name="IDEvent" value="<%=gIDEvent%>">
					<input type="hidden" name="epID" value="<%=epID%>">
					</td>
				</tr>
			  	<tr>
					<td align="center">
					<input type="submit" name="Submit" value="Save">
					<input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();">
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
			</table>			     
		</td>
	</tr>
</table>
</form>
<%
end if
call closeDb()
%>
</div>
</body>
</html>