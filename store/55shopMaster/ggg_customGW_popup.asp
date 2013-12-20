<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<%
'on error resume next
dim f, query, conntemp, rstemp
call openDB()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Edit Order Custom Input Fields</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body style="background-image: none;">
<center>
<%IF request("action")<>"" then 
	if request("action")="update" then
		pidProductOrdered=request("idProductOrdered")
		pcv_GWOpt=request("IDGW")
		if (pcv_GWOpt="") then
			pcv_GWOpt="0"
		else
			pcv_GWOpt=replace(pcv_GWOpt,"'","''")
		end if
		if pcv_GWOpt="0" then
			pcv_GWNote=""
			pcv_GWPrice="0"
		else
			pcv_GWNote=request("GWNote")
			if pcv_GWNote<>"" then
				pcv_GWNote=replace(pcv_GWNote,"'","''")
			end if
			pcv_GWPrice=request("GWPrice")
			if pcv_GWPrice<>"" then
				pcv_GWPrice=replace(pcv_GWPrice,"'","''")
			end if
		end if
		
		query="UPDATE productsOrdered SET pcPO_GWOpt=" & pcv_GWOpt & ",pcPO_GWNote='" & pcv_GWNote & "',pcPO_GWPrice=" & pcv_GWPrice & " WHERE idProductOrdered="&pidProductOrdered&";"
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		set rs=nothing
	end if
	
	IF request("action")="del" then
		pidProductOrdered=request("idProductOrdered")
		IDGW=request("IDGW")
		query="UPDATE productsOrdered SET pcPO_GWOpt=0,pcPO_GWNote='',pcPO_GWPrice=0 where idProductOrdered="&pidProductOrdered&";"
		set rs=conntemp.execute(query)
		set rs=nothing
	END IF
	%>
	
	<table class="pcCPcontent">
	<tr>
		<th>Edit Gift Wrapping Option</th>
	</tr>
	<tr>
		<td>
			<br>
			<br>
			<div class="pcCPmessage"><%if request("action")="del" then%>This gift wrapping option was deleted successfully!<%else%>This gift wrapping option has been updated!<%end if%></div>
			<br>
			<br>
			<br>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" name="Back" value="Close Window" onClick="opener.location='AdminEditOrder.asp?ido=<%=request("idOrder")%>&action=upd'; self.close();" class="iBtnGrey">
		</td>
	</tr>
	</table>
<%ELSE%>
	<form name="form1" method="post" action="ggg_customGW_popup.asp?action=update" class="pcForms">
	<input type="hidden" name="idProductOrdered" value="<%=request("idProductOrdered")%>">					  
	<input type="hidden" name="idOrder" value="<%=request("idOrder")%>">
	<table class="pcCPcontent">
	<tr>
		<th colspan="2">Edit Gift Wrapping Option</th></tr>
	<tr>
		<td colspan="2">Edit the gift wrapping option and click on the <em>Update</em> button.</td>
	</tr>
	<%pidProductOrdered=request("idProductOrdered")
	query="select ProductsOrdered.pcPO_GWOpt,ProductsOrdered.pcPO_GWNote,ProductsOrdered.pcPO_GWPrice from ProductsOrdered WHERE idProductOrdered="&pidProductOrdered&";"
	set rsGW=connTemp.execute(query)
	pcv_GWOpt=rsGW("pcPO_GWOpt")
	pcv_GWNote=rsGW("pcPO_GWNote")
	pcv_GWPrice=rsGW("pcPO_GWPrice")
	set rsGW=nothing

	tmpPrice=""
	query="select pcGW_IDOpt,pcGW_OptName,pcGW_OptPrice from pcGWOptions order by pcGW_OptName asc"
	set rsGW=connTemp.execute(query)%>
	<tr>
		<td nowrap>Option Name:</td>
		<td>
			<select name="IDGW" onChange="javascript:UpdPrice(this.value);">
		    <option value="0" selected>Select Gift wrapping option</option>
			<%do while not rsGW.eof%>
				<option value="<%=rsGW("pcGW_IDOpt")%>" <%if clng(rsGW("pcGW_IDOpt"))=clng(pcv_GWOpt) then%>selected<%end if%>><%=rsGW("pcGW_OptName")%></option>
				<%
				tmpPrice=tmpPrice & "if (IDOpt == " & rsGW("pcGW_IDOpt") & ") { tprice=" & rsGW("pcGW_OptPrice") & "; }" & vbcrlf
				rsGW.MoveNext
			loop
			set rsGW=nothing
			%>
			</select>
			<script>
			function UpdPrice(IDOpt)
			{
				var tprice=0;
				<%if tmpPrice<>"" then%>
					<%=tmpPrice%>
				<%end if%>
				document.form1.GWPrice.value=tprice;
			}
			</script>
		</td>
	</tr>
	<tr>
		<td valign="top">Gift Note:</td>
		<td>
			<textarea name="GWNote" rows="6" cols="25"><%=pcv_GWNote%></textarea>
	    </td>
	</tr>
	<tr>
		<td>Price:</td>
		<td>
			<input type="text" name="GWPrice" value="<%=pcv_GWPrice%>" size="30">
	    </td>
	</tr>
	<tr>
		<td align="center" colspan="2">
			<input type="submit" name="Submit" value="Update" class="submit2">&nbsp;
			<input type="button" name="Back" value="Close" onClick="self.close();" class="iBtnGrey">
	    </td>
	</tr>
	</table>
	</form>
<%END IF%>
</center>
</body>
</html>