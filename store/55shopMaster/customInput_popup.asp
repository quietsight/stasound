<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/stringfunctions.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<%
'on error resume next
dim f, mySQL, conntemp, rstemp
call openDB()
%>
<html>
<head>
<title>Edit Order Custom Input Fields</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="pcCPmain" style="width:450px; background-image: none;">
<% 
if request("action")="update" then
	pidProductOrdered=request("idProductOrdered")
	xCnt=request.form("xCnt")
	pxfield1=getUserInput(request.form("x0"),500)
	pxfield2=getUserInput(request.form("x1"),500)
	pxfield3=getUserInput(request.form("x2"),500)
	'// Repair HTML for the Line Return
	pxfield1=replace(pxfield1,"&lt;BR&gt;","<BR>")
	pxfield2=replace(pxfield2,"&lt;BR&gt;","<BR>")
	pxfield3=replace(pxfield3,"&lt;BR&gt;","<BR>")
	xCnt=0
	tempArray=""
	if pxfield1<>"" then
		pxfield1=replace(pxfield1,vbCrlf,"<BR>")
		pXfieldDescrip1=request.form("xDesc0")
		tempArray=tempArray&pXfieldDescrip1&": "&pxfield1
		xCnt=1
	end if
	if pxfield2<>"" then
		pxfield2=replace(pxfield2,vbCrlf,"<BR>")
		pXfieldDescrip2=request.form("xDesc1")
		if xCnt=1 then
			tempArray=tempArray& "|"
		end if
		xCnt=1
		tempArray=tempArray&pXfieldDescrip2&": "&pxfield2
	end if
	if pxfield3<>"" then
		pxfield3=replace(pxfield3,vbCrlf,"<BR>")
		pXfieldDescrip3=request.form("xDesc2")
		if xCnt=1 then
			tempArray=tempArray& "|"
		end if
		xCnt=1
		tempArray=tempArray&pXfieldDescrip3&": "&pxfield3
	end if
		
	query="UPDATE productsOrdered SET xfdetails='"&tempArray&"' WHERE idProductOrdered="&pidProductOrdered&";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
%>
	<table class="pcCPcontent">
		<tr>
			<th>Edit Custom Input Field</th>
		</tr>
		<tr>
			<td><br>
				<br>
				<b>This order has been updated!<br>
				<br>
			</b><br>
				</td>
		</tr>
		<tr>
			<td>
				<p align="center"><input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();"></td>
		</tr>
	</table>
<%ELSE%>
<form name="form1" method="post" action="customInput_popup.asp?action=update" class="pcForms">
<table class="pcCPcontent">
  <tr>
    <th>Edit Custom Input Field</th>
  </tr>
  <tr>
    <td>Edit the current field and click on the &quot;update&quot; button.</td>
  </tr>
  <tr>
    <td><table width="98%" border="0" align="center" cellpadding="4" cellspacing="0">
      <% c=request("c")
			pidProductOrdered=request("idProductOrdered")
			query="select xfdetails From ProductsOrdered WHERE idProductOrdered="&pidProductOrdered&";"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=conntemp.execute(query)
			pxfDetails=rs("xfdetails") 
			
			if pxfDetails<>"" then 
				xfArray=split(pxfdetails,"|")
				for xf=0 to ubound(xfArray)
					tempXf=xfArray(xf)
					xSplitArray=split(tempXf,": ") 
					if int(c)=int(xf) then
						%>
						<tr valign="top">
						  <td><%=xSplitArray(0)%>:</td>
	  </tr>
						<tr valign="top">
							<td><input type="hidden" name="xdesc<%=c%>" value="<%=xSplitArray(0)%>"><textarea name="x<%=c%>" cols="30" rows="5"><%=replace(xSplitArray(1),"<BR>",vbCrlf)%></textarea>
							</td>
						</tr>
					<% else %>
						<input type="hidden" name="xdesc<%=xf%>" value="<%=xSplitArray(0)%>">
						<input type="hidden" name="x<%=xf%>" value="<%=xSplitArray(1)%>">
					<% end if 
				next %>
				<input type="hidden" name="xCnt" value="<%=ubound(xfArray)%>">
			<% end if %>
			<input type="hidden" name="idProductOrdered" value="<%=pidProductOrdered%>">
      <tr>
        <td><input type="submit" name="Submit" value="Update">
        <input type="button" name="Back" value="Close" onClick="self.close();">
        </td>
        </tr>
      <tr>
        <td align="center">&nbsp;</td>
      </tr>
    </table>      
    </td>
  </tr>
</table>
</form>
<%END IF%>
</div>
</body>
</html>
