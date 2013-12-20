<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %> 
<% Response.Expires = -1 %>
<% response.Buffer=true 
Server.ScriptTimeout = 120 %>
<% PmAdmin="1*6*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/taxSettings.asp"--> 
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<%
dim conntemp, rs, query
%>
<html>
<head>
    <title>Add/Edit Fallback State Tax Rate</title>
</head>
<body>
<center>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
    <tr>
        <td bgcolor="#6666cc"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Add/Edit Fallback State Tax Rate</font></strong></td>
    </tr>
    <% if request("ro")="1" then
        strRStateCode=request("rstate")
        call openDB()
        query="SELECT stateName FROM states WHERE stateCode='"&strRStateCode&"';"
        set rs=Server.CreateObject("ADODB.RecordSet")
        set rs=conntemp.execute(query)
        strRStateName=rs("stateName")
        set rs=nothing
        call closeDB()
        %>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td><p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">You have successfully created a Fallback Tax Rate for &nbsp;<%=strRStateName%>&nbsp;</font></p>
            </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
            <td>
                <p align="center"><input type="button" name="Back" value="Close Window" onClick="opener.location.reload(); self.close();"></td>
        </tr>
    <% else %>
        <% call openDB()
		query="SELECT stateCode, stateName FROM states ORDER BY stateName;"
		set rs=Server.CreateObject("ADODB.RecordSet")
		set rs=conntemp.execute(query)
		%>
		<tr>
			<td>
				<form action="../includes/PageCreateTaxSettings.asp" method="post" name="form1" id="form1">
					<table width="100%" border="0" cellspacing="0" cellpadding="4">
						<tr> 
							<td width="17%"><font size="2" face="Arial, Helvetica, sans-serif">State:</font></td>
							<td width="83%"> <select name="taxRateState">
							<% do until rs.eof
								strStateCode=rs("stateCode")
								strStateName=rs("stateName")
								varStatePre=0
								stateArray=split(ptaxRateState,", ")
								for i=0 to ubound(stateArray)
									if strStateCode=stateArray(i) then
										varStatePre=1
									end if
								next 
								if varStatePre=0 then %>
									<option value="<%=strStateCode%>"><%=strStateName%></option>
								<% end if 
								rs.moveNext
							loop 
							set rs=nothing
							call closedb()
							%>
							</select>
							<input type="hidden" name="RateOnly" value="1">
                            <input type="hidden" name="Page_Name" value="taxsettings.asp"> 
							<input type="hidden" name="refpage" value="AddTaxRateState_popup.asp">
							</td>
						</tr>
						<tr> 
							<td nowrap><font size="2" face="Arial, Helvetica, sans-serif">Fallback Tax Rate:</font></td>
							<td nowrap> <input name="taxRateDefault" type="text" size="4"> % </td>
						</tr>
						<tr> 
						  <td><div align="right">
                                <input type="hidden" name="PopForm" value="YES">
                                <input type="radio" name="taxSNH" value="YN">
							</div></td>
						  <td><font size="2" face="Arial, Helvetica, sans-serif">Tax shipping</font></td>
						</tr>
						<tr> 
						  <td><div align="right"> 
							  <input type="radio" name="taxSNH" value="YY">
							</div></td>
						  <td><font size="2" face="Arial, Helvetica, sans-serif">Tax shipping and handling together</font></td>
						</tr>
						<tr> 
						  <td><div align="right"> 
							  <input name="taxSNH" type="radio" value="NN" checked>
							</div></td>
						  <td><font size="2" face="Arial, Helvetica, sans-serif">Do not tax shipping or handling</font></td>
						</tr>
						<tr> 
						  <td>&nbsp;</td>
						  <td>&nbsp;</td>
						</tr>
						<tr> 
						  <td>&nbsp;</td>
						  <td><input type="submit" name="Add" value="Add"></td>
						</tr>
					</table>
				</form>
			</td>
		</tr>
	<% end if %>
</table>
</center>
</body>
</html>
