<% pageTitle = "Customer Import Wizard - Map fields" %>
<% section = "mngAcc" %>
<%PmAdmin=7%><!--#include file="adminv.asp"--> 
<%
if ucase(right(session("importfile"),4))=".CSV" then
response.redirect "cwstep2.asp?append=" & request("append") & "&movecat=" & request("movecat")
end if%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="cwcheckfields.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
on error resume next
Dim connTemp,query,rs

append=request("append")
if append<>"" then
session("append")=append
end if
movecat=request("movecat")
if movecat<>"" then
else
movecat="1"
end if
session("movecat")=movecat
if append="1" then
requiredfields = 1
else
requiredfields = 4
end if

sub displayerror(msg)
%>
<!--#include file="pcv4_showMessage.asp"-->
<%
end sub%>

    <table class="pcCPcontent">
    <tr>
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Select product data file</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2a.gif"></td>
        <td><strong>Map fields</strong></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4.gif"></td>
        <td><font color="#A8A8A8"><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</font></td>
    </tr>
    </table>
		<br>
	<% 
	if PPD="1" then
		FileXLS = "/"&scPcFolder&"/pc/catalog/" & session("importfile")
	else
		FileXLS = "../pc/catalog/" & session("importfile")
	end if
	
	Set cnnExcel = Server.CreateObject("ADODB.Connection")
	cnnExcel.Open "DRIVER={Microsoft Excel Driver (*.xls)};" & " DBQ=" & Server.MapPath(FileXLS) & ";"
	Set rsExcel = Server.CreateObject("ADODB.Recordset")
	rsExcel.open "SELECT * FROM IMPORT;", cnnExcel 
	if Err.number<>0 then
		session("importfilename")=""%>
		<script>
		location="msg.asp?message=30";
		</script><%
	else
	iCols = rsExcel.Fields.Count
	if iCols<requiredfields then
	 session("importfilename")=""%>
	 <script>
	 location="msg.asp?message=29";
	 </script><%
	end if
	end if
	validfields=0
   for i=0 to iCols-1
    if trim(rsExcel.Fields.Item(I).Name)<>"" then
     validfields=validfields+1
    end if
   next
  	if validfields<requiredfields then
   	 session("importfilename")=""%>
	 <script>
	 location="msg.asp?message=29";
	 </script><%
	end if
	session("totalfields")=iCols
	msg=request.querystring("msg")
	if msg<>"" then 
		displayerror(msg)%>
		<br>
	<%end if %>

	<div style="padding: 7px;">Use the drop-down menus below to map existing fields in your customer database, located on the left side of the page under 'From' to ProductCart database fields, which are located on the right side of the page under 'To'.</div>
	<form method="post" action="cwstep3-xls.asp" class="pcForms">
	<table class="pcCPcontent">
    <tr>
    	<td colspan="2"></td>
    </tr>
	<tr>
		<th width="50%">From:</th>
		<th width="50%">To:</th>
	</tr>
    <tr>
    	<td colspan="2"></td>
    </tr>
	<% validfields=0
	for i=0 to iCols-1
	FiName=rsExcel.Fields.Item(I).Name
	if trim(FiName)<>"" then
		if left(FiName,1)=chr(34) then
			FiName=mid(FiName,2,len(FiName))
		end if
		if right(FiName,1)=chr(34) then
			FiName=mid(FiName,1,len(FiName)-1)
		end if
		validfields=validfields+1%>
		<tr>
			<td width="50%" style="border-bottom: 1 solid #666666"><%=FiName%><input type=hidden name="F<%=validfields%>" value="<%=FiName%>" ><input type=hidden name="P<%=validfields%>" value="<%=i%>" ></td>
			<td width="50%" style="border-bottom: 1 solid #666666">
				<select size="1" name="T<%=validfields%>">
					<option value="   ">   </option>
					<option value="E-mail Address">E-mail Address</option>
					<option value="Password">Password</option>
					<option value="Customer Type">Customer Type</option>
					<option value="First Name">First Name</option>
					<option value="Last Name">Last Name</option>
					<option value="Company">Company</option>
					<option value="Phone">Phone</option>
					<option value="Fax">Fax</option>
					<option value="Address">Address</option>
					<option value="Address 2">Address 2</option>
					<option value="City">City</option>
					<option value="State Code (US/Canada)">State Code (US/Canada)</option>
					<option value="Province">Province</option>
					<option value="Postal Code">Postal Code</option>
					<option value="Country Code">Country Code</option>
					<option value="Shipping Company">Shipping Company</option>
					<option value="Shipping Address">Shipping Address</option>
					<option value="Shipping Address 2">Shipping Address 2</option>
					<option value="Shipping City">Shipping City</option>
					<option value="Shipping State Code (US/Canada)">Shipping State Code (US/Canada)</option>
					<option value="Shipping Province">Shipping Province</option>
					<option value="Shipping Postal Code">Shipping Postal Code</option>
					<option value="Shipping Country Code">Shipping Country Code</option>
					<option value="Current Reward Points Balance">Current Reward Points Balance</option>
					<option value="Pricing Category ID">Pricing Category ID</option>
					<option value="Shipping Email Address">Shipping Email Address</option>
					<option value="Shipping Phone">Shipping Phone</option>
					
					<%'Start Special Customer Fields
					call opendb()
					query="SELECT pcCField_ID,pcCField_Name,-1,'' FROM pcCustomerFields order by pcCField_ID asc;"
					set rs=connTemp.execute(query)
					
					session("cp_cw_custfields")=""
					session("cp_cw_HaveCustField")=""
					
					if not rs.eof then
						session("cp_cw_custfields")=rs.getRows()
						pcArr=session("cp_cw_custfields")
						set rs=nothing
						
						For k=0 to ubound(pcArr,2)%>
						<option value="<%=pcArr(1,k)%>"><%=pcArr(1,k)%></option>						
						<%Next
					end if
					
					set rs=nothing
					call closedb()
					'End of Special Customer Fields%>
					
					<option value="Newsletter Subscription">Newsletter Subscription</option>                            
					<%'MailUp-S%>
					<option value="Opt-in MailUp List IDs">Opt-in MailUp List IDs</option>
					<option value="Opt-out MailUp List IDs">Opt-out MailUp List IDs</option>
					<%'MailUp-E%>                            
					<%if request("T" & validfields)<>"" then%>
						<option value="<%=request("T" & validfields)%>" selected><%=request("T" & validfields)%></option>
					<%else
						FiName1=""
						FiName1=CheckField(FiName)
						if FiName1<>"" then%>
							<option value="<%=FiName1%>" selected><%=FiName1%></option>
						<%end if
					end if%>
				</select>
			</td>
		</tr>
    <%end if
	Next%>  
    <tr>
    	<td colspan="2"><hr></td>
    </tr>                 
    <tr>
		<td colspan="2">      
			<input type="hidden" name="validfields" value="<%=validfields%>">         
			<input type="submit" name="submit" value="Map Fields" class="submit2">
            &nbsp;<input type="reset" name="reset" value="Reset">
		</td>
	</tr>
	</table>
	</form>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->