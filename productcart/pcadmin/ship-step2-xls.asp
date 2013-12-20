<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Import 'Order Shipped' Information - Map fields" %>
<% section = "orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<%if ucase(right(session("importfile"),4))=".CSV" then
response.redirect "ship-step2.asp?append=" & request("append") & "&movecat=" & request("movecat")
end if%>
<!--#include file="../includes/settings.asp"-->
<!--#include file="ship-checkfields.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<% append=request("append")
if append<>"" then
	session("append")=append
else
	append=session("append")
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

sub displayerror(msg)%>
<!--#include file="pcv4_showMessage.asp"-->
<% end sub %>

<table class="pcCPcontent">
    <tr>  
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td  width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Upload data file</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2a.gif"></td>
        <td><font color="#000000"><strong>Map fields</strong></font></td>
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


		<% if PPD="1" then
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
	<% end if %>
	
    <div style="margin: 10px;">
    Use the drop-down menus below to map existing fields in your data file, located on the left side of the page under
    'From' to ProductCart database fields, which are located on the right side of the page under
    'To'.
    </div>
    <form method="post" action="ship-step3-xls.asp" class="pcForms">
        <table class="pcCPcontent">
            <tr>
                <th width="50%">From:</th>
                <th width="50%">To:</th>
            </tr>
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
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
                                <td style="border-bottom: 1px solid #ccc;"><%=FiName%><input type=hidden name="F<%=validfields%>" value="<%=FiName%>" ><input type=hidden name="P<%=validfields%>" value="<%=i%>" ></td>
                                <td style="border-bottom: 1px solid #ccc;">
                                <select size="1" name="T<%=validfields%>">
                                    <option value="   ">   </option>
                                    <option value="Order ID">Order ID</option>
                                    <option value="Ship">Ship</option>
                                    <option value="Send Mail">Send Mail</option>
                                    <option value="Ship Date">Ship Date</option>
                                    <option value="Method">Method</option>
                                    <option value="Tracking Number">Tracking Number</option>                            
                                    <%if request("T" & validfields)<>"" then%>
                                        <option value="<%=request("T" & validfields)%>" selected><%=request("T" & validfields)%></option>
                                    <% else
                                        FiName1=""
                                        FiName1=CheckField(FiName)
                                        if FiName1<>"" then%>
                                            <option value="<%=FiName1%>" selected><%=FiName1%></option>
                                        <% end if
                                    end if %>
                                </select>
                                </td>
                            </tr>
                        <% end if
                    next %>   
               <tr>
               	<td colspan="2">
                	<div style="margin: 10px;">
                    <input type="hidden" name="validfields" value="<%=validfields%>">         
                    <input type="submit" name="submit" value="Map Fields" class="submit2">&nbsp; 
                    <input type="reset" name="reset" value="Reset"> 
                    </div>
        		</td>
              </tr>                
          </table>         
    </form>
<!--#include file="AdminFooter.asp"-->