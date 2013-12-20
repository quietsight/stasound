<% pageTitle = "Import 'Order Shipped' Information - Instructions" %>
<% section = "orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
on error resume next
%>
<form class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>

		<div class="pcCPmessage" style="width: 500px;">IMPORTANT <span style="font-weight: normal;">- Carefully read the <a href="http://wiki.earlyimpact.com/productcart/orders_importing_shipping_info" target="_blank">Product Import Wizard documentation</a> before attempting to import or update product information.</span></div>
		<div style="padding-top:10px;"><input type="button" value="Proceed to Import Wizard" class="submit2" onClick="location.href='ship-index_import.asp'">&nbsp;
		<%
            CSVFile = "importlogs/ship-prologs.txt"
            findit = Server.MapPath(CSVfile)
            Set fso = server.CreateObject("Scripting.FileSystemObject")
            Err.number=0
            MyTest=1
            Set f = fso.OpenTextFile(findit, 1)
            if Err.number>0 then
            MyTest=0
            Err.number=0
            Err.Description=0
            end if
            if MyTest=1 then
            Topline = f.Readline
            InsTop=""
            if TopLine="IMPORT" then
            InsTop="Import"
            end if
            if TopLine="UPDATE" then
            InsTop="Update"
            end if
            if InsTop<>"" then
        %>
            <input type="button" value="Undo Last <%=InsTop%>" onClick="javascript:if (confirm('You are about to undo your last data <%=InsTop%>. All the information added to the database during the import/update will be removed. ProductCart saved a log of the information imported/updated in the files pcadmin/importlogs/ship-prologs.txt. You should NOT use this feature if you have further updated the orders after having imported/updated shipped data. Are you sure you want to complete this action?')) location='ship-undoimport.asp'">&nbsp;
		<%
        end if
        f.close
        set f=nothing
        end if
        set fso=nothing
        %>
        </div>
		</td>
	</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->