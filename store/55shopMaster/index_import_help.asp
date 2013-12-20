<% pageTitle = "Import Wizard - Instructions" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%on error resume next%>
<form class="pcForms">
<table class="pcCPcontent">
<tr>
	<td valign="top">
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>

		<div class="pcCPmessage" style="width: 500px;">IMPORTANT <span style="font-weight: normal;">- Carefully read the <a href="http://wiki.earlyimpact.com/productcart/products_import" target="_blank">Product Import Wizard documentation</a> before attempting to import or update product information.</span></div>
		<p style="padding-top:10px;" align="center"><input type="button" value="Proceed to Import Wizard" onClick="location.href='index_import.asp'" class="submit2">&nbsp;
		<%
		CSVFile = "importlogs/prologs.txt"
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
				<input type="button" value="Undo Last <%=InsTop%>" class="submit2" onClick="javascript:if (confirm('You are about to undo your last product <%=InsTop%>. All the information added to the database during the import/update will be removed. ProductCart saved a log of the information imported/updated in the files pcadmin/importlogs/prologs.txt and pcadmin/importlogs/catlogs.txt. You should NOT use this feature if you have further updated the product catalog after having imported/updated product data. Are you sure you want to complete this action?')) location='undoimport.asp'">&nbsp;
			<%
			end if
			f.close
			set f=nothing
		end if
		set fso=nothing%>
		<input type="button" value="Import Guide" onClick="window.open('http://wiki.earlyimpact.com/productcart/products_import')">
		</p>
	</td>
</tr>
<tr>
	<td>
		<p><input type="button" value="Export for Re-Import" onClick="location.href='ReverseImport_step1.asp'">&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=438')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>&nbsp;<input type="button" value="Import Product Additional Images" onClick="location.href='iistep1.asp'"></p>
	</td>
</tr>
</table>
</form>
<!--#include file="AdminFooter.asp"-->