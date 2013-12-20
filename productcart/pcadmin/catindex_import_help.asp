<% pageTitle = "Category Import Wizard - Instructions" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
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
<table class="pcCPcontent">
<tr>
	<td>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
        
		<div class="pcCPmessage" style="width: 500px;">IMPORTANT <span style="font-weight: normal;">- Carefully read the <a href="http://www.earlyimpact.com/support/userguides/importCategories.asp" target="_blank">Category Import Wizard documentation</a> before attempting to import or update category information.</span></div>
		<p align="center">
		<input type="button" value="Proceed to Category Import Wizard" class="submit2" onClick="location.href='catindex_import.asp'">&nbsp;
		<%
		CSVFile = "importlogs/categorylogs.txt"
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
				<input type="button" value="Undo Last <%=InsTop%>" class="submit2" onClick="javascript:if (confirm('You are about to undo your last Category <%=InsTop%>. All the information added to the database during the import/update will be removed. ProductCart saved a log of the information imported/updated in the file pcadmin/importlogs/categorylogs.txt. You should NOT use this feature if you have further updated the Category information after having imported/updated Category data. Are you sure you want to complete this action?')) location='undocatimport.asp'">&nbsp;
			<%
			end if
			f.close
			set f=nothing
		end if
		set fso=nothing%>
	<input type="button" value="Help" class="ibtnGrey" onClick="window.open('http://www.earlyimpact.com/support/userguides/importCategories.asp')">
	</p>
	</td>
</tr>
<tr>
	<td>
		<p><input type="button" value="Export for Re-Import" onClick="location.href='ReverseCatImport_step1.asp'" class="ibtnGrey"></p>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->