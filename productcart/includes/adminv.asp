<% response.Buffer=true 
'// Verifies if admin is logged, so as not send to login page
if session("admin")="0" OR session("admin")="" then
	response.Write("You do not have proper rights to access this page.")
	response.End()
end if
%>