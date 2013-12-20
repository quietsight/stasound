<%
if scVersion<>"4.7" AND scVersion<>"4.7b" then
	updtrigger=1
	updDBScript="upddb_v47.asp"
	updSubVersion=""
else
	updtrigger=0
	updDBScript=""
end if
%>