<% 
'// Check for same-page message
IF msg<>"" THEN
	pcStrMsg=trim(msg)
	pcvMessageType=msgType
'// Check for querystrings and forms
ELSE
	pcStrMsg=trim(request.querystring("msg"))
	if pcStrMsg="" then
		pcStrMsg=trim(request.querystring("message"))
	end if
	pcvMessageType=request.querystring("s")
END IF
	
if pcStrMsg<>"" then
	if not validNum(pcvMessageType) then pcvMessageType=0
	if pcvMessageType=1 then %>
	<div class="pcCPmessageSuccess"><%=pcStrMsg%></div>
<% 
	else 
%>
	<div class="pcCPmessage"><%=pcStrMsg%></div>
<% 
	end if 
end if
%>