<!--INCLUDE JQUERY - START -->
<%
private const scJQuery="1"
%>
<link type="text/css" href="../includes/jquery/themes/redmond/jquery-ui-1.8.9.custom.css" rel="stylesheet" />
<%
user_agent = request.servervariables("HTTP_USER_AGENT")
if Instr(ucase(user_agent),"CHROME")>0 then
	%>
    <link type="text/css" rel="stylesheet" href="onepagecheckout.css" />
    <link type="text/css" rel="stylesheet" href="onepagecheckoutGoogle.css" />
<% else %>
	<link type="text/css" rel="stylesheet" href="onepagecheckout.css" />
<% end if %>
<script type="text/javascript" src="../includes/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../includes/jquery/jquery.bgiframe.pack.js"></script>
<script type="text/javascript" src="../includes/jquery/jquery-ui-1.8.9.custom.min.js"></script>
<script type="text/javascript" src="../includes/jquery/jquery.validate.min.OPC.js"></script>
<script type="text/javascript" src="../includes/jquery/jquery.form.js"></script>
<!--INCLUDE JQUERY - END -->