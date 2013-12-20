<%
'////////////////////////////////////////////////////////////////////
'// START - Check for WWW and redirect if absent
'////////////////////////////////////////////////////////////////////
Dim strDomain, strPath, strQueryString, strURL, strHttpsDomain, intDoRedirect, intRedirectType

'// Redirect to maintain consistent URL
intDoRedirect = 0 ' 0 = NO; 1 = YES

'// Type of redirect
'// 1 = from without "www" to with "www" (e.g. mystore.com -> www.mystore.com)
'// 2 = from with "www" to without "www" (e.g. www.mystore.com -> mystore.com)
intRedirectType = 1

IF intDoRedirect = 1 THEN
	strDomain = Request.ServerVariables("HTTP_HOST")
	strPath = Request.ServerVariables("URL")
	strQueryString = Request.ServerVariables("QUERY_STRING")
	strHttpsDomain = ucase(Request.ServerVariables("HTTPS"))
	'// Clean up and concatenate
	if trim(strQueryString)<>"" then
		strQueryString = "?" & strQueryString
	end if
	if strHttpsDomain="ON" then
		strURL = "https://" & strDomain & strPath & strQueryString
	else
		strURL = "http://" & strDomain & strPath & strQueryString
	end if
	If len(strURL)>0 Then
		If instr(strURL,"404;")>0 Then
			strURL = "/" & Right(strURL,Len(strURL)-instr(strURL,":80")-3)			
			if strHttpsDomain="ON" then
				strURL = "https://" & strDomain & left(strURL,len(strURL))
			else
				strURL = "http://" & strDomain & left(strURL,len(strURL))
			end if
		end if
	End If
	'// Check for "www" and redirect to it if absent
	if intRedirectType=1 then
		if instr(strURL,"www")=0 then
			strURL = replace(strURL,"://","://www.")
			Response.Status="301 Moved Permanently" 
			Response.AddHeader "Location", strURL
		end if
	else
		if instr(strURL,"www")<>0 then
			strURL = replace(strURL,"://www.","://")
			Response.Status="301 Moved Permanently" 
			Response.AddHeader "Location", strURL
		end if
	end if
END IF
'////////////////////////////////////////////////////////////////////
'// END - Check for WWW and redirect if absent
'////////////////////////////////////////////////////////////////////
%>
<%
private const scIncHeader="1"
%>
<!--#include file="inc_jquery.asp" -->
<% 'SB S %>
<% If scSBStatus = "1" Then %>
<script type="text/javascript" src="<%=gv_RootURL%>/Widget/widget.js"></script>
<% End If %> 
<!--#include file="inc_sb.asp"-->
<% 'SB E %>
<%
'// ProductCart v4.5 - START: Built-in Integration with Google Analytics
if trim(scGoogleAnalytics)<>"" and not IsNull(scGoogleAnalytics) then
%>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', '<%=scGoogleAnalytics%>']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
<%
end if
if pcStrPageName<>"noscript.asp" then
if (pcStrPageName = "configurePrd.asp") OR (pcStrPageName = "Reconfigure.asp") then
%>
<noscript>
<meta http-equiv="refresh" content="0;URL=noscript.asp"/>
</noscript>
<%
end if
end if
'// ProductCart v4.5 - END
%>