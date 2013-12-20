<!--#include file="../includes/pcSBClassInc.asp"-->
<!--#include file="../includes/pcSBBase64.asp"-->
<!--#include file="../includes/pcSBSettings.asp"-->
<!--#include file="SubscriptionBridge/spry_includes.asp"-->
<%
Dim tmp_setup
tmp_setup=0

call opendb()
query="SELECT Setting_RegSuccess FROM SB_Settings WHERE Setting_ID=1;"
set rs=connTemp.execute(query)
if not rs.eof then
	tmp_setup=1
end if
set rs=nothing
call closedb()

if tmp_setup=0 and (pageName<>"sb_Default.asp" and pageName<>"sb_manageAcc.asp") then
	response.Redirect("sb_Default.asp")
end if
%>