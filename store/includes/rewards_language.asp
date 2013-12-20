<%
if session("rewards_language")="" then
   session("rewards_language")="english"
end if
Dim dictRewardsLanguage
set dictRewardsLanguage = CreateObject("Scripting.Dictionary")
 
dictRewardsLanguage.Add "english_CustPref_11",   "View your " 

dictRewardsLanguage.Add "english_order_AA", "Enter Points to Use Towards this Purchase:" 
 
dictRewardsLanguage.Add "english_orderverify", RewardsLabel   
 
%>