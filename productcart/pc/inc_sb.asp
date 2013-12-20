<%
private const scSB="1"

Dim pSubscriptionID, pcv_intBillingFrequency, pcv_strBillingPeriod, pcv_intBillingCycles, pSubStartImmed,pSubStartFromPurch, pSubStart, pcv_intTrialCycles,pcv_curTrialAmount,pSubStartDate, pSubType, pSubInstall, pcv_intIsTrial, pSubAddToMail, pSubReOccur, pcv_intBillingCyclesUntDate,pcv_intTrialCyclesUntDate
%>
<!--#include file="../includes/pcSBSettings.asp"-->
<!--#include file="../includes/pcSBBase64.asp"-->
<!--#include file="../includes/pcSBClassInc.asp"-->
<!--#include file="../includes/pcSBHelperInc.asp"-->
<link type="text/css" rel="stylesheet" href="subscriptionBridge.css" />