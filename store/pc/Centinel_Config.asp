<!--#include file="Centinel_DateFormat.js" -->
<%
'===================================================================================================
'= Cardinal Commerce (http://www.cardinalcommerce.com)
'= This configuration file centralizes all Centinel related configurables.
'= These values are required to be defined to enable the samples for work properly.
'=	
'= Transaction Testing URL : https://centineltest.cardinalcommerce.com/maps/txns.asp
'=
'= Your Production Transaction URL, Processor Id, and Merchant Id were assigned to you
'= upon registration for the Cardinal Centinel service.
'= 
'= Term URL is the fully qualified URL to the Centinel_Authenticate.asp file that is provided in theses samples.
'===================================================================================================

call opendb()

query="SELECT pcPay_Centinel.pcPay_Cent_TransactionURL, pcPay_Centinel.pcPay_Cent_ProcessorId, pcPay_Centinel.pcPay_Cent_MerchantID, pcPay_Centinel.pcPay_Cent_Active, pcPay_Centinel.pcPay_Cent_Password FROM pcPay_Centinel WHERE (((pcPay_Centinel.pcPay_Cent_ID)=1));"

set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
TRANSACTIONURL = rs("pcPay_Cent_TransactionURL") 'TRANSACTIONURL
PROCESSORID = rs("pcPay_Cent_ProcessorId") 'PROCESSORID
MERCHANTID = rs("pcPay_Cent_MerchantID") 'MERCHANTID
pcPay_Cent_Password = rs("pcPay_Cent_Password")
pcPay_Cent_Active = rs("pcPay_Cent_Active")
set rs=nothing
call closedb()
if scSSL="0" or scSSL="" then
	tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/Centinel_Authenticate.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
else
	tempURL=replace(( scSslURL&"/"&scPcFolder&"/pc/Centinel_Authenticate.asp"),"//","/")
	tempURL=replace(tempURL,"https:/","https://")
	tempURL=replace(tempURL,"http:/","http://") 
end if

TERMURL = tempURL 'TERMURL
MESSAGEVERSION	= "1.7" 'MESSAGEVERSION
'///////////////////////////////////////////////
'// EDIT YOUR CARDINAL COMMERCE PASSWORD
TRANSACTIONPWD = pcPay_Cent_Password 'TRANSACTIONPWD
'///////////////////////////////////////////////

RESOLVETIMEOUT = 10000
SENDTIMEOUT    = 10000
CONNECTTIMEOUT = 10000
RECEIVETIMEOUT = 10000
%>



