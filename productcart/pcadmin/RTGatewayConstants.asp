<%
'// Gateway code assignments. Each gateway is assigned a unique gwCode. Order array so that gateways are shown alphabetically.
dim pcGWArray, pcGWClassification
pcGWArray="//,13,39,1,57,29,33,52,32,60,67,65,42,54,31,37,58,30,5,8,11,27,55,59,64,12,47,48,4,26,49,63,70,24,35,56,10,9,//"

'Gateway Friendly Names - For each new gateway, add a friendly name
function gwShowDesc()
	select case varNonActive
		case 1
			gwa="1"
			pcGWPaymentDesc="AuthorizeNet"
			pcGWPaymentURL="http://reseller.authorize.net/application.asp?id=220675"
			pcGWClassification="BANK"
		case 2
			gwa="2"
			pcGWPaymentDesc="PayFlow Pro"
			pcGWPaymentURL="http://PayPal.com"
			pcGWClassification="BANK"
		case 3
			gwpp="1"
			pcGWPaymentDesc="PayPal"
			pcGWPaymentURL="http://PayPal.com"
			pcGWClassification="ALLINONE"
		case 4
			gwpsi="1"
			pcGWPaymentDesc="PSiGate - XML API &amp; HTML API"
			pcGWPaymentURL="http://www.psigate.com/"
			pcGWClassification="BANK"
		case 5
			gwit="1" 
			pcGWPaymentDesc="iTransact, Inc."
			pcGWPaymentURL="http://www.itransact.com/"
			pcGWClassification="BANK"
		case 8
			gwlp="1"
			pcGWPaymentDesc="LinkPoint"
			pcGWPaymentURL="http://www.firstdata.com/linkpoint/"
			pcGWClassification="BANK"
		case 9
			gwa="2"
			pcGWPaymentDesc="PayFlow Link"
			pcGWPaymentURL="http://PayPal.com"
			pcGWClassification=""
		case 10
			gwwp="1"
			pcGWPaymentDesc="WorldPay - Select Junior"
			pcGWPaymentURL="http://support.worldpay.com/integrations/jnr/index.html"
			pcGWClassification="BANK"
		case 11
			gwmoneris="1"
			pcGWPaymentDesc="Moneris - eSelect Plus Direct Post"
			pcGWPaymentURL="https://www3.moneris.com/"
			pcGWClassification="BANK"
		case 12		
			gwPxPay="1"
			pcGWPaymentDesc="Payment Express &reg; PX Pay"
			pcGWPaymentURL="http://www.paymentexpress.com/technical_resources/ecommerce_hosted/pxpay.html"
			pcGWClassification="BANK"
		case 13
			gw2Checkout="1"
			pcGWPaymentDesc="2Checkout (2CO)"
			pcGWPaymentURL="http://www.2checkout.com/community/"
			pcGWClassification="ALLINONE"
		case 15
			gwfast="1"
			pcGWPaymentDesc="Fast Transact"
			pcGWPaymentURL="http://www.fasttransactonline.com/"
			pcGWClassification="NULL"
		case 19
			gwecho="1"
			pcGWPaymentDesc="ECHO - Electronic Clearing House, Inc"
			pcGWPaymentURL="http://www.echo-inc.com/"
			pcGWClassification="NULL"
		case 22
			gwconcord="1"
			pcGWPaymentDesc="Concord"
			pcGWPaymentURL="http://www.concordefsnet.com"
			pcGWClassification="NULL"
		case 23
			gwklix="1"
			pcGWPaymentDesc="viaKLIX"
			pcGWPaymentURL="https://www2.viaklix.com/Admin/main.asp"
			pcGWClassification="NULL"
		case 24
			gwtclink="1"
			pcGWPaymentDesc="TrustCommerce - TCLink"
			pcGWPaymentURL="http://www.trustcommerce.com/tclink.php"
			pcGWClassification="BANK"
		case 26
			gwprotx="1"
			pcGWPaymentDesc="Sage Pay (Protx)"
			pcGWPaymentURL="http://www.sagepay.com/"
			pcGWClassification="BANK"
		case 27
			gwnetbill="1"
			pcGWPaymentDesc="NETbilling"
			pcGWPaymentURL="http://www.netbilling.com/index.php"
			pcGWClassification="BANK"
		case 29
			gwBluePay="1"
			pcGWPaymentDesc="BluePay"
			pcGWPaymentURL="http://www.bluepay.com/"
			pcGWClassification="ALLINONE"
		case 30
			gwIntSecure="1"
			pcGWPaymentDesc="InternetSecure"
			pcGWPaymentURL="http://www.internetsecure.com/"
			pcGWClassification="BANK"
		case 31
			gwEway="1"
			pcGWPaymentDesc="eWay"
			pcGWPaymentURL="http://www.eway.com.au/"
			pcGWClassification="ALLINONE"
		case 32
			gwCys="1"
			pcGWPaymentDesc="CyberSource - Simple Order API"
			pcGWPaymentURL="http://www.cybersource.com/support_center/implementation/downloads/simple_order/matrix.html"
			pcGWClassification="BANK"
		case 33
			gwCBN="1"
			pcGWPaymentDesc="ChecksByNet by CrossCheck, Inc"
			pcGWPaymentURL="http://www.checksbynet.com/"
			pcGWClassification="ALTERNATIVE"
		case 35
			gwUep="1"
			pcGWPaymentDesc="USA ePay"
			pcGWPaymentURL="http://www.usaepay.com/"
			pcGWClassification="BANK"
		case 37
			gwFastCharge="1"
			pcGWPaymentDesc="Fastcharge"
			pcGWPaymentURL="http://www.fastcharge.com/"
			pcGWClassification="BANK"
		case 39
			gwACH="1"
			pcGWPaymentDesc="ACH Direct, Inc"
			pcGWPaymentURL="http://www.achdirect.com/"
			pcGWClassification="BANK"
		case 40
			gwNETOne="1"
			pcGWPaymentDesc="NET1 / Sage Payments"
			pcGWPaymentURL="http://www.sagepayments.com/Products.aspx?PageId=Net1Gateway"
			pcGWClassification="NULL"
		case 42
			gwEPN="1"
			pcGWPaymentDesc="eProcessing Network, LLC"
			pcGWPaymentURL="http://www.eprocessingnetwork.com/"
			pcGWClassification="BANK"
		case 43
			gwTripleDeal="1"
			pcGWPaymentDesc="Triple Deal"
			pcGWPaymentURL="http://corporate.tripledeal.com/"
			pcGWClassification="NULL"
		case 44
			gwHSBC="1"
			pcGWPaymentDesc="HSBC"
			pcGWPaymentURL="http://www.hsbc.co.uk/1/2/business/home"
			pcGWClassification="NULL"
		case 45			
			gwParaData="1"
			pcGWPaymentDesc="ParaData"
			pcGWPaymentURL="http://www.paypros.com/"
			pcGWClassification="NULL"
		case 47			
			gwPaymentExpress="1"
			pcGWPaymentDesc="Payment Express &reg; PX Post"
			pcGWPaymentURL="http://www.paymentexpress.com/products/non_hosted/pxpost_product.html"
			pcGWClassification="BANK"
		case 48			
			gwSecPay="1"
			pcGWPaymentDesc="PayPoint (formerly SECPay)"
			pcGWPaymentURL="http://www.secpay.com/secpay/index.php/secpay.html"
			pcGWClassification="NULL"
		case 49
			gwSkipJack="1"
			pcGWPaymentDesc="Skipjack"
			pcGWPaymentURL="http://www.Skipjack.com/"
			pcGWClassification="BANK"
		case 51
			gwEmerchant="1"
			pcGWPaymentDesc="Fasthosts eMerchant"
			pcGWPaymentURL="http://www.fasthosts.co.uk/ecommerce/merchant-account/"
			pcGWClassification="NULL"
		case 52			
			gwCP="1"
			pcGWPaymentDesc="ChronoPay"
			pcGWPaymentURL="http://www.chronopay.com/"
			pcGWClassification="BANK"
		case 54		
			gweMoney="1"
			pcGWPaymentDesc="ETS - EMoney<sup>TM</sup>"
			pcGWPaymentURL="https://www.etsms.com/ASP/emoney.htm"
			pcGWClassification="BANK"
		case 55		
			gwOgone="1"
			pcGWPaymentDesc="Ogone"
			pcGWPaymentURL="http://www.ogone.com/"
			pcGWClassification="BANK"
		case 56	
			gwVirtualMerchant="1"
			pcGWPaymentDesc="Virtual Merchant"
			pcGWPaymentURL="https://www.myvirtualmerchant.com/"
			pcGWClassification="BANK"
		case 57		
			gwBeanStream="1"
			pcGWPaymentDesc="Beanstream"
			pcGWPaymentURL="http://www.beanstream.com/website/index.asp"
			pcGWClassification="ALLINONE"
		case 58	
			gwGlobalPay="1"
			pcGWPaymentDesc="Global Pay"
			pcGWPaymentURL="http://www.globalpaymentsinc.com/"
			pcGWClassification="BANK"
		case 59	
			gwOmega="1"
			pcGWPaymentDesc="Omega"
			pcGWPaymentURL="http://www.omegap.com/"
			pcGWClassification="BANK"
		case 60	' echeck = 61
			gwDowCommmerce="1"
			pcGWPaymentDesc="Dow Commerce"
			pcGWPaymentURL="http://www.dowcommerce.com/"
			pcGWClassification="ALLINONE"
	    case 63	
			gwTotalWeb="1"
			pcGWPaymentDesc="TotalWeb Solutions"
			pcGWPaymentURL="http://www.totalwebsolutions.com"
			pcGWClassification="BANK"
		case 64	
			gwPayJunction="1"
			pcGWPaymentDesc="Pay Junction - QuickLink"
			pcGWPaymentURL="http://www.payjunction.com/"
			pcGWClassification="ALLINONE"
		case 65	
			gwTSG="1"
			pcGWPaymentDesc="EC Suite - Transaction Gateway System"
			pcGWPaymentURL="http://www.ecsuite.com/"
			pcGWClassification="ALLINONE"
		case 67	
			gwEIG="1"
			pcGWPaymentDesc="NetSource Commerce Gateway"
			pcGWPaymentURL="http://www.earlyimpact.com/"
			pcGWClassification="BANK"
		case 70
		 	gwTXP="1"
			pcGWPaymentDesc="Transaction Express™ - TransFirst"
			pcGWPaymentURL="https://www.transfirst.com/products/online-solutions/payment-gateway.htm"
			pcGWClassification="BANK"
		case 80
			gwPPA="1"
			pcGWPaymentDesc="PayPal Payments Advanced"
			pcGWPaymentURL="http://www.PayPal.com"
		case 46
			gwPPP="1"
			pcGWPaymentDesc="PayPal Payments Pro"
			pcGWPaymentURL="http://www.PayPal.com"
		case 53
			gwPPPUK="1"
			pcGWPaymentDesc="PayPal Payments Pro UK"
			pcGWPaymentURL="http://www.PayPal.com"
		case 999999
			gwPPEx="1"
			pcGWPaymentDesc="PayPal Express"
			pcGWPaymentURL="http://www.PayPal.com"
			pcGWClassification="ALLINONE"
		case 0
			pcGWPaymentDesc=""
			pcGWPaymentURL=""
			pcGWClassification=""
		end select
end function

function gwCallEdit()
	select case pcv_EditGW
		case "6"
			call OffLineCCEdit()
		case "7"
			call OffLineCustomEdit()
		case "2"
			call gwPFPEdit()
		case "3"
			call gwPPEdit()
		case "4"
			call gwpsiEdit()
		case "5"
			call gwitEdit()
		case "8"
			call gwlpEdit()
		case "9"
			call gwPFLinkEdit()
		case "10"
			call gwwpEdit()
		case "11"
			call gwmonerisEdit()	
		case "13" '2Checkout
			call gw2CheckoutEdit()
		case "15"
			call gwfastEdit()
		case "1" 'Authorize SIM
			call gwaEdit()
		case "19"
			call gwechoEdit()	
		case "22"
			call gwconcordEdit()
		case "23"
			call gwklixEdit()
		case "24"
			call gwtclinkEdit()	
		case "26"
			call gwprotxEdit()	
		case "27"
			call gwnetbillEdit()
		case "29"
			call gwBluePayEdit()
		case "30"
			call gwIntSecureEdit()
		case "31"
			call gwEwayEdit()
		case "32"
			call gwcysEdit()
		case "33"
			call gwCBNEdit()
		case "35"
			call gwuepEdit()
		case "37"
			call gwfastchargeEdit()
		case "39"
			call gwACHEdit()
		case "40"
			call gwNETOneEdit()
		case "42"
			call gwEPNEdit()
		case "43"
			call gwTripleDealEdit()	
		case "44"
			call gwHSBCEdit()
		case "45"
			call gwParaDataEdit()
		case "47"
			call gwPaymentExpressEdit()
		case "48"
			call gwSECPayEdit()
		case "48"
			call gwPxPayEdit()
		case "49"
			call gwSkipJackEdit()
		case "51"
			call gwEMerchantEdit()
		case "52"
			call gwCPEdit()
		case "12"
			call gwPxPayEdit()
		case "54"
			call gweMoneyEdit()
		case "55"
			call gwOgoneEdit()
		case "56"
			call gwVMPayEdit()
		case "57"
			call gwBeanStreamEdit()
		case "58"
			call gwGlobalPayEdit()
		case "59"
			call gwOmegaEdit()
		case "60"' echeck = 61
			call gwDowComEdit()
		case "63"
			call gwTotalWebEdit()
		case "64" 
			call gwPayJunctionEdit()
		case "65" 
			call gwTGSEdit()
		case "67" 
			call gwEIGEdit()
		case "70"
			call gwTXPEdit()
		case "80"
			call gwPPAEdit()
		case "46"
			call gwPPPEdit()
		case "53"
			call gwPPPUKEdit()
		case "99"
			call gwPFLEdit()
		case "999999"
			call gwPPExEdit()
	end select
end function

function gwCallAdd()
	select case pcv_AddGW
		case "2"
			call gwPFP()	
		case "3"
			call gwPP()	
		case "4"
			call gwPsi()	
		case "5"
			call gwit()
		case "8"
			call gwlp()
		case "9"
			call gwPFLink()
		case "10"
			call gwwp()
		case "11"
			call gwmoneris()	
		case "13" '2Checkout
			call gw2Checkout()
		case "15"
			call gwfast()
		case "1" 'Authorize SIM
			call gwa()
		case "19"
			call gwecho()	
		case "23"
			call gwklix()
		case "24"
			call gwtclink()	
		case "26"
			call gwprotx()	
		case "27"
			call gwnetbill()
		case "29"
			call gwBluePay()
		case "30"
			call gwIntSecure()
		case "31"
			call gwEway()
		case "32"
			call gwcys()
		case "33"
			call gwCBN()
		case "35"
			call gwuep()
		case "37"
			call gwfastcharge()
		case "39"
			call gwACH()
		case "40"
			call gwNETOne()
		case "42"
			call gwEPN()
		case "43"
			call gwTripleDeal()	
		case "44"
			call gwHSBC()
		case "45"
			call gwParaData()
		case "47"
			call gwPaymentExpress()
		case "48"
			call gwSECPay()
		case "48"
			call gwPxPay()
		case "49"
			call gwSkipJack()
		case "51"
			call gwEMerchant()
		case "52"
			call gwCP()
		case "12"
			call gwPxPay()
		case "54"
			call gweMoney()
		case "55"
			call gwOgone()
		case "56"
			call gwVMPay()
		case "57"
			call gwBeanStream()
		case "58"
			call gwGlobalPay()
		case "59"
			call gwOmega()
		case "60" ' echeck = 61
			call gwDowCom()
		case "63" 
			call gwTotalWeb()
		case "64" 
			call gwPayJunction()
		case "65" 
			call gwTGS()
		case "67" 
			call gwEIG()
		case "70"
			call gwTXP()
		case "80"
			call gwPPA()
		case "46"
			call gwPPP()
		case "53"
			call gwPPPUK()
		case "99"
			call gwPFL()
		case "999999"
			call gwPPEx()
	end select
end function
%>