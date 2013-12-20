<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<% pageTitle="Purge Credit Card Numbers" %>
<% Section="" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%dim query, rs, conntemp
dim intSuccessCnt
intSuccessCnt=0
dim intPurgeprocess
intPurgeprocess=0
cGateway=request("GW")
if request.Form("PurgeNumbers")<>"" then
	intPurgeprocess=1
	dim strSuccessData
	strSuccessData=""
	'get the count
	pc_CardCnt=request.Form("iCnt")
	dim i
	for i=1 to pc_CardCnt
		'see if checkbox is checked
		if request.Form("idOrder"&i)="1" then
			tempOrderId=request.Form("ccOrderID"&i)
			tempDisplayOrderId=request.Form("pOrderID"&i)
			'/get ccnumber from database
			call opendb()
			query="SELECT creditCards.cardnumber, creditCards.pcSecurityKeyID FROM creditCards WHERE idOrder="&tempOrderId&";"
			set rs=server.CreateObject("ADODB.RecordSet") 
			set rs=connTemp.execute(query)
			tempSecurityKeyID=rs("pcSecurityKeyID")
			tempCCNum=rs("cardnumber")
			tempfour=pcf_PurgeCardNumber(tempCCNum,tempSecurityKeyID)
			tempSecurityKeyID=""
			tempCCNum=""
			select case cGateway
				case "authorders"
					query="UPDATE authorders SET ccnum='"&tempfour&"' WHERE idauthorder="&tempOrderId&";"		
				
				case "lporders"
					query="UPDATE pcPay_LinkPointAPI SET pcPay_LPAPI_CCNum='"&tempfour&"' WHERE pcPay_LPAPI_ID="&tempOrderId&";"
					
				case "pfporders"
					query="UPDATE pfporders SET acct='"&tempfour&"' WHERE idpfporder="&tempOrderId&";"
					
				case "netbillorders"
					query="UPDATE netbillorders SET ccnum='"&tempfour&"' WHERE idnetbillorder="&tempOrderId&";"
					
				case "pcPay_USAePay_Orders"
					query="UPDATE pcPay_USAePay_Orders SET ccnum='"&tempfour&"' WHERE idePayOrder="&tempOrderId&";"
					
				case "pcPay_EIG_Authorize"
					query="UPDATE pcPay_EIG_Authorize SET ccnum='"&tempfour&"' WHERE idauthorder="&tempOrderId&";"
					
				case else				
					query="UPDATE creditcards SET cardnumber='"&tempfour&"', seqcode='na' WHERE idOrder="&tempOrderId&";"
					
			end select				
			set rs=server.CreateObject("ADODB.RecordSet") 
			set rs=connTemp.execute(query)
			
			call closedb()
			intSuccessCnt=intSuccessCnt+1
			strSuccessData=strSuccessData&"Credit card number successfully purged for <strong>order #"&(scpre+int(tempDisplayOrderId))&"</strong>&nbsp;|&nbsp;<a href=Orddetails.asp?id="&int(tempDisplayOrderId)&">View Order</a><BR>"
		end if
	next
end if

if request.Form("search1")<>"" then
'show results 
	call opendb()
	pcv_fromMonth=request.Form("fromMonth")
	pcv_fromDay=request.Form("fromDay")
	pcv_fromYear=request.Form("fromYear")
	pcv_fromDate=pcv_fromMonth&"/"&pcv_fromDay&"/"&pcv_fromYear
	pcv_toMonth=request.Form("toMonth")
	pcv_toDay=request.Form("toDay")
	pcv_toYear=request.Form("toYear")
	pcv_toDate=pcv_toMonth&"/"&pcv_toDay&"/"&pcv_toYear
	pcv_orderStatus=request.Form("orderStatus")
	pcv_captured=request.Form("captured")
	select case cGateway
		case "authorders"
			
			'// START: Authorize.net			
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((authorders.captured)=1)"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT authorders.idauthorder, authorders.idOrder, orders.orderDate, authorders.ccnum, authorders.pcSecurityKeyID FROM authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT authorders.idauthorder, authorders.idOrder, orders.orderDate, authorders.ccnum, authorders.pcSecurityKeyID FROM authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT authorders.idauthorder, authorders.idOrder, orders.orderDate, authorders.ccnum, authorders.pcSecurityKeyID FROM authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT authorders.idauthorder, authorders.idOrder, orders.orderDate, authorders.ccnum, authorders.pcSecurityKeyID FROM authorders INNER JOIN orders ON authorders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: Authorize.net
			
		case "pcPay_EIG_Authorize"
			
			'// START: EIG			
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((pcPay_EIG_Authorize.captured)=1)"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT pcPay_EIG_Authorize.idauthorder, pcPay_EIG_Authorize.idOrder, orders.orderDate, pcPay_EIG_Authorize.ccnum, pcPay_EIG_Authorize.pcSecurityKeyID FROM pcPay_EIG_Authorize INNER JOIN orders ON pcPay_EIG_Authorize.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_EIG_Authorize.idauthorder, pcPay_EIG_Authorize.idOrder, orders.orderDate, pcPay_EIG_Authorize.ccnum, pcPay_EIG_Authorize.pcSecurityKeyID FROM pcPay_EIG_Authorize INNER JOIN orders ON pcPay_EIG_Authorize.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT pcPay_EIG_Authorize.idauthorder, pcPay_EIG_Authorize.idOrder, orders.orderDate, pcPay_EIG_Authorize.ccnum, pcPay_EIG_Authorize.pcSecurityKeyID FROM pcPay_EIG_Authorize INNER JOIN orders ON pcPay_EIG_Authorize.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_EIG_Authorize.idauthorder, pcPay_EIG_Authorize.idOrder, orders.orderDate, pcPay_EIG_Authorize.ccnum, pcPay_EIG_Authorize.pcSecurityKeyID FROM pcPay_EIG_Authorize INNER JOIN orders ON pcPay_EIG_Authorize.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: EIG
			
		case "lporders"
		
			'// START: Link Point API			
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((pcPay_LinkPointAPI.pcPay_LPAPI_Captured)=1)"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT pcPay_LinkPointAPI.pcPay_LPAPI_ID, pcPay_LinkPointAPI.idOrder, orders.orderDate, pcPay_LinkPointAPI.pcPay_LPAPI_CCNum, pcPay_LinkPointAPI.pcSecurityKeyID FROM pcPay_LinkPointAPI INNER JOIN orders ON pcPay_LinkPointAPI.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_LinkPointAPI.pcPay_LPAPI_ID, pcPay_LinkPointAPI.idOrder, orders.orderDate, pcPay_LinkPointAPI.pcPay_LPAPI_CCNum, pcPay_LinkPointAPI.pcSecurityKeyID FROM pcPay_LinkPointAPI INNER JOIN orders ON pcPay_LinkPointAPI.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT pcPay_LinkPointAPI.pcPay_LPAPI_ID, pcPay_LinkPointAPI.idOrder, orders.orderDate, pcPay_LinkPointAPI.pcPay_LPAPI_CCNum, pcPay_LinkPointAPI.pcSecurityKeyID FROM pcPay_LinkPointAPI INNER JOIN orders ON pcPay_LinkPointAPI.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_LinkPointAPI.pcPay_LPAPI_ID, pcPay_LinkPointAPI.idOrder, orders.orderDate, pcPay_LinkPointAPI.pcPay_LPAPI_CCNum, pcPay_LinkPointAPI.pcSecurityKeyID FROM pcPay_LinkPointAPI INNER JOIN orders ON pcPay_LinkPointAPI.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: Link Point API
				
		case "pfporders"

			'// START: Payflo Pro
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((pfporders.captured)="&pcv_captured&")"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&" AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pfporders.idpfporder, pfporders.idOrder, orders.orderDate, pfporders.acct, pfporders.pcSecurityKeyID FROM pfporders INNER JOIN orders ON pfporders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&" AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: Payflo Pro
			
		case "netbillorders"
		
			'// START: Netbilling
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((netbillorders.captured)=1)"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT netbillorders.idnetbillorder, netbillorders.idOrder, orders.orderDate, netbillorders.ccnum, netbillorders.pcSecurityKeyID FROM netbillorders INNER JOIN orders ON netbillorders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT netbillorders.idnetbillorder, netbillorders.idOrder, orders.orderDate, netbillorders.ccnum, netbillorders.pcSecurityKeyID FROM netbillorders INNER JOIN orders ON netbillorders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT netbillorders.idnetbillorder, netbillorders.idOrder, orders.orderDate, netbillorders.ccnum, netbillorders.pcSecurityKeyID FROM netbillorders INNER JOIN orders ON netbillorders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT netbillorders.idnetbillorder, netbillorders.idOrder, orders.orderDate, netbillorders.ccnum, netbillorders.pcSecurityKeyID FROM netbillorders INNER JOIN orders ON netbillorders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: Netbilling
			
		case "pcPay_USAePay_Orders"
			
			'// START: USAePay
			if pcv_captured="1" then
				pcv_capturedquery=" AND ((pcPay_USAePay_Orders.captured)=1)"
			else
				pcv_capturedquery=""
			end if
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT pcPay_USAePay_Orders.idePayOrder, pcPay_USAePay_Orders.idOrder, orders.orderDate, pcPay_USAePay_Orders.ccCard, pcPay_USAePay_Orders.pcSecurityKeyID FROM pcPay_USAePay_Orders INNER JOIN orders ON pcPay_USAePay_Orders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_USAePay_Orders.idePayOrder, pcPay_USAePay_Orders.idOrder, orders.orderDate, pcPay_USAePay_Orders.ccCard, pcPay_USAePay_Orders.pcSecurityKeyID FROM pcPay_USAePay_Orders INNER JOIN orders ON pcPay_USAePay_Orders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&") ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT pcPay_USAePay_Orders.idePayOrder, pcPay_USAePay_Orders.idOrder, orders.orderDate, pcPay_USAePay_Orders.ccCard, pcPay_USAePay_Orders.pcSecurityKeyID FROM pcPay_USAePay_Orders INNER JOIN orders ON pcPay_USAePay_Orders.idOrder = orders.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT pcPay_USAePay_Orders.idePayOrder, pcPay_USAePay_Orders.idOrder, orders.orderDate, pcPay_USAePay_Orders.ccCard, pcPay_USAePay_Orders.pcSecurityKeyID FROM pcPay_USAePay_Orders INNER JOIN orders ON pcPay_USAePay_Orders.idOrder = orders.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)"&pcv_capturedquery&"  AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if
			'// END: USAePay
			
		case else	

			'// START: Offline Credit Card			
			if pcv_orderStatus=0 then
				if scDB="SQL" then
					query="SELECT creditCards.idOrder, creditCards.cardnumber, creditCards.pcSecurityKeyID, orders.orderDate, orders.orderDate, orders.orderStatus FROM orders INNER JOIN creditCards ON orders.idOrder = creditCards.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"')) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT creditCards.idOrder, creditCards.cardnumber, creditCards.pcSecurityKeyID, orders.orderDate, orders.orderDate, orders.orderStatus FROM orders INNER JOIN creditCards ON orders.idOrder = creditCards.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#)) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			else
				if scDB="SQL" then
					query="SELECT creditCards.idOrder, creditCards.cardnumber, creditCards.pcSecurityKeyID, orders.orderDate, orders.orderDate, orders.orderStatus FROM orders INNER JOIN creditCards ON orders.idOrder = creditCards.idOrder WHERE (((orders.orderDate)='"&pcv_fromDate&"' OR (orders.orderDate)>'"&pcv_fromDate&"') AND ((orders.orderDate)<'"&pcv_toDate&"' OR (orders.orderDate)='"&pcv_toDate&"') AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				else
					query="SELECT creditCards.idOrder, creditCards.cardnumber, creditCards.pcSecurityKeyID, orders.orderDate, orders.orderDate, orders.orderStatus FROM orders INNER JOIN creditCards ON orders.idOrder = creditCards.idOrder WHERE (((orders.orderDate)=#"&pcv_fromDate&"# OR (orders.orderDate)>#"&pcv_fromDate&"#) AND ((orders.orderDate)<#"&pcv_toDate&"# OR (orders.orderDate)=#"&pcv_toDate&"#) AND ((orders.orderStatus)="&pcv_orderStatus&")) ORDER BY orders.orderDate DESC, orders.idorder DESC;"
				end if
			end if	
			'// END: Offline Credit Card	
	
	end select

	set rs=server.CreateObject("ADODB.RecordSet") 
	'response.Write(query)
	'response.End()
	set rs=connTemp.execute(query)
	
	'show results
	dim iCnt
	iCnt=0
	if NOT rs.eof then 
		%>		
		<form name="form1" method="post" action="creditCardPurge.asp" class="pcForms">
        	<input type="hidden" name="GW" value="<%=cGateway%>">
            <p style="padding-left: 10px;">Select the  
            <%
            select case cGateway
                case "authorders"
                    response.Write("Authorize.net")			
                case "lporders"
                    response.Write("LinkPoint API")	
                case "pfporders"
                    response.Write("Payflo Pro")
                case "netbillorders"
                    response.Write("Netbilling")
                case "pcPay_USAePay_Orders"
                    response.Write("USAePay")
                case "pcPay_EIG_Authorize"
                    response.Write("EIG")
                case else				
                    response.Write("Credit Card Orders")
            end select	
            %>
            Orders for which you would like to purge credit card information.</p>
			<table class="pcCPcontent" style="width: 400px;">
                <tr><td colspan="3" class="pcCPspacer"></td></tr>
				<tr> 
					<th nowrap>&nbsp;</th>
					<th nowrap>Date of Order</th>
					<th nowrap>Order Number</th>
				</tr>
				<tr><td colspan="3" class="pcCPspacer"></td></tr>
				<% 
				do until rs.eof

					select case cGateway
						case "authorders"
							'// Authorize.net
							'pcv_idauthorder=rs("idauthorder")
							pcv_idOrder=rs("idauthorder")	
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("ccnum")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")
							
						case "pcPay_EIG_Authorize"
							'// EIG
							'pcv_idauthorder=rs("idauthorder")
							pcv_idOrder=rs("idauthorder")	
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("ccnum")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")

						case "lporders"
							'// Link Point API
							pcv_idOrder=rs("pcPay_LPAPI_ID")
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("pcPay_LPAPI_CCNum")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")
							
						case "pfporders"
							'// Payflo Pro
							pcv_idOrder=rs("idpfporder")
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("acct")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")

						case "netbillorders"
							'// Netbill
							pcv_idOrder=rs("idnetbillorder")
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("ccnum")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")
							
						case "pcPay_USAePay_Orders"
							'// USA ePay
							pcv_idOrder=rs("idePayOrder")
							pcv_intOrderId=rs("idOrder")
							pcv_orderDate=rs("orderDate")			
							pcv_CCnumber=rs("ccCard")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")
			
						case else	
							'// Offline Credit Card			
							pcv_idOrder=rs("idOrder")
							pcv_intOrderId=rs("idOrder")
							pcv_CCnumber=rs("cardnumber")
							pcv_SecurityKeyID = rs("pcSecurityKeyID")
							pcv_orderDate=rs("orderDate")
					end select
					
					pcv_SecurityPass = pcs_GetKeyUsed(pcv_SecurityKeyID)

					if isNull(pcv_CCnumber) OR pcv_CCnumber="" then
						pcv_CCnumber="*"
					end if
					
					pcv_DecryptedCC=enDeCrypt(pcv_CCnumber, pcv_SecurityPass)
					if (instr(pcv_DecryptedCC,"*")) OR (pcv_CCnumber="*") then
					else
						iCnt=iCnt+1
						%>
						<tr>
							<td width="20">
                                <input name="idOrder<%=iCnt%>" type="checkbox" value="1" checked>
                                <input type="hidden" name="ccOrderID<%=iCnt%>" value="<%=pcv_idOrder%>">
                                <input type="hidden" name="pOrderID<%=iCnt%>" value="<%=pcv_intOrderId%>">
                            </td>
							<td><%=pcv_orderDate%></td>
						    <td><a href="ordDetails.asp?id=<%=pcv_intOrderId%>" target="_blank"><%=(scpre+int(pcv_intOrderId))%></a></td>
						</tr>
					<% end if
					rs.movenext
				loop
				call closedb() %>
				<input name="iCnt" type="hidden" value="<%=iCnt%>">
				<% if iCnt=0 then %>
                    <tr>
                        <td colspan="3"><div class="pcCPmessage">No records found. <a href="creditCardPurge.asp">Back</a>.</div></td>
                    </tr>
				<% else %>
                    <tr><td colspan="3" class="pcCPspacer"></td></tr>
                    <tr>
                        <td colspan="3" nowrap>
                            <a href="javascript:checkAll();">Check All</a> |&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
                        </td>
                    </tr>
                    <tr><td colspan="3" class="pcCPspacer"></td></tr>
                    <tr>
                        <td colspan="3">
                            <input name="PurgeNumbers" type="submit" class="submit2" value="Purge Credit Card Numbers">
                            &nbsp;&nbsp;
                            <input type="button" class="ibtnGrey" value="Back" onClick="location='creditCardPurge.asp';">
                        </td>
                    </tr>
				<% end if %>
			</table>
			<% if iCnt>0 then %>
				<script language="JavaScript">
                    <!--
                    function checkAll() {
                    for (var j = 1; j <= <%=iCnt%>; j++) {
                    box = eval("document.form1.idOrder" + j); 
                    if (box.checked == false) box.checked = true;
                         }
                    }
                    
                    function uncheckAll() {
                    for (var j = 1; j <= <%=iCnt%>; j++) {
                    box = eval("document.form1.idOrder" + j); 
                    if (box.checked == true) box.checked = false;
                         }
                    }
                    -->
                </script>
            <% end if %>
    	</form>
	<% else %>
        <form name="form1" method="post" action="creditCardPurge.asp" class="pcForms">
            <p style="padding-left: 10px;">Select 
				<%
				select case cGateway
					case "authorders"
						response.Write("Authorize.net")			
					case "pcPay_PayPal_Authorize"
						response.Write("PayPal Website Payments Pro")	
					case "pfporders"
						response.Write("Payflo Pro")
					case "netbillorders"
						response.Write("Netbilling")
					case "pcPay_USAePay_Orders"
						response.Write("USAePay")
					case "pcPay_EIG_Authorize"
						response.Write("EIG")	
					case else				
						response.Write("Credit Card Orders")
				end select	
				%>
                Orders
            </p>
            <table class="pcCPcontent" style="width: 400px;">
                <tr><td colspan="3" class="pcCPspacer"></td></tr>
                <tr> 
                    <th nowrap>&nbsp;</th>
                    <th nowrap>Date of Order</th>
                    <th nowrap>Order Number</th>
                </tr>
                <tr><td colspan="3" class="pcCPspacer"></td></tr>
                <tr>
                    <td colspan="3">
                    	<div class="pcCPmessage">No records found. <a href="creditCardPurge.asp">Back</a>.</div>
                    </td>
                </tr>
                <tr><td colspan="3" class="pcCPspacer"></td></tr>
            </table>
        </form>
	<% end if %>
<% else
	if intPurgeprocess=0 then %>
		
        <table class="pcCPcontent">
        	<tr><td colspan="3" class="pcCPspacer"></td></tr>
			<tr> 
				<th>Search Orders by Date</th>
			</tr>
			<tr><td colspan="3" class="pcCPspacer"></td></tr>
			<tr> 
				<td>Enter a date range and select the status from the drop-menu below:  </td>
			</tr>
			<tr> 
				<td>
                    <form action="creditCardPurge.asp" method="post"  name="CCPurgeSearch" class="pcForms">
                        <%
                        FromDate=Date()
                        FromDate=FromDate-13
                        ToDate=Date()
                        %>
                        <table cellpadding="4" cellspacing="2">
                            <tr>
                                <td>Gateway:</td>
                                <td>
                                    <select id="GW" name="GW" onChange="showCapture();">
                                        <option value="offline">Offline Credit Card</option>
                                        <option value="authorders">Authorize.Net</option>
                                        <option value="pfporders">PayPal Payflow Pro</option>
                                        <option value="netbillorders">Netbilling</option>
                                        <option value="pcPay_USAePay_Orders">USAePay</option>
                                        <option value="lporders">LinkPoint API</option>
                                        <option value="pcPay_EIG_Authorize">NetSource Commerce Gateway</option>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td width="61">Date From:</td>
                                <td width="525"  valign="top">Month: 
                                    <input type=text name="fromMonth" value="<%=month(FromDate)%>" size="2" maxlength="4">
                                    Day:
                                    <input type=text name="fromDay" value="<%=day(FromDate)%>" size="2" maxlength="4"> 
                                    Year:
                                    <select name="fromYear">
                                        <% Dim varYear
                                        varYear=year(now) %>
                                        <option value="<%=varYear-4%>"><%=varYear-4%></option>
                                        <option value="<%=varYear-3%>"><%=varYear-3%></option>
                                        <option value="<%=varYear-2%>"><%=varYear-2%></option> 
                                        <option value="<%=varYear-1%>"><%=varYear-1%></option>
                                        <option value="<%=varYear%>" selected><%=varYear%></option>
                                </select></font></td>
                            </tr>
                            <tr>
                                <td>Date To:</td>
                                <td>Month:
                                    <input type=text name="toMonth" value="<%=month(ToDate)%>" size="2" maxlength="4">
                                    Day:
                                    <input type=text name="toDay" value="<%=day(ToDate)%>" size="2" maxlength="4">
                                    Year:
                                    <select name="toYear">
                                    <% 
                                    varYear=year(now) %>
                                        <option value="<%=varYear-4%>"><%=varYear-4%></option>
                                        <option value="<%=varYear-3%>"><%=varYear-3%></option>
                                        <option value="<%=varYear-2%>"><%=varYear-2%></option> 
                                        <option value="<%=varYear-1%>"><%=varYear-1%></option>
                                        <option value="<%=varYear%>" selected><%=varYear%></option>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Status:</td>
                                <td>
                                    <select name="orderStatus" id="orderStatus">
                                        <option value="0" selected>All orders (regardless of order status)</option>
                                        <option value="3">Processed</option>
                                        <option value="4">Shipped</option>
                                        <option value="5">Canceled</option>
                                        <option value="6">Return</option>
                                    </select> 
                                </td>
                          	</tr>
                            <tr>
                                <td colspan="2">
                                   <div id="capture" style="display:none"><input name="captured" type="checkbox" value="1">  Only show captured orders</div>
                                	<script language="javascript">
                                		function showCapture() {
											var a;
											a =  document.getElementById("GW").value;
											if (a=='offline') {
												document.getElementById("capture").style.display = "none";
											} else {
												document.getElementById("capture").style.display = "";
											}										
										}
										showCapture();
                                    </script>
                                </td>
                            </tr>
                            <tr>
                                <td>&nbsp;</td>
                                <td>
                                	<input type="submit" name="search1" value="View " class="submit2">
									&nbsp;&nbsp;
                        			<input type="button" class="ibtnGrey" value="Back" onClick="location='creditCardPurge_index.asp';">
                                </td>
                            </tr>
                    	</table>
					</form>
				</td>
			</tr>			
			<tr>
				<td>&nbsp;</td>
			</tr>			
		</table>
	<% else 
		if intSuccessCnt=0 then %>
			<table class="pcCPcontent">
				<tr>
					<td>No credit card numbers were purged. You must check at least one order.  <td>
				</tr>
				<tr>
					<td>
                    	<a href="creditCardPurge.asp">Back</a>
                    </td>
				</tr>
			</table>
		<% else %>
			<table class="pcCPcontent">
				<tr>
					<td><font color="#FF0000"><strong><%=intSuccessCnt%></strong></font>&nbsp;Credit card numbers were successfully purged for the selected orders:<br>
						<% if strSuccessData<>"" then %>
							<br><%=strSuccessData%><br>
						<% end if %>
					</td>
				</tr>
				<tr>
					<td><p>&nbsp;</p>
					<p><a href="resultsAdvancedAll.asp?B1=View%2BAll&dd=1">Manage Orders</a></p></td>
				</tr>
			</table>
		<% end if %>
	<% end if %>
<% end if %>
<!--#include file="AdminFooter.asp"-->