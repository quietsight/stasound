<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="ProductCart Online Help - Troubleshooting Utility" %>
<% Section="" %>
<%PmAdmin=0%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/ppdstatus.inc"-->
<%
pcPageName="pcTSUtility.asp"

Dim conntemp, query, rs, rs2, Obj

' get payment types
err.clear
err.number=0
call openDb()  

Dim HaveImgUplResizeObjs
Dim pcv_UploadObj
Dim	pcv_ResizeObj

HaveImgUplResizeObjs=0
pcv_UploadObj=0
pcv_ResizeObj=0

query="SELECT gwCode, active FROM paytypes;"
set rs=Server.CreateObject("ADODB.Recordset")     
set rs=connTemp.execute(query)
if err.number <> 0 then
	set rs=nothing
	call closedb()
	response.redirect "techErr.asp?error="& Server.Urlencode("Error in pcTSUtility: "&Err.Description) 
end if
if NOT rs.eof then 
	iCnt=1
	actGW=0
	do until rs.eof
		varTemp=rs("gwCode")
		varActive=rs("active")
		if varTemp<>"0" then
			select case varTemp
			case 1
				gwa="1"
				actGW=1
			case 2
				gwvpfp="1"
				actGW=1
			case 3
				gwpp="1"
				gwppstandard="1"
				actGW=1
			case 4
				gwpsi="1"
				actGW=1
			case 5
				gwit="1" 
				actGW=1
			case 8
				gwlp="1"
				actGW=1
			case 9
				gwvpfl="1"
				actGW=1
			case 10
				gwwp="1"
				actGW=1
			case 11
				gwmoneris="1"
				actGW=1
			case 13
				gw2Checkout="1"
				actGW=1
			case 14
				gwAIM="1"
				actGW=1
			case 15
				gwfast="1"
				actGW=1
			case 18
				gwtcs="1"
				actGW=1
			case 19
				gwecho="1"
				actGW=1
			case 22
				gwconcord="1"
				actGW=1
			case 23
				gwklix="1"
				actGW=1
			case 24
				gwtclink="1"
				actGW=1
			case 26
				gwprotx="1"
				actGW=1
			case 27
				gwnetbill="1"
				actGW=1
			case 29
				gwBluePay="1"
				actGW=1
			case 30
				gwIntSecure="1"
				actGW=1
			case 31
				gwEway="1"
				actGW=1
			case 32
				gwCys="1"
				actGW=1
			case 33
				gwCBN="1"
				actGW=1
			case 34
				gwPaymentech="1"
				actGW=1
			case 35
				gwUep="1"
				actGW=1
			case 39
				gwACH="1"
				actGW=1
			case 40
				gwNETOne="1"
				actGW=1
			case 41
				gwGestPay="1"
				actGW=1
			case 42
				gwEPN="1"
				actGW=1
			case 43
				gwTripleDeal="1"
				actGW=1
			case 44
				gwHSBC="1"
				actGW=1
			case 45			
				gwParaData="1"
				actGW=1
			case 46
				gwpp="1"
				PayPalWP="1"
				actGW=1
			case 47			
				gwPaymentExpress="1"
				actGW=1
			case 48			
				gwSecPay="1"
				actGW=1
			case 49
			    gwSkipJack="1"
				actGW=1
		    case 51
				gwEmerchant="1"
				actGW=1
			case 52			
				gwCP="1"
				actGW=1
			case 12		
				gwPxPay="1"
				actGW=1
			case 54		
				gweMoney="1"
				actGW=1
			case 55		
				gwOgone="1"
				actGW=1
			case 56	
				gwVirtualMerchant="1"
				actGW=1
			case 57		
				gwBeanStream="1"
				actGW=1
			case 58	
				gwGlobalPay="1"
				actGW=1
			case 59	
				gwOmega="1"
				actGW=1
			case 60	' echeck = 61
				gwDowCommmerce="1"
				actGW=1
			case 63	
				gwTotalWeb="1"
				actGW=1
			 case 64	
				gwPayJunction="1"
				actGW=1
			case 999999
				gwpp="1"
				PayPalExp="1"
				actGW=1
			end select
		end if
		rs.moveNext
	loop
end if
set rs=nothing
'call closedb()


'Simple array for component names, 2-D array for class names (multiple possible classes per component)
Dim strComponent(9)
Dim strClass(9,2)

'The following require an installed class:
strComponent(0) = "USAePay"
strComponent(1) = "ACHDirect"
strComponent(2) = "FastCharge"
strComponent(3) = "ParaData"
strComponent(4) = "PSIGate"
strComponent(5) = "Cybersource"
strComponent(6) = "PayPal"
strComponent(7) = "Pay Junction (WinHttp.WinHttpRequest.5) "
strComponent(8) = "Pay Junction (WinHttp.WinHttpRequest.5.1)"
strComponent(9) = "LinkPoint API"

'The component class names
strClass(0,0) = "USAePayXChargeCom2.XChargeCom2"
strClass(1,0) = "SendPmt.clsSendPmt"
strClass(2,0) = "ATS.SecurePost"
strClass(3,0) = "Paygateway.EClient.1"
strClass(4,0) = "MyServer.PsiGate"
strClass(5,0) = "CyberSourceWS.MerchantConfig"
strClass(6,0) = "Com.paypal.sdk.COMNetInterop.COMAdapter"
strClass(7,0) = "WinHttp.WinHttpRequest.5"
strClass(8,0) = "WinHttp.WinHttpRequest.5.1"
strClass(9,0) = "LpiCom_6_0.LPOrderPart"
strClass(9,1) = "LpiCom_6_0.LinkPointTxn"

Function IsObjInstalled(intClassNum)
On Error Resume Next
'This function tests the classes for the component indicated by the passed-in number.  Uses elements in the strClasses array above, correlated with the component names in the strComponent array.  Returns a string with non-present class names if any are missing, otherwise returns a ZLS.

	'Increase this constant to reflect the ubound of the classes array if a component requires more classes.
	CONST CLASSBOUND = 2

	Dim objTest, j
	Dim strError
	
	'init
	strError = ""

	'Test up to possible number of classes for each component...
	for j = 0 to CLASSBOUND

		'depending on whether or not there's a class name present in the array.
		If Not IsEmpty(strClass(intClassNum, j)) Then
			Set objTest = Server.CreateObject(strClass(intClassNum, j))
			If Err.Number = 0 Then
				Set objTest = Nothing
			Else
				'If the class test failed, create or append to an error string for reporting results.
				If IsObject(objTest) Then Set objTest = Nothing
				If strError = "" Then
					strError = strClass(intClassNum, j)
				Else
					strError = strError & ", " & strClass(intClassNum, j)
				End If
			End If
		Else
			Exit For
		End If
	Next

	'Report any resulting errors.
	IsObjInstalled = strError

End Function
%>

<!--#include file="AdminHeader.asp"-->
<style>
	#pcCPmain h2 {
		margin-top: 25px;
	}
</style>
<script language="JavaScript"><!--
<!-- Hide me
function wincomcheckobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=no,status=no,width=300,height=250')
	myFloater.location.href=fileName;
	}
function wingwaycheckobj(fileName)
	{
	myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=475,height=425')
	myFloater.location.href=fileName;
	}
// show me-->
//-->
</script>

<% '// Start Web server readiness%>
<form name="form1" method="post" action="pcTSUtilitySave.asp">
<table class="pcCPcontent">
  	<tr>
    	<td>
			<h2 style="margin-top: 0px;">Version and License</h2>
			<p>You are using ProductCart <strong>version <%=scVersion&scSubVersion%><% if scSP<>"" then Response.Write(" Service Pack " & scSP) end if %><% if PPD="1" then Response.Write(" PPD") end if %></strong></p>
            <% if session("PmAdmin")="19" then %>
      		<p>Your ProductCart License Key is: <strong><%=scCrypPass%></strong></p>
            <% end if %>
                        
     		<h2>Store URL</h2>
            <p>The current Store URL is: <strong><%=scStoreURL%></strong>
            <p>The folder name variable (<em>includes/productcartFolder.asp</em>) is set to: <strong><%=scPcFolder%></strong></p>		
			<%
				Dim StoreUrl
				StoreUrl=0
				if instr(scStoreURL,"http://")>0 or instr(scStoreURL,"https://")>0 then
					StoreUrl=1
				end if
				
				Dim TempStoreUrl, StoreUrlUseIP
				StoreUrlUseIP=0
				TempStoreUrl = scStoreURL
				TempStoreUrl = replace(TempStoreUrl,".","")
				TempStoreUrl = replace(TempStoreUrl,"/","")
				TempStoreUrl = replace(TempStoreUrl,"http:","")
				TempStoreUrl = replace(TempStoreUrl,"https:","")
				if validNum(TempStoreUrl) then StoreUrlUseIP=1
				
				if StoreUrl=1 then %>
					<table width="100%">
						<tr> 
							<td width="4%"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
							<td width="96%">The Store URL appears to be configured properly. <% if StoreUrlUseIP=1 then %>However, it is using an IP address.<% end if %></td>
						</tr>
					</table>
				<% else %>
					<table>
						<tr> 
							<td width="4%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
							<td width="96%">Your Store URL ( <b><%=scStoreURL%></b> ) is NOT configured properly.<br>
								There is an issue with the &quot;Store URL&quot; stored in the file &quot;includes/storeconstants.asp&quot;. The Store URL must be the full URL to the folder on your server that contains ProductCart, including &quot;http://&quot;. To correct the problem, download the file &quot;includes/storeconstants.asp&quot; to your desktop using your favorite FTP program, open it in Notepad or an HTML editor, edit the line that contains the store address (&quot;scStoreURL&quot; constant), save the file and reupload it to your server.
							</td>
						</tr>
					</table>
				<% end if %>
                <% if StoreUrlUseIP=1 then %>
                	<table>
						<tr> 
							<td width="4%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
						  <td width="96%"><strong>Using an IP address </strong><br>The &quot;Store URL&quot; currently appears to be using an IP address. This is probably because when you originally installed ProductCart, there was no domain name pointing to the store. If now there is a domain name pointing to the store, that domain name should replace the IP address in the &quot;Store URL&quot; (<a href="http://wiki.earlyimpact.com/how_to/change_store_url" target="_blank">how to edit the Store URL</a>). Otherwise, there can be a number of technical issues.</td>
						</tr>
					</table>
                <% end if %>
                <% if instr(scStoreURL, "/"&scPcFolder) then %>
                	<table>
						<tr> 
							<td width="4%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
						  <td width="96%"> There is a possible issue with the &quot;Store URL&quot; stored in the file &quot;includes/storeconstants.asp&quot;. The &quot;Store URL&quot; path currently contains the directory &quot;/<%=scPcFolder%>&quot;; which is the same name as your &quot;ProductCart&quot; directory name.   The URL must be the full URL to the ProductCart directory, but should  NOT contain the ProductCart directory name. To correct the problem, download the file &quot;includes/storeconstants.asp&quot; to your desktop using your favorite FTP program, open it in Notepad or an HTML editor, edit the line that contains the store address (&quot;scStoreURL&quot; constant), save the file and reupload it to your server. </td>
						</tr>
					</table>
                <% end if %>

      		<h2>Parent Paths</h2>
	        <% on error resume next
				dim allstop
				allstop=0 
				dim findit
				if PPD="1" then
					PageName="/"&scPcFolder&"/includes/diagtxt.txt"
				else
					PageName="../includes/diagtxt.txt"
				end if
				findit=Server.MapPath(PageName)
				if err.number=0 OR PPD="1" then %>
					<table>
						<tr> 
							<td width="4%"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
							<td width="96%">
								<% if PPD="1" then
								response.write "This store appears to be using the Parent Paths <strong>Disabled</strong> version of ProductCart"
								else
								response.write "This store appears to be using the Parent Paths <strong>Enabled</strong> version of ProductCart"
								end if %>
							</td>
						</tr>
					</table>
				<% else
					allstop=1 %>
					<table>
						<tr> 
							<td width="4%"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
							<td width="96%">Parent Paths are disabled on this server. Please use the version of ProductCart that is included in the &quot;Parent Paths Disabled&quot; folder within the ProductCart installation files</td>
						</tr>
					</table>
				<% end if %>

				<% if allstop=0 then %>
				<h2>Folder permissions</h2>
					<% 
						Dim fso, f, errpermissions, errdelete_includes, errwrite_includes, errwrite_others
						errpermissions=0
						errdelete_includes=0
						errwrite_includes=0
						errwrite_others=0
						Set fso=server.CreateObject("Scripting.FileSystemObject")
						Set f=fso.GetFile(findit)
						Err.number=0
						f.Delete
						if Err.number>0 then
							errdelete_includes=1
							errpermissions=1
							Err.number=0
						end if
						'Set f=nothing
									
						Set f=fso.OpenTextFile(findit, 2, True)
						f.Write "test done"
						if Err.number>0 then
							errwrite_includes=1
							errpermissions=1
							Err.number=0
						end if
						
						if PPD="1" then
							PageName="/"&scPcFolder&"/pc/diagtxt.txt"
						else			
							PageName="../pc/diagtxt.txt"
						end if
						findit=Server.MapPath(PageName)
						Set f=fso.OpenTextFile(findit, 2, True)
						f.Write "test done"
						if Err.number>0 then
							errwrite_others=1
							errpermissions=1
							Err.number=0
						end if
									
						f.Close
						Set fso=nothing
						Set f=nothing
						if errpermissions=0 then %>
							<table>
								<tr> 
									<td width="4%"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
									<td width="96%">Folder permissions have been assigned correctly.</td>
								</tr>
							</table> 
						<% else %>
							<% if errwrite_others=1 or errwrite_includes=1 then %> 
                <table>
                  <tr> 
                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                    <td width="95%"><font color="#CC3950">You need to assign 'read/write' permissions to the 'productcart' folder and all of its subfolders.</font></td>
                  </tr>
                </table>
								<% end if
								if errdelete_includes=1 then %>
                <table>
                  <tr> 
                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                    <td width="95%"><font color="#CC3950">You need to assign 'read/write/delete' permissions to the 'productcart/includes' folder and all of its subfolders.</font></td>
                  </tr>
                  <tr> 
                </table>
                <% end if
				end if %>
              <h2>Database connection string</h2>
                <table>
                  <tr> 
                    <td width="4%" valign="top"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
                    <td width="96%">The connection string is valid. When you activated ProductCart, you entered <strong>MS <%=scDB%></strong> as your database type. If this is no longer the case:
										<ul>
											<li>Download the file &quot;includes/storeconstants.asp&quot; to your desktop via FTP.</li>
											<li>Open it with Notepad or your HTML editor and edit the scDB constant. If you are using MS Access, it should read scDB=&quot;Access&quot;. If you are using MS SQL, it should read scDB=&quot;SQL&quot;.</li>
											<li>Save the file and upload it back to your Web server.</li>
										</ul>
										</td>
                      </tr>
                    </table>
                
                    <h2>Available Upload Components</h2>
                    <table>
                    	<tr>
                        	<td colspan="3">Having an upload component installed on your Web server allows you to upload images and other files directly from the ProductCart Control Panel. It is also required for the <a href="http://wiki.earlyimpact.com/productcart/products_adding_new#how_to_upload_images" target="_blank">Image Upload &amp; Resize Feature</a>.</td>
                        </tr>
                      <tr> 
                        <td width="34%" nowrap style="color:#336699;">SoftArtisans FileUp</td>
                        <td width="5%">
							<% If Not IsMailObjInstalled("SoftArtisans.FileUp") Then
								response.write "<img src=images/pcv4_icon_alert_sm.gif>"
							else
								response.write "<img src=images/pc_checkmark_sm.gif>"
								uplComCnt = uplComCnt + 1
							end if %></td>
                            <td width="62%" nowrap style="color:#336699;"> 
                              <% If Not IsMailObjInstalled("SoftArtisans.FileUp") Then %>
                               Not Installed 
                                <% Else %>
                                Installed
                                <% End If %>
                            </td>
                        </tr>
                      <tr> 
                        <td width="34%" nowrap style="color:#336699;">aspSmartUpload</td>
                        <td width="5%">
							<% If Not IsMailObjInstalled("aspSmartUpload.SmartUpload") Then
								response.write "<img src=images/pcv4_icon_alert_sm.gif>"
							else
								response.write "<img src=images/pc_checkmark_sm.gif>"
								uplComCnt = uplComCnt + 1
							end if %>
                          </td>
                          <td width="62%" nowrap style="color:#336699;"> 
                              <% If Not IsMailObjInstalled("aspSmartUpload.SmartUpload") Then %>
                               Not Installed 
                                <% Else %>
                                Installed
                                <% End If %>
                          </td>
                      </tr>
                      <tr> 
                        <td width="34%" nowrap style="color:#336699;">Persits Upload</td>
                        <td width="5%">
							<% If Not IsMailObjInstalled("Persits.Upload") Then
								response.write "<img src=images/pcv4_icon_alert_sm.gif>"
							else
								response.write "<img src=images/pc_checkmark_sm.gif>"
								uplComCnt = uplComCnt + 1
							end if %>
                            </td>
                            <td width="62%" nowrap style="color:#336699;"> 
                              <% If Not IsMailObjInstalled("Persits.Upload") Then %>
                               Not Installed 
                                <% Else %>
                                Installed
                                <% End If %>
                            </td>
                        </tr>
                    </table>
					<% if uplComCnt>0 then %>
                        <table>
                          <tr> 
                            <td width="4%" valign="top"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
                            <td width="96%">Compatible upload components are available on this server.</td>
                          </tr>
                        </table>
					<% else %>
                        <table>
                          <tr> 
                            <td width="5%" valign="middle" align="right"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                            <td width="95%" style="color:#CC3950">We are unable to detect a supported upload component on this server. ProductCart supports the components listed above. Please contact your Web hosting company and ask them to install an upload component for you. Otherwise, you can upload images using FTP software.</td>
                          </tr>
                        </table>
					<% end if %>
                    
                    <h2>Available Image Processing Components</h2>
                    <table>
                    	<tr>
                        	<td colspan="3">This is required for the <a href="http://wiki.earlyimpact.com/productcart/products_adding_new#how_to_upload_images" target="_blank">Image Upload &amp; Resize Feature</a>.</td>
                        </tr>
                      <tr> 
                        <td width="34%" nowrap style="color:#336699;">Persits Jpeg (ASPJpeg)</td>
                        <td width="5%">
							<% If Not IsMailObjInstalled("Persits.Jpeg") Then
								response.write "<img src=images/pcv4_icon_alert_sm.gif>"
							else
								response.write "<img src=images/pc_checkmark_sm.gif>"
								imgComCnt = imgComCnt + 1
							end if %>
                            </td>
                            <td width="62%" nowrap style="color:#336699;"> 
                              <% If Not IsMailObjInstalled("Persits.Jpeg") Then %>
                               Not Installed 
                                <% Else %>
                                Installed
                                <% End If %>
                            </td>
                        </tr>
                      <tr> 
                        <td width="34%" nowrap style="color:#336699;">Asp Image</td>
                        <td width="5%">
							<% If Not IsMailObjInstalled("AspImage.Image") Then
								response.write "<img src=images/pcv4_icon_alert_sm.gif>"
							else
								response.write "<img src=images/pc_checkmark_sm.gif>"
								imgComCnt = imgComCnt + 1
							end if %>
                            </td>
                            <td width="62%" nowrap style="color:#336699;"> 
                              <% If Not IsMailObjInstalled("AspImage.Image") Then %>
                               Not Installed 
                                <% Else %>
                                Installed
                                <% End If %>
                            </td>
                        </tr>
                    </table>
					<% if imgComCnt>0 then %>
                        <table>
                          <tr> 
                            <td width="4%" valign="top">
							<img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
                            <td width="96%">Compatible image processing components are available on this server.</td>
                          </tr>
                        </table>
					<% else %>
                        <table>
                          <tr> 
                            <td width="5%" valign="middle" align="right"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                            <td width="95%" style="color:#CC3950">We are unable to detect an image processing components on this server. ProductCart supports the components listed above. Please contact your Web hosting company, and ask them to install an upload component for you.</td>
                          </tr>
                        </table>
					<% end if %>
                
              <h2>Available E-mail Components</h2>
              <%
					Dim theComponent(7)
					Dim theComponentName(7)
					
					' the components
					theComponent(0) = "CDONTS.NewMail"
					theComponent(1) = "Bamboo.SMTP"
					theComponent(2) = "SMTPsvg.Mailer"
					theComponent(3) = "JMail.SMTPMail"
					theComponent(4) = "JMail.Message"
					theComponent(5) = "CDO.Message"
					theComponent(6) = "ABMailer.Mailman"
					theComponent(7) = "Persits.MailSender"
					
					' the name of the components
					theComponentName(0) = "CDONTS"
					theComponentName(1) = "Bamboo SMTP"
					theComponentName(2) = "ServerObjects ASPMail"
					theComponentName(3) = "JMail 3.7"
					theComponentName(4) = "JMail 4"
					theComponentName(5) = "CDOSYS"
					theComponentName(6) = "ABMailer"
					theComponentName(7) = "Persits ASPMail"
					
					Function IsMailObjInstalled(strClassString)
						 On Error Resume Next
						 ' initialize default values
						 IsMailObjInstalled = False
						 Err = 0
						 ' testing code
						 Dim xTestObj
						 Set xTestObj = Server.CreateObject(strClassString)
						 If 0 = Err Then IsMailObjInstalled = True
						 ' cleanup
						 Set xTestObj = Nothing
						 Err = 0
					End Function
				%>
                <table>
                <% Dim i, emailComCnt
					emailComCnt=0
					For i=0 to UBound(theComponent) %>
                	<tr> 
                    	<td width="70%" nowrap><%= theComponentName(i)%></td>
	                    <td nowrap> 
    	                  <% If Not IsMailObjInstalled(theComponent(i)) Then %>
        		               	NOT INSTALLED
                	      <% Else %>
						  		<% emailComCnt=emailComCnt+1 %>
		                        <b>INSTALLED</b>	
        	              <% End If %>
            	        </td>
					</tr>
                  <% Next %>
                </table>
				<% if emailComCnt>0 then %>
                <table>
                  <tr> 
                    <td width="4%" valign="top"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
                    <td width="96%">Compatible e-mail components are available on this server.</td>
                  </tr>
                </table>
				<% else %>
				<table>
                  <tr> 
                    <td width="4%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                    <td width="96%">We are unable to detect a supported E-mail component on this server. ProductCart supports the components listed above. Please contact your Web hosting company and ask them to install an e-mail component for you.</td>
                  </tr>
              </table>
			<% end if %>
            <% 'if UPS is active show info 
			query="SELECT ShipmentTypes.active, ShipmentTypes.userID, ShipmentTypes.password, ShipmentTypes.AccessLicense FROM ShipmentTypes WHERE (((ShipmentTypes.shipmentDesc)='UPS') AND ((ShipmentTypes.active)<>0));"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=connTemp.execute(query)
			if err.number <> 0 or rs.eof then
			else
				pcv_UPS_UserID=rs("userID")
				pcv_UPS_Password=rs("password")
				pcv_UPS_AccessLicense=rs("AccessLicense")
				%>
              	<h2>UPS License</h2>
              	<p>If you have an account with UPS with negotiated rates; you  will need to activate your OnLine Tools account that was created when you  register through the ProductCart application. Provide UPS with the following  credentials. They will then match your UPS negotiated account with the account  created in ProductCart.</p>
              	<br>
              	<p><strong>Access License Number:</strong>  <%=pcv_UPS_AccessLicense%><br>
				  <strong>User Id:</strong> <%=pcv_UPS_UserID%><br>
				  <strong>Password:</strong> <%=pcv_UPS_Password%><br>
              	</p>
			<% End if %>
              <h2>XML Parser</h2>
			<% 
            select case scXML
                case ".2.6"
                    strXMLUse="MSXML2 v2.6"
                    session("XMLUse")=".2.6"
                case ".3.0"
                    strXMLUse="MSXML2 v3.0"
                    session("XMLUse")=".3.0"
                case ".4.0"
                    strXMLUse="MSXML2 v4.0"
                    session("XMLUse")=".4.0"
                case ".5.0"
                    strXMLUse="MSXML2 v5.0"
                    session("XMLUse")=".5.0"
                case ".6.0"
                    strXMLUse="MSXML2 v6.0"
                    session("XMLUse")=".6.0"
                case else
                    strXMLUse="MSXML2 default"
                    session("XMLUse")=""
            end select
            %>	
              <p><b>Current XML Parser</b>: the store is currently using the <b><%=strXMLUse%></b> XML parser</p>
              <p>&nbsp;</p>
              <p><b>Available XML Parsers on Server</b>
              <% dim xml, XMLAvailable, XMLUse, XML_checked, XML_Err_reason, XML_Err_reason_2, XML3_checked, XML3_Err_reason, XML3_Err_reason_2, XML4_checked, XML4_Err_reason, XML4_Err_reason_2
									xml = "<?xml version=""1.0"" encoding=""UTF-16""?><cjb></cjb>"
									XMLAvailable=0
									XMLUse=0
									XML_checked = ""
									XML_Err_reason = "<b>INSTALLED</b>"
									XML_Err_reason_2 = ""
									XML26_checked = ""
									XML26_Err_reason = "<b>INSTALLED</b>"
									XML26_Err_reason_2 = ""
									XML3_checked = ""
									XML3_Err_reason = "<b>INSTALLED</b>"
									XML3_Err_reason_2 = ""
									XML4_checked = ""
									XML4_Err_reason = "<b>INSTALLED</b>"
									XML4_Err_reason_2 = ""
									XML5_checked = ""
									XML5_Err_reason = "<b>INSTALLED</b>"
									XML5_Err_reason_2 = ""
									XML6checked = ""
									XML6_Err_reason = "<b>INSTALLED</b>"
									XML6_Err_reason_2 = ""

									testURL="https://www.ups.com/ups.app/xml/Rate" 
									
									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument")
									x.async = false 
									if x.loadXML(xml) then
									 XML_checked="checked"
									end if
									if err.number<>0 then
									 XML_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
									 XML_checked=""
									 err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML3_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=1
											end if
										end if
									end if
									
									dim intReqXML
									intReqXML=0
									
									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument.2.6")
									x.async = false 
									if x.loadXML(xml) then
										XML26_checked="checked"
									end if
									if err.number<>0 then
										XML26_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
										XML26_checked=""
										err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp.2.6")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML26_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML26_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=26
												intReqXML=1
											end if
										end if
									end if

									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument.3.0")
									x.async = false 
									if x.loadXML(xml) then
										XML3_checked="checked"
									end if
									if err.number<>0 then
										XML3_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
										XML3_checked=""
										err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp.3.0")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML3_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML3_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=3
												intReqXML=1
											end if
										end if
									end if
									
									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument.4.0")
									x.async = false 
									if x.loadXML(xml) then
										XML4_checked="checked"
									end if
									if err.number<>0 then
										XML4_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
										XML4_checked=""
										err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp.4.0")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML4_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML4_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=4
												intReqXML=1
											end if
										end if
									end if
									
									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument.5.0")
									x.async = false 
									if x.loadXML(xml) then
										XML5_checked="checked"
									end if
									if err.number<>0 then
										XML5_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
										XML5_checked=""
										err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp.5.0")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML5_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML5_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=5
												intReqXML=1
											end if
										end if
									end if
									
									err.number=0
									Set x = server.CreateObject("Msxml2.DOMDocument.6.0")
									x.async = false 
									if x.loadXML(xml) then
										XML6_checked="checked"
									end if
									if err.number<>0 then
										XML6_Err_reason="NOT INSTALLED"  'Error - <b>"&err.description&"</b>"
										XML6_checked=""
										err.number=0
									else
										Set srvUPSXmlHttp = server.createobject("Msxml2.serverXmlHttp.6.0")
										srvUPSXmlHttp.open "POST", testURL, false
										if err.number<>0 then
											XML6_Err_reason_2="<b>"&err.description&"</b>"
											err.number=0
										else
											srvUPSXmlHttp.send(xml)
											if err.number<>0 then
												XML6_Err_reason_2="<b>"&err.description&"</b>"
												err.number=0
											else
												XMLAvailable=1
												XMLUse=6
												intReqXML=1
											end if
										end if
									end if															
									%>               
              
              <table width="100%">
                <tr>
                  <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value="" <% if scXML="" AND XML_Err_reason_2="" then %>checked<%end if%><% if NOT (XML_Err_reason="<b>INSTALLED</b>" AND XML_Err_reason_2="") then %>disabled=1<%end if%>></td> 
                  <td width="18%" nowrap>MSXML Parser</td>
                  <td width="77%" nowrap id="msxmlreason"><%=XML_Err_reason%><% if XML_Err_reason="<b>INSTALLED</b>" AND XML_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML_Err_reason_2%><% end if %></td>
                </tr>
                <tr>
                    <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value=".2.6"  <% if scXML=".2.6" AND XML26_Err_reason_2="" then %>checked<%end if%><% if NOT (XML26_Err_reason="<b>INSTALLED</b>" AND XML26_Err_reason_2="") then %>disabled=1<%end if%>></td> 
                  <td width="18%" nowrap>MSXML2 v2.6 Parser</td>
                  <td id="XML26reason"><%=XML26_Err_reason%><% if XML26_Err_reason="<b>INSTALLED</b>" AND XML26_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML26_Err_reason_2%><% end if %></td>
								</tr>
                <tr>
                    <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value=".3.0" <% if scXML=".3.0" AND XML3_Err_reason_2="" then %>checked<%end if%><% if NOT (XML3_Err_reason="<b>INSTALLED</b>" AND XML3_Err_reason_2="") then %>disabled=1<%end if%>></td> 
                  <td width="18%" nowrap>MSXML2 v3.0 Parser</td>
                  <td id="XML3reason"><%=XML3_Err_reason%><% if XML3_Err_reason="<b>INSTALLED</b>" AND XML3_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML3_Err_reason_2%><% end if %></td>
								</tr>
								<tr>
									  <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value=".4.0" <% if scXML=".4.0" AND XML4_Err_reason_2="" then %>checked<%end if%><% if NOT (XML4_Err_reason="<b>INSTALLED</b>" AND XML4_Err_reason_2="") then %>disabled=1<%end if%>></td>
									<td width="18%" nowrap>MSXML2 v4.0 Parser</td>
									<td id="XML4reason"><%=XML4_Err_reason%><% if XML4_Err_reason="<b>INSTALLED</b>" AND XML4_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML4_Err_reason_2%><% end if %></td>
								</tr>
								<tr>
									  <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value=".5.0" <% if scXML=".5.0" AND XML5_Err_reason_2="" then %>checked<%end if%><% if NOT (XML5_Err_reason="<b>INSTALLED</b>" AND XML5_Err_reason_2="") then %>disabled=1<%end if%>></td>
									<td width="18%" nowrap>MSXML2 v5.0 Parser</td>
									<td id="XML5reason"><%=XML5_Err_reason%><% if XML5_Err_reason="<b>INSTALLED</b>" AND XML5_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML5_Err_reason_2%><% end if %></td>
								</tr>
								<tr>
									  <td width="5%" valign="top" nowrap><input type="radio" name="ChangeXMLParser" value=".6.0" <% if scXML=".6.0" AND XML6_Err_reason_2="" then %>checked<%end if%><% if NOT (XML6_Err_reason="<b>INSTALLED</b>" AND XML6_Err_reason_2="") then %>disabled=1<%end if%>></td>
									<td width="18%" nowrap>MSXML2 v6.0 Parser</td>
									<td id="XML6reason"><%=XML6_Err_reason%><% if XML6_Err_reason="<b>INSTALLED</b>" AND XML6_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML6_Err_reason_2%><% end if %></td>
								</tr>
								</table>

                <% if XMLAvailable=1 then %>
                <table>
				<% if intReqXML=0 then %>
					<tr> 
						<td width="4%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
						<td width="96%">The XML Parser version 3.0 is not installed or has returned errors while trying to connect. This will cause problems if you decide to use UPS as a dynamic shipping provider. Contact your hosting provider and ask them to install or reinstall the XML Parser version 3.0.</td>
					</tr>
				<% else %>
					<tr> 
						<td width="4%" valign="top"><img src="images/pc_checkmark_sm.gif" width="16" height="16"></td>
						<td width="96%">The required XML Parsers are installed.</td>
					</tr>
					<% 
					'// Do not recommend 6 if a lower (fully supported) version is available. Google Checkout doesnt work with 6.
					if XMLUse=6 then
						if XML5_Err_reason="<b>INSTALLED</b>" AND XML5_Err_reason_2="" then
							XMLUse=5
						end if
						if XML4_Err_reason="<b>INSTALLED</b>" AND XML4_Err_reason_2="" then
							XMLUse=4
						end if
						if XML3_Err_reason="<b>INSTALLED</b>" AND XML3_Err_reason_2="" then
							XMLUse=3
						end if
					end if
					select case XMLUse
						case 1
							strXMLUse="We recommend you use the MSXML2 default parser."
						case 26
							strXMLUse="We recommend you use MSXML2 v2.6 parser."
						case 3
							strXMLUse="We recommend you use MSXML2 v3.0 parser."
						case 4
							strXMLUse="We recommend you use MSXML2 v4.0 parser."
						case 5
							strXMLUse="We recommend you use MSXML2 v5.0 parser."
						case 6
							strXMLUse="We recommend you use MSXML2 v6.0 parser."
					end select
					%><br>

					<tr>
					  <td colspan="2" valign="top"><br>
					    <p>To update your XML Parser, select the parser of your choice above and click the &quot;Update&quot; button below.</p>
					    <p><%=strXMLUse%></p>
					    <p>&nbsp;</p>
					    <p>
					      <input type="submit" name="Submit" value="Update" class="submit2">
</p></td>
					  </tr>
				<% end if %>
                </table>
                <% else %>
                <table>
                  <tr> 
                    <td width="5%" valign="top"><img src="images/pc_error_sm.gif" width="18" height="18"></td>
                    <td width="95%">Please contact your Web hosting company and ask them to install Microsoft's free XML Parser 3.0 SP2 or above.</td>
                  </tr>
                </table>
                <% end if %>
              <% end if %>
            <p align="center">
				<% if noConnDb=1 then %>
					There is a problem with your database connection string.<br /><br />
				<% end if %>
			</p>
			
      		<h2>Server Information</h2>		
			<table>
			<%
				Response.Write "<tr><td>SERVER NAME</td><td>=</td><td><b>" & request.servervariables("SERVER_NAME") & "</b></td></tr>" & vbcrlf
				Response.Write "<tr><td>SERVER PORT</td><td>=</td><td><b>" & request.servervariables("SERVER_PORT") & "</b></td></tr>" & vbcrlf
				Response.Write "<tr><td>SERVER PORT SECURE</td><td>=</td><td><b>" & request.servervariables("SERVER_PORT_SECURE") & "</b></td></tr>" & vbcrlf
				Response.Write "<tr><td>SERVER PROTOCOL</td><td>=</td><td><b>" & request.servervariables("SERVER_PROTOCOL") & "</b></td></tr>" & vbcrlf
				Response.Write "<tr><td>SERVER SOFTWARE</td><td>=</td><td><b>" & request.servervariables("SERVER_SOFTWARE") & "</b></td></tr>" & vbcrlf
				'Response.Write "<tr><td>IP ADDRESS</td><td>=</td><td><b>" & request.servervariables("REMOTE_ADDR") & "</b></td></tr>" & vbcrlf
				'Response.Write "<tr><td>BROWSER</td><td>=</td><td><b>" & request.servervariables("HTTP_USER_AGENT") & "</b></td></tr>" & vbcrlf
			%> 
			</table>
          </td>
     </tr>
	 <% '// End Web Server Readiness%>

	 <% '// Start Payment Gateway Components%>
	 <tr>
	   	<td>
		   <h2>Payment Gateway Components</h2>
			<% if actGW=1 then %>
				<p><strong>The following real-time payment options have been configured:</strong>
				  <a href="javascript:wingwaycheckobj('GWayComcheck.asp')"> ( Payment Gateway Component Check... ) </a>
				</p>

				<table>
					<!-- No Components/Classes -->
					<% if gw2Checkout="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">2Checkout is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwGestPay="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Banca Sella GestPay is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwbofa="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Bank of America eStore Solution is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwBluePay="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">BluePay&#8482; is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwCBN="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">ChecksByNet is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwconcord="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Concord EFSnet is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwIntSecure="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">InternetSecure is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwit="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">ITransact is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwPaymentExpress="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Payment Express - PX Post is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwprotx="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Sage Page (Protx) is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwtcs="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Trine CS is no longer supported</td>
					</tr>
					<% end if %>
					<% if gwtclink="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">TrustCommerce (TCLink) is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwvpfl="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PayPal Payflow Link is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwklix="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">viaKLIX is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					<% if gwwp="1" then %>
					<tr> 
						<td height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">WorldPay is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					
					
					
					<!-- XML Parser -->
					<%
					'"FedEx"
					'"UPS"
					'"USPS"
					'"CanadaPost"
					%>

					<% strErr = "Msxml2.serverXMLHTTP" & scXML %>

					<% if gwa="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">AuthorizeNet is <b>ENABLED</b> 
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwecho="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Echo is <b>ENABLED</b> 
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwEPN="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">eProcessing Network is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwEway="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">eWay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwHSBC="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">HSBC is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwMoneris="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Moneris is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwNetBilling="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Netbilling is <b>ENABLED</b> 
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwNETOne="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Net1 Payment Services is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwPaymentech="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Paymentech is <b>ENABLED</b> 
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwSecPay="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">SECPay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwTripleDeal="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">TripleDeal is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					<% if gwEmerchant="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Fasthosts eMerchant is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					
				   <% if gwCP="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">ChronoPay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
				<% if gwPxPay="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Payment Express &reg; PX Pay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gweMoney="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">ETS - EMoney<sup>TM</sup> is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwOgone="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Ogone is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwVirtualMerchant="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Virtual Merchant is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwBeanStream="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">BeanStream is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					
					<% if gwGlobalPay="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Global Pay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwOmega="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Omega Pay is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwDowCommmerce="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Dow Commmerce is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwTotalWeb="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">TotalWeb is <b>ENABLED</b>  
						<% If XMLAvailable=0 Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>


					<!-- 3rd Party Installed Components -->
					<% if gwlp="1" then
						'//Check if basic or API
						query="SELECT lp_yourpay FROM linkpoint;"
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=connTemp.execute(query)
						VAR_LPType = rs("lp_yourpay")
						VAR_LPEnabled = "Basic"
						If VAR_LPType = "API" Then
							VAR_LPEnabled = "API"
						End If
						If VAR_LPType = "YES" Then
							VAR_LPEnabled = "YourPay"
						End If
						'//End Check
						%>
                        <tr> 
                            <td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
                            <td height="21">Linkpoint <%=VAR_LPEnabled%> is <b>ENABLED</b>.<br>
                            <% If VAR_LPEnabled = "API" Then
								strErr = IsObjInstalled(9)
								If strErr <> "" Then %>
									The required components ("<b><%=strErr%></b>") are <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
								<% Else %>
									The required components are <b>INSTALLED</b>
								<% End If %>
                        	<% End If %></td>
                        </tr>
					<% end if %>                    
					
					<% if gwuep="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">USAePay is <b>ENABLED</b> 
						<% strErr = IsObjInstalled(0)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwACH="1" then %>
					<tr>
						<td height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">ACH Direct is <b>ENABLED</b>  
						<% strErr = IsObjInstalled(1)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwfast="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Fast Transact is <b>ENABLED</b>  
						<% strErr = IsObjInstalled(2)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwParaData="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">Paradata is <b>ENABLED</b>  
						<% strErr = IsObjInstalled(3)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwvpfp="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PayPal Payflow Pro is <b>ENABLED</b></td>
					</tr>
					<% end if %>
					
					<% if gwpsi="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PSIGate is <b>ENABLED</b>  
						<% strErr = IsObjInstalled(4)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required component is <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
					<% if gwcys="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">CyberSource is <b>ENABLED</b> 
						<% strErr = IsObjInstalled(5)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required components are <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>

					<% if gwpp="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21">PayPal is <b>ENABLED</b>  
						<%
						if gwppstandard="1" then
							%>
							&nbsp;-&nbsp;Since you are using PayPal Standard, no component needs to be installed on the server.
							<%
						elseif PayPalExp="1" then
							%>
							&nbsp;-&nbsp;You are using PayPal Express Checkout.
							<%
						else
							%>
							&nbsp;-&nbsp;You are using Website Payments Pro.
							<%
						end if
						%>
						</td>
					</tr>
					<% end if %>
					 <% if gwSkipJack="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21"> SkipJack is <b>ENABLED</b> 
						<% strErr = IsObjInstalled(9)
						If strErr <> "" Then %>
							and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						<% Else %>
							and the required components are <b>INSTALLED</b>
						<% End If %>
						</td>
					</tr>
					<% end if %>
					
				  <% if gwPayJunction="1" then %>
					<tr> 
						<td width="6%" height="21"><b><img src="images/lighton.gif" width="18" height="22"></b></td>
						<td height="21"> Pay Junction is <b>ENABLED</b> 
					 <% 
					     strErr = IsObjInstalled(7)
						strErr2 = IsObjInstalled(8)
						if  strErr <> "" and  strErr2 <> "" then%>						
						 and the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
						 or the required component ("<b><%=strErr%></b>") is <b><span style='color: #FF0000'>NOT INSTALLED</span></b>
					 <% Else%>					
							and the required components are <b>INSTALLED</b>
					 <% End if %>
						</td>
					</tr>
					<% end if %>
				</table>
			<% else %>
				<p><strong>No real-time payment options have been configured</strong></p>
			<% end if %>
		</td>
	</tr>
	<% '// End Payment Gateway Components%>
	
	<% '// Start Product Weight Check %>
	<tr>
	   	<td>
			<h2>Product Weight Check</h2>
			<%  ' Do Product Weight Check
				err.clear
				err.number=0

				query="SELECT idProduct, sku, description FROM products "
				query = query & "where weight<1 and active=-1 and noshipping=0 AND pcprod_QtyToPound=0 "
				query = query & "ORDER BY idProduct;"	
				set rs2=Server.CreateObject("ADODB.Recordset")     
				rs2.PageSize=5

				rs2.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
				'// Page Count
				iProductsPageCount=rs2.PageCount
				if err.number <> 0 then
					set rs2=nothing
					call closedb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error in pcTSUtility: 3"&Err.Description) 
				end if
				prdStr=""
				idProduct=0
				wCount = 0
				if NOT rs2.eof then %>
                    <p><strong>Possible Shipping Issue:</strong> When products do not have a weight assigned, the store cannot properly calculate shipping charges for any shipping service that is based on the order weight. The following products do not have a weight assigned:</p>
					<p>
						<ul>

						<%	do while not rs2.eof and wCount < rs2.PageSize
							idProduct=rs2("idProduct")
							prdStr=rs2("description")&" ("&rs2("sku")&")"	%>
		
							<li><a href="FindProductType.asp?id=<%=idProduct%>"><%response.write prdStr%></a></li>
							<%
							wCount=wCount+1
							rs2.moveNext
						loop
						set rs2=nothing					
						%>
						<%  if iProductsPageCount > 1 then %>
							<li style="padding-top: 10px;"><a href="javascript:wingwaycheckobj('prdWC_popup.asp')">View complete list...</a></li>
						<%	end if %>
                    	</ul>
					</p>
				<% else %>
                <table>
                  <tr> 
                    <td width="5%" valign="top"><img src="images/pc_checkmark_sm.gif" width="18" height="18"></td>
                    <td width="95%"><strong>All shipping products have a weight assigned.</strong></td>
                  </tr>
                </table>
			<% end if %>
		</td>
	</tr>
	<% '// End Product Weight Check %>
	
	<% '// Start SSL Check%>
	<tr>
		<td>
			<h2>Secure Socket Layer (SSL) Check</h2>
			<ul>
			<%
			'Response.Write "<li>HTTPS on your server is turned <b>" & ucase(request.servervariables("HTTPS")) & " </b></li>" & vbcrlf
			if scSSL="1" then 
				Response.Write "<li>SSL for your Store is <b>ENABLED</b></li>" & vbcrlf
			else
				Response.Write "<li>SSL for your Store is <b>DISABLED</b></li>" & vbcrlf
			end if

			if scSSLUrl<>"" then
				if inStr(scSSLUrl,"http://") then
					Response.Write "<li>Your Store's SSL URL <b>" & scSSLURL & " </b> does not use the <strong>HTTPS</strong> protocol. Make sure you have entered it correctly. It should start with &quot;https://&quot;." & vbcrlf
				else
					session("SSLCheck")="SSL Check Successful."
					dim strRedirectSSL
					Response.Write "<li>Your Store's SSL URL is <b>" & scSSLURL & " </b>" & vbcrlf		
					if scSSL="1" AND scIntSSLPage="1" then
						strRedirectSSL=replace((scSslURL&"/"&scPcFolder&"/"&scAdminFolderName&"/pcSSLCheck.asp"),"//","/")
						strRedirectSSL=replace(strRedirectSSL,"https:/","https://")
						strRedirectSSL=replace(strRedirectSSL,"http:/","http://")
						Response.Write " - <a href=""javascript:;;"" onclick=""javascript:wincomcheckobj('"&strRedirectSSL&"')"">Run SSL Check...</a>" & vbcrlf
					end if
					Response.Write "</li>" & vbcrlf
				end if
			else
				Response.Write "<li>Your Store's SSL URL is <b>NOT DEFINED</b></li>" & vbcrlf			
			end if

			%>
			</ul>
		</td>
	</tr>
	<% '// End SSL heck %>
	
	<% '// Start Record Count %>
	<tr>
		<td>
			<h2>Product Database Information</h2>
				<ul>
					<%
					query="SELECT idProduct FROM products ORDER BY idProduct;"	
					set rs3=Server.CreateObject("ADODB.Recordset")     
					rs3.Open query, conntemp
					if err.number <> 0 then
						set rs3=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in pcTSUtility: 3"&Err.Description) 
					end if
					productCount = 0
					if NOT rs3.eof then					
						do while NOT rs3.eof
							productCount = productCount + 1
						rs3.movenext
						loop
					end if
					set rs3=nothing
					%>
					<li>The total number of <strong>products</strong> is <strong><%=productCount%></strong>.
					<br><i>This number represents the total record count from the &quot;products&quot; table. It includes inactive and <a href="http://www.earlyimpact.com/faqs/afmviewfaq.asp?faqid=379" target="_blank">deleted products</a> and can be useful when troubleshooting performance issues</i></li>
					<%
					query="SELECT idCategory FROM categories ORDER BY idCategory;"	
					set rs3=Server.CreateObject("ADODB.Recordset")
					rs3.Open query, conntemp
					if err.number <> 0 then
						set rs3=nothing
						call closedb()
						response.redirect "techErr.asp?error="& Server.Urlencode("Error in pcTSUtility: 3"&Err.Description) 
					end if
					categoryCount = 0
					if NOT rs3.eof then					
						do while NOT rs3.eof
							categoryCount = categoryCount + 1
						rs3.movenext
						loop
					end if
					set rs3=nothing
					%>
					<li>The total number of <strong>categories</strong> is <strong><%=categoryCount%></strong>.
					<br><i>This number represents the total record count from the &quot;categories&quot; table.</i></li>
				</ul>
		</td>
	</tr>
	<% '// End Record Count %>

</table></form>
<% call closedb() %>
<!--#include file="AdminFooter.asp"-->