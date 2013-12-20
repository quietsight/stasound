<!--#include file="pcSetupHeader.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/secureadminfolder.asp"-->
<!--#include file="ppdstatus.inc"-->
<% 
'// Clear any sessions for the connection string
session("tmpConnectionString")=""
%>
        <form name="form1" action="step2_results.asp" method="post" class="pcForms">
                <h1>Step 2: Web server readiness utility test results</h1>
                <p>Here are the results of our tests:</p>
                <p><strong>Parent Paths</strong></p>
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
                            <td width="4%"><img src="pc_checkmark_sm.gif" width="16" height="16"></td>
                            <td width="96%">
                                <% if PPD="1" then
                                    response.write "Not applicable. No action needs to be taken."
                                else
                                    response.write "Parent paths are enabled. No action needs to be taken."
                                end if %>
                            </td>
                        </tr>
                    </table>
				<% else
					allstop=1 %>
					<table>
						<tr> 
							<td width="4%"><img src="pc_error_sm.gif" width="18" height="18"></td>
							<td width="96%">Parent Paths are disabled on this server. Please use the version of ProductCart that is included in the &quot;Parent Paths Disabled&quot; folder within the ProductCart installation files, or ask your Web hosting company to &quot;enable parent paths&quot; (<a href="http://wiki.earlyimpact.com/productcart/install#parent_paths_enabled_and_disabled" target="_blank">more information</a>).</td>
						</tr>
					</table>
				<% end if %>
				<hr>
				<% if allstop=0 then %>
                    <p><strong>Folder permissions</strong></p>
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
						
					if PPD=1 then
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
					
					if PPD=1 then
						PageName="/"&scPcFolder&"/"&scAdminFolderName&"/"&replace(Date,"/","")&".txt"
					else			
						PageName="../"&scAdminFolderName&"/"&replace(Date,"/","")&".txt"
					end if
					
					findit=Server.MapPath(PageName)
					
					Set objCPFSO = Server.CreateObject ("Scripting.FileSystemObject")
					Set objCPFile = objCPFSO.OpenTextFile (findit, 8, True, 0)
					objCPFile.WriteLine "test done"
					objCPFile.Close
					set objCPFSO = nothing
					set objCPFile = nothing

					if err.number<>0 then
						errwrite_pcadmin=1
						errpermissions=1
						Err.number=0
					end if
					
					if errpermissions=0 then %>
                        <table>
                            <tr> 
                                <td width="4%"><img src="pc_checkmark_sm.gif" width="16" height="16"></td>
                                <td width="96%">Folder permissions have been assigned correctly.</td>
                            </tr>
                        </table> 
                    <% else %>
						<% if errwrite_others=1 or errwrite_includes=1 or errwrite_pcadmin=1 then %> 
                            <table>
                              <tr> 
                                <td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
                                <td width="95%"><font color="#CC3950">You need to assign 'read/write' permissions to the 'productcart' folder and all of its subfolders.</font></td>
                              </tr>
                            </table>
						<% end if
						if errdelete_includes=1 then %>
							<table>
								<tr> 
								<td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
								<td width="95%"><font color="#CC3950">You need to assign 'read/write/delete' permissions to the 'productcart/includes' folder and all of its subfolders.</font></td>
								</tr>
							</table>
						<% end if
						if errwrite_pcadmin=1 then %>
							<table>
								<tr> 
								<td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
								<td width="95%"><font color="#CC3950">You need to assign 'read/write/delete' permissions to the 'productcart/<%=scAdminFolderName%>' folder and all of its subfolders.</font></td>
								</tr>
							</table>
						<% end if
                    end if %>
                    <hr>
                    <p><strong>Database connection string</strong></p>
                    <% dim noConnDb, pcv_InvalidChar
                    noConnDb=0
                    scDSN=request.form("dbConn")
                    pcv_InvalidChar=""
                    if instr(scDSN, "#") then
                        pcv_InvalidChar="#"
                    end if
                    set connTemp=server.createobject("adodb.connection")
                    err.number=0
                    connTemp.Open scDSN  
                    if err.number = 0 then
						'//Check for invalid character
						if pcv_InvalidChar="" then
							'// Check if database has a ProductCart table [admins]
							query="SELECT * FROM admins;"
							set rs=server.CreateObject("ADODB.RecordSet")
							set rs=connTemp.execute(query)
							if err.number=0 then %>
								<table>
								  <tr> 
									<td width="4%" valign="top"><img src="pc_checkmark_sm.gif" width="16" height="16"></td>
									<td width="96%">The connection string is valid. Use this connection string on the next window.</td>
								  </tr>
								</table>
							<% else
								noConnDb=2 %>
                                <table>
                                  <tr> 
                                    <td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">The connection string is valid, it appears that you have not yet populated your database. If you are using SQL, please run the required SQL Script to create the necessary tables in your database to use ProductCart.</font></td>
                                  </tr>
                                </table>
                            <% end if
						else 
							noConnDb=1 %>
							<table>
							  <tr> 
								<td width="4%" valign="top"><img src="pc_error_sm.gif" width="16" height="16"></td>
								<td width="96%">The connection string contains an invalid character "<%=pcv_InvalidChar%>" which is not supported by ProductCart.</td>
							  </tr>
							</table>
						<% end if
					else 
						noConnDb=1 %>
						<table>
						  <tr> 
							<td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
							<td width="95%"><font color="#CC3950">We were unable to connect to the database. <a href="javascript:win('connectionStrings.asp')">Click here</a> for <a href="javascript:win('connectionStrings.asp')">examples</a> of supported database connection strings.</font></td>
						  </tr>
						</table>
					<% end if %>
                    <hr>
                    <p><strong>Available E-mail Components</strong></p>
                    <%
					Dim theComponent(4)
					Dim theComponentName(4)
					
					' the components
					theComponent(0) = "CDO.Message"
					theComponent(1) = "CDONTS.NewMail"
					theComponent(2) = "Bamboo.SMTP"
					theComponent(3) = "SMTPsvg.Mailer"
					theComponent(4) = "JMail.SMTPMail"
					theComponent(5) = "Persits.MailSender"
					
					' the name of the components
					theComponentName(0) = "CDOSYS"
					theComponentName(1) = "CDONTS"
					theComponentName(2) = "Bamboo SMTP"
					theComponentName(3) = "ServerObjects ASPMail"
					theComponentName(4) = "JMail"
					theComponentName(5) = "Persits ASPMail"
					
					Function IsObjInstalled(strClassString)
						On Error Resume Next
						' initialize default values
						IsObjInstalled = False
						Err = 0
						' testing code
						Dim xTestObj
						Set xTestObj = Server.CreateObject(strClassString)
						If 0 = Err Then IsObjInstalled = True
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
                            <td width="34%" nowrap style="color:#336699;"><%= theComponentName(i)%></td>
                            <td width="5%">
								<% If Not IsObjInstalled(theComponent(i)) Then
									emailcom_checked=""
								else
									emailcom_checked="checked"
									emailComCnt=emailComCnt+1
								end if %><input type="checkbox" name="emailcom" value="ON" disabled <%=emailcom_checked%> class="clearBorder"></td>
                                <td width="62%" nowrap style="color:#336699;"> 
                                  <% If Not IsObjInstalled(theComponent(i)) Then %>
                                   Not Installed 
                                    <% Else %>
                                    Installed
                                    <% End If %>
                                </td>
                            </tr>
						  <% Next %>
                        </table>
						<% if emailComCnt>0 then %>
                            <table>
                              <tr> 
                                <td width="4%" valign="top"><img src="pc_checkmark_sm.gif" width="16" height="16"></font></td>
                                <td width="96%">Compatible e-mail components are available on this server.</td>
                              </tr>
                            </table>
						<% else %>
                            <table>
                              <tr> 
                                <td width="4%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
                                <td width="96%"><font color="#CC3950">We are unable to detect a supported E-mail component on this server. ProductCart supports the components listed above. Please contact your Web hosting company and ask them to install an e-mail component for you.</font></td>
                              </tr>
                            </table>
						<% end if %>
                        <hr>
                        <p><strong>XML Parser</strong></p>
							<% dim xml, XMLAvailable, XMLUse, XML_checked, XML_Err_reason, XML_Err_reason_2, XML3_checked, XML3_Err_reason, XML3_Err_reason_2, XML4_checked, XML4_Err_reason, XML4_Err_reason_2
							xml = "<?xml version=""1.0"" encoding=""UTF-16""?><cjb></cjb>"
							XMLAvailable=0
							XMLUse=0
							XML_checked = ""
							XML_Err_reason = "Installed"
							XML_Err_reason_2 = ""
							XML26_checked = ""
							XML26_Err_reason = "Installed"
							XML26_Err_reason_2 = ""
							XML3_checked = ""
							XML3_Err_reason = "Installed"
							XML3_Err_reason_2 = ""
							XML4_checked = ""
							XML4_Err_reason = "Installed"
							XML4_Err_reason_2 = ""
							XML5_checked = ""
							XML5_Err_reason = "Installed"
							XML5_Err_reason_2 = ""
							XML6_checked = ""
							XML6_Err_reason = "Installed"
							XML6_Err_reason_2 = ""

							testURL="https://www.ups.com/ups.app/xml/Rate"
							 
							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument")
							x.async = false 
							if x.loadXML(xml) then
							 XML_checked="checked"
							end if
							set x=nothing
							
							if err.number<>0 then
								XML_Err_reason=err.description
								XML_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=1
									end if
								end if
								set srvXmlHttp=nothing
							end if
							
							dim intReqXML
							intReqXML=0
							
							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument.2.6")
							x.async = false 
							if x.loadXML(xml) then
								XML26_checked="checked"
							end if
							set x=nothing
							if err.number<>0 then
								XML26_Err_reason=err.description
								XML26_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.2.6")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML26_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML26_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=2.6
										intReqXML=1
									end if
								end if
								set srvXmlHttp=nothing
							end if

							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument.3.0")
							x.async = false 
							if x.loadXML(xml) then
								XML3_checked="checked"
							end if
							set x=nothing
							if err.number<>0 then
								XML3_Err_reason=err.description
								XML3_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.3.0")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML3_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML3_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=3
										intReqXML=1
										XML3Trump="Y"
									end if
								end if
								set srvXmlHttp=nothing
							end if

							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument.4.0")
							x.async = false 
							if x.loadXML(xml) then
								XML4_checked="checked"
							end if
							set x=nothing
							if err.number<>0 then
								XML4_Err_reason=err.description
								XML4_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.4.0")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML4_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML4_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=4
										intReqXML=1
									end if
								end if
								set srvXmlHttp=nothing
							end if

							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument.5.0")
							x.async = false 
							if x.loadXML(xml) then
								XML5_checked="checked"
							end if
							set x=nothing
							if err.number<>0 then
								XML5_Err_reason=err.description
								XML5_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.5.0")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML5_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML5_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=5
										intReqXML=1
									end if
								end if
								set srvXmlHttp=nothing
							end if

							err.clear
							Set x = server.CreateObject("Msxml2.DOMDocument.6.0")
							x.async = false 
							if x.loadXML(xml) then
								XML6_checked="checked"
							end if
							set x=nothing
							if err.number<>0 then
								XML6_Err_reason=err.description
								XML6_checked=""
								err.clear
							else
								Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp.6.0")
								srvXmlHttp.open "POST", testURL, false
								if err.number<>0 then
									XML6_Err_reason_2=err.description
									err.clear
								else
									srvXmlHttp.send(xml)
									if err.number<>0 then
										XML6_Err_reason_2=err.description
										err.clear
									else
										XMLAvailable=1
										XMLUse=6
										intReqXML=1
									end if
								end if
								set srvXmlHttp=nothing
							end if
							
							if XMLAvailable=1 AND XML3Trump="Y" then
								XMLUse=3
								intReqXML=1
							end if
							%>               
                            <table>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML Parser</td>
                                  <td width="6%"><input type="checkbox" name="msxml" value="ON" disabled <%=XML_checked%> class="clearBorder"></td>
                                  <td nowrap id="msxmlreason" style="color:#336699;"><%=XML_Err_reason%><% if XML_Err_reason="Installed" AND XML_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML_Err_reason_2%><% end if %></td>
                              </tr>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML2 v2.6 Parser</td>
                                                    <td width="6%"><input type="checkbox" name="XML26" value="ON" disabled <%=XML26_checked%> class="clearBorder"></td>
                                                    <td id="XML26reason" style="color:#336699;"><%=XML26_Err_reason%><% if XML26_Err_reason="Installed" AND XML26_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML26_Err_reason_2%><% end if %></td>
                                                </tr>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML2 v3.0 Parser</td>
                                                    <td width="6%"><input type="checkbox" name="XML3" value="ON" disabled <%=XML3_checked%> class="clearBorder"></td>
                                                    <td id="XML3reason" style="color:#336699;"><%=XML3_Err_reason%><% if XML3_Err_reason="Installed" AND XML3_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML3_Err_reason_2%><% end if %></td>
                                                </tr>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML2 v4.0 Parser</td>
                                                    <td width="6%"><input type="checkbox" name="XML4" value="ON" disabled <%=XML4_checked%> class="clearBorder"></td>
                                                    <td id="XML4reason" style="color:#336699;"><%=XML4_Err_reason%><% if XML4_Err_reason="Installed" AND XML4_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML4_Err_reason_2%><% end if %></td>
                                                </tr>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML2 v5.0 Parser</td>
                                                    <td width="6%"><input type="checkbox" name="XML5" value="ON" disabled <%=XML5_checked%> class="clearBorder"></td>
                                                    <td id="XML5reason" style="color:#336699;"><%=XML5_Err_reason%><% if XML5_Err_reason="Installed" AND XML5_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML5_Err_reason_2%><% end if %></td>
                                                </tr>
                              <tr> 
                                <td width="29%" nowrap style="color:#336699;">MSXML2 v6.0 Parser</td>
                                                    <td width="6%"><input type="checkbox" name="XML6" value="ON" disabled <%=XML6_checked%> class="clearBorder"></td>
                                                    <td id="XML6reason" style="color:#336699;"><%=XML6_Err_reason%><% if XML6_Err_reason="Installed" AND XML6_Err_reason_2<>"" then %>&nbsp;with Errors&nbsp;-&nbsp;<%=XML6_Err_reason_2%><% end if %></td>
                                                </tr>
                            </table>
							<% if XMLAvailable=1 then %>
                                <table>
									<% if intReqXML=0 then %>
										<tr> 
											<td width="4%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
											<td width="96%"><font color="#CC3950">The XML Parser version 3.0 is not installed or has returned errors while trying to connect. This will cause problems if you decide to use UPS as a dynamic shipping provider. Contact your hosting provider and ask them to install or reinstall the XML Parser version 3.0.</font></td>
										</tr>
									<% else %>
										<tr> 
                                            <td width="4%" valign="top"><img src="pc_checkmark_sm.gif" width="16" height="16"></font></td>
                                            <td width="96%">The required XML Parsers are installed.</td>
										</tr>
										<% 
										select case XMLUse
											case 1
												strXMLUse="ProductCart will use the MSXML2 default parser."
												session("XMLUse")=""
											case 26
												strXMLUse="ProductCart will use the MSXML2 v2.6 parser."
												session("XMLUse")=".2.6"
											case 3
												strXMLUse="ProductCart will use MSXML2 v3.0 parser."
												session("XMLUse")=".3.0"
											case 4
												strXMLUse="ProductCart will use MSXML2 v4.0 parser."
												session("XMLUse")=".4.0"
											case 5
												strXMLUse="ProductCart will use MSXML2 v5.0 parser."
												session("XMLUse")=".5.0"
											case 6
												strXMLUse="ProductCart will use MSXML2 v6.0 parser."
												session("XMLUse")=".6.0"
										end select
										%>
                                        <tr> 
                                            <td valign="top">&nbsp;</td>
                                            <td><%=strXMLUse%></td>
                                        </tr>
									<% end if %>
                                </table>
							<% else %>
                                <table>
                                  <tr> 
                                    <td width="5%" valign="top"><img src="pc_error_sm.gif" width="18" height="18"></td>
                                    <td width="95%"><font color="#CC3950">Please contact your Web hosting company and ask them to install Microsoft's free XML Parser 3.0 SP2 or above.</font></td>
                                  </tr>
                                </table>
							<% end if %>
						<% end if %>
						<hr>
                        <p align="center">
						<% if noConnDb=1 or noConnDb=2 then 
							if noConnDb=1 then %>
								<font color="#CC3950">Please click the "Back" button and update your database connection string.</font>
                           <% else %>
								<font color="#CC3950">Please create your database by running the required SQL Script against your SQL Database.</font>
                           <%  end if %>    
                                <br /><br />
						<% else 
							'// Put the connection string in Session to grab in next screen
							session("tmpConnectionString")=scDSN 
							%>
            				<input name="next" type="button" value="Proceed to Step 3" onClick="location.href='step3.asp'" class="submit2">
						<% end if %>
						<input name="back" type="button" value="Back" onClick="location.href='step2.asp'" class="ibtnGrey">
						</p>
                </form>
<!--#include file="pcSetupFooter.asp"-->