<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/taxsettings.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp" --> 
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<%
taxLoc=0
Dim mySQL, conntemp
call openDb()
Dim pStateCode, pCountryCode, pZip, pidPayment
pStateCode=request.QueryString("S")
pCountryCode=request.QueryString("C")
pZip=request.QueryString("Z")

'Trim PostalCode
if len(pZip)>5 then
	pZip=left(pZip,5)
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head><title>Tax Rate</title>
<link href="pcv4_ControlPanel.css" rel="stylesheet" type="text/css"></head>
<body style="background-image: none;">
	<table class="pcCPMainTable" style="width:450px;">
		<tr>
			<td>
				<table class="pcCPcontent">
					<tr> 
						<td width="71%"><b>Tax Rate</b></td>
						<td width="29%"></td>
					</tr>
							<%
							If ptaxfile=1 then
								' Let's now dynamically retrieve the current directory
								Dim sScriptDir
								if PPD="1" then
									filename="/"&scPcFolder&"/pc" & "/"
								else
									filename="../pc" & "/"
								end if
								sScriptDir=filename
				
								'get the file name
								dim Filename
								Filename=ptaxfilename
								Const ForReading = 1, ForWriting = 2, ForAppending = 8 
						
								dim FSO
								set FSO = Server.CreateObject("scripting.FileSystemObject") 
								
								'map the logincal path to the physical system path
								Dim Filepath
								Filepath=Server.MapPath(sScriptDir) & "\tax\" & Filename
								if NOT FSO.FileExists(Filepath) then
									response.write "<h3><i><font color=red>File " & imp_filename & " does not exist</font></i></h3>"
								end if
								
								If pCountryCode="US" then
									'see if state is a taxable one and then flag
									taxStateArray=split(ptaxRateState,", ")
									taxRateArray=split(ptaxRateDefault,", ")
									intTaxableState=0
									for i=0 to ubound(taxStateArray)-1
										if taxStateArray(i)=pStateCode then
											'flag
											intTaxableState=1
											intTaxRateDefault=taxRateArray(i)
										end if
									next
									
									dim f
									set f=FSO.GetFile(Filepath)
							
									'set ts = f.OpenAsTextStream(ForReading, -2) 
									Dim TextStream
									set TextStream=f.OpenAsTextStream(ForReading, -2) 
							
									zipCnt=0
									optionStr=""
									do While NOT TextStream.AtEndOfStream
										'line is found, now write it to new string
										Line=TextStream.readline
										'ignore first line
										if instr(ucase(Line), "ZIP") then
											iArray=split(Line,",")
											'loop to find correct array for each
											for q=0 to ubound(iArray)
												if iArray(q)="ZIP_CODE" then
													ZIP_NUM=q
													'response.write q&"<BR>"
												end if
												if iArray(q)="COUNTY_NAME" then
													COUNTY_NAME_NUM=q
													'response.write q&"<BR>"
												end if
												if iArray(q)="CITY_NAME" then
													CITY_NAME_NUM=q
													'response.write q&"<BR>"
												end if
												if iArray(q)="TOTAL_SALES_TAX" then
													TOTAL_SALES_TAX_NUM=q
													'response.write q&"<BR>"
												end if
												if iArray(q)="TAX_SHIPPING_ALONE" then
													TAX_SHIPPING_ALONE_NUM=q
													'response.write q&"<BR>"
												end if
												if iArray(q)="TAX_SHIPPING_AND_HANDLING_TOGETHER" then
													TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM=q
													'response.write q&"<BR>"
												end if
											next
											'response.end
										else
											'SEE IF MORE THEN ONE ZIP CODE EXIST
											if instr(Line, pZip) then
												zipCnt=zipCnt+1
												zArray=split(Line,",")
												ZIP_CODE=zArray(ZIP_NUM)
												if trim(ZIP_CODE)<>trim(pZip) then
												else
													COUNTY_NAME=zArray(COUNTY_NAME_NUM)
													CITY_NAME=zArray(CITY_NAME_NUM)
													TOTAL_SALES_TAX=zArray(TOTAL_SALES_TAX_NUM)
													TAX_SHIPPING_ALONE=zArray(TAX_SHIPPING_ALONE_NUM)
													TAX_SHIPPING_AND_HANDLING_TOGETHER=zArray(TAX_SHIPPING_AND_HANDLING_TOGETHER_NUM)
													'optionStr=optionStr&"<option value="""&TOTAL_SALES_TAX&""">"&CITY_NAME&" - "&COUNTY_NAME&"</option>"
													optionStr=optionStr&"<tr><FORM NAME='inputForm"&zipCnt&"' onSubmit='return setForm"&zipCnt&"();'><td width='71%'>"&CITY_NAME&" - "&COUNTY_NAME&": "&TOTAL_SALES_TAX*100&"%<INPUT NAME='inputField"&zipCnt&"' TYPE='hidden' VALUE='"&TOTAL_SALES_TAX*100&"'></td><td width='29%'><INPUT TYPE='SUBMIT' name='UPD' VALUE='Select Rate' onSubmit='return setForm();'></td></form></tr>"
												end if
											end if
										end if
									loop
							
									TextStream.Close
									If zipCnt=0 AND intTaxableState=1 then
										taxLoc=taxLoc+(intTaxRateDefault/100) 
									End If
				
								End if %>
								<%=optionStr%>
								<% response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf
								for i=1 to zipCnt
									response.write "function setForm"&i&"() {"&vbCrlf
									response.write "opener.document.EditOrder.calculateTaxPercentage.value = document.inputForm"&i&".inputField"&i&".value;"&vbCrlf
									response.write "self.close();"&vbCrlf
									response.write "return false;"&vbCrlf
									response.write "}"&vbCrlf
								next
								response.write "//--></SCRIPT>"&vbCrlf
							else
								'tax Per Place segment
								mySQL="SELECT taxLoc, taxDesc FROM taxLoc WHERE ((stateCode='" &pStateCode& "' AND stateCodeEq=-1) OR (stateCode IS NULL) OR (stateCode<>'" &pStateCode& "' AND stateCodeEq=0)) AND ((CountryCode='"&pCountryCode&"' AND CountryCodeEq=-1) OR (CountryCode IS NULL) OR (CountryCode<>'" &pCountryCode& "' AND CountryCodeEq=0)) AND ((zip='" &pZip& "' AND zipEq=-1) OR (zip IS NULL) OR (zip<>'" &pZip& "' AND zipEq=0));"
								set rs=conntemp.execute(mySQL)
								
								if err.number <> 0 then
									call closeDb()
									response.redirect "techErr.asp?error="&Server.Urlencode("Error in tax 1: "&err.description)
								end If
								
								if  rs.eof then 
								 ' there are no taxes defined for that zone
								 taxLoc=0
								end if
								taxCnt=0
								do until rs.eof 
									taxLoc=rs("taxLoc")
									taxDesc=rs("taxDesc")
									taxCnt=taxCnt+1 %>
									<tr>
									<FORM NAME="inputForm<%=taxCnt%>" onSubmit="return setForm<%=taxCnt%>();">
									<td width="71%"><strong><%=taxDesc%></strong>: <%=TaxLoc*100%>%<INPUT NAME="inputField<%=taxCnt%>" TYPE="hidden" VALUE="<%=TaxLoc*100%>"></td>
									<td width="29%"><INPUT TYPE="SUBMIT" name="UPD" VALUE="Select Rate" onSubmit="return setForm();"></td>
									</form>
									</tr>
									<% rs.movenext
								loop
								response.write "<SCRIPT LANGUAGE=JAVASCRIPT><!--"&vbCrlf&vbCrlf
								dim pcv_strTaxSelectType
								pcv_strTaxSelectType=Request("T")
								if pcv_strTaxSelectType<>"" then								
									for s=1 to taxCnt
										response.write "function setForm"&s&"() {"&vbCrlf
										response.write "opener.document.EditOrder.calculateTaxPercentage"&s&".value = document.inputForm"&s&".inputField"&s&".value;"&vbCrlf
										response.write "self.close();"&vbCrlf
										response.write "return false;"&vbCrlf
										response.write "}"&vbCrlf
									next
								else
									for s=1 to taxCnt
										response.write "function setForm"&s&"() {"&vbCrlf
										response.write "opener.document.EditOrder.calculateTaxPercentage.value = document.inputForm"&s&".inputField"&s&".value;"&vbCrlf
										response.write "self.close();"&vbCrlf
										response.write "return false;"&vbCrlf
										response.write "}"&vbCrlf
									next
								end if
								response.write "//--></SCRIPT>"&vbCrlf 
							End if
							%>
					</table>
				</td>
			</tr>
	</table>
</div>
</body>
</html>