<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="sds_LIv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"--> 
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<!--#include file="header.asp"-->

<% 
Dim connTemp,rs,query,pshipmentDetails

pcv_IdOrder=request("idorder")
if pcv_IdOrder="" then
	pcv_IdOrder=0
end if
pcv_PrdList=""
pcv_count=request("count")
if pcv_count="" then
	pcv_count=0
end if

if (pcv_IdOrder=0) or (pcv_count=0) then
	response.redirect "menu.asp"
end if

For i=1 to pcv_count
	if request("C" & i)="1" then
		pcv_PrdList=pcv_PrdList & request("IDPrd" & i) & ","
	end if
Next
	%>
<div id="pcMain">
	<table class="pcMainTable">
	<tr>
		<td valign="top">
		<table class="pcShowContent">
		<tr>
			<td colspan="6"><h1><%response.write dictLanguage.Item(Session("language")&"_sds_viewpast_1c")%> - <%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_1")%> <%=(scpre+int(pcv_IdOrder))%></h1></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		<tr>
			<td width="5%" align="center"><img border="0" src="images/step1.gif"></td>
			<td width="28%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_3")%></font></td>
			<td width="5%" align="center"><img border="0" src="images/step2a.gif"></td>
			<td width="28%" nowrap><b><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_4")%></b></td>
			<td width="5%" align="center"><img border="0" src="images/step3.gif"></td>
			<td width="29%" nowrap><font color="#A8A8A8"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_5")%></font></td>
		</tr>
		<tr>
			<td colspan="6" class="pcSpacer"></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>

	<Form name="form1" method="post" action="sds_ShipOrderWizard3.asp?action=add" class="pcForms">
		<table class="pcMainTable">
		<tr>
		<th colspan="2"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_14")%></th>
		</tr>
		<tr>
			<td colspan="2" class="pcSpacer"></td>
		</tr>
        
		<%
        ' START Shipment type
		call opendb()
		query="SELECT shipmentDetails FROM orders WHERE idOrder = " & pcv_IdOrder
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
		if not rs.EOF then
			pshipmentDetails=rs("shipmentDetails")
		end if
		set rs=nothing
		call closedb()
		if pshipmentDetails<>"" and not isNull(pshipmentDetails) then
        %>
            <tr>           
                <td colspan="2"><%response.write dictLanguage.Item(Session("language")&"_sds_custviewpastD_16")%>
                    <% 
                    Service=""
                    If pSRF="1" then
                        response.write ship_dictLanguage.Item(Session("language")&"_noShip_b")
                    else
                        'get shipping details...
                        shipping=split(pshipmentDetails,",")
                        if ubound(shipping)>1 then
                            if NOT isNumeric(trim(shipping(2))) then
                                varShip="0"
                                response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
                            else
                                Service=shipping(1)
                            end if
                            if len(Service)>0 then
                                response.write Service
                            End If
                        else
                            varShip="0"
                            response.write ship_dictLanguage.Item(Session("language")&"_noShip_a")
                        end if
                    end if
                    %>
                </td>
            </tr>
        <%
            if pOrdShipType=0 then
                pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_18")
            else
                pDisShipType=dictLanguage.Item(Session("language")&"_sds_custviewpastD_19")
            end if
            if varShip<>"0" then
        %>
                <tr> 
                    <td colspan="2"><%=dictLanguage.Item(Session("language")&"_sds_custviewpastD_17")%><%=pDisShipType%></td>
                </tr>
        <%
        	end if
		end if
        ' END Shipment Type
        %>

		<tr>
			<td colspan="2" class="pcSpacer"><hr></td>
		</tr>
        <tr>
        	<td colspan="2">
            	<table class="pcShowContent">
                    <tr>
                        <td nowrap="nowrap"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_15")%></td>
                        <td><input type="text" name="pcv_method" value="" size="30"></td>
                    </tr>
                    <tr>
                        <td nowrap="nowrap"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_16")%></td>
                        <td><input type="text" name="pcv_tracking" value="" size="30"></td>
                    </tr>
            <%
                        Dim varMonth, varDay, varYear
                        varMonth=Month(Date)
                        varDay=Day(Date)
                        varYear=Year(Date) 
                        dim dtInputStr
                        dtInputStr=(varMonth&"/"&varDay&"/"&varYear)
                        if scDateFrmt="DD/MM/YY" then
                            dtInputStr=(varDay&"/"&varMonth&"/"&varYear)
                        end if
            %>
                    <tr>
                        <td nowrap="nowrap"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_17")%></td>
                        <td><input type="text" name="pcv_shippedDate" value="<%=dtInputStr%>" size="30"> <i><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_17a")%></i></td>
                    </tr>
                    <tr>
                        <td valign="top"><%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_18")%></td>
                        <td><textarea name="pcv_AdmComments" size="40" rows="8" cols="40"></textarea></td>
                    </tr>
                    <tr>
                        <td colspan="2"><hr></td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td><input type="image" src="<%=rslayout("pcLO_finalShip")%>" name="submit1" value="<%response.write dictLanguage.Item(Session("language")&"_sds_shiporderwizard_19")%>" border="0" id="submit">&nbsp;<a href="javascript:history.back();"><img src="<%=rslayout("back")%>" border="0"></a>
                        <input type="hidden" name="PrdList" value="<%=pcv_PrdList%>">
                        <input type="hidden" name="idorder" value="<%=pcv_IdOrder%>">
                        <input type="hidden" name="count" value="<%=pcv_count%>">
                        </td>
                    </tr>
                  </table>
				</td>
			</tr>
		</table>
	</Form>

<table class="pcMainTable">
	<tr>
    	<td>&nbsp;</td>
    </tr>
    <tr>
        <td><div align="center"><a href="sds_MainMenu.asp"><%response.write(dictLanguage.Item(Session("language")&"_CustPref_1"))%></a> - <a href="sds_ViewPast.asp"><%response.write(dictLanguage.Item(Session("language")&"_sdsMain_3"))%></a></div></td>
    </tr>
</table>
</div>
<!--#include file="footer.asp"-->