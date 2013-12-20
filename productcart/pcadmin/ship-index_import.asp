<% pageTitle = "Import 'Order Shipped' Information - Upload/Locate Data File" %>
<% section = "orders" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
if request("action")="select" then
	if request("ways")="1" then
		response.redirect "ship-index_import1.asp"
	else
		if request("ways")="2" then
		response.redirect "ship-index_import2.asp"
		end if
	end if
end if
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
    <tr>  
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td  width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
        <td width="95%"><b><font color="#000000">Upload data file</font></b></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2.gif"></td>
        <td><font color="#A8A8A8">Map fields</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4.gif"></td>
        <td><font color="#A8A8A8">Import results</font></td>
    </tr>
</table>

<table class="pcCPcontent">
    <tr>  
        <td colspan="2">

		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>

    	<% if (request.querystring("nextstep")=1) then %>
        
            <form action="ship-step2.asp" method="post" class="pcForms">
                <table class="pcCPcontent">
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr>
                        <td>
                        <input type="hidden" name="append" value="1"> Existing order data will be updated if the Order ID imported in your import data file is an exact match with an existing Order ID.
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr>
                        <td>
                         <input type="submit" name="Go" value="Go to Step 2 >>" class="submit2">
                        </td>
                    </tr>
                </table>
            </form>
            
		<%else%>

            <form method="post" action="ship-index_import.asp?action=select">
            <table class="pcCPcontent">
                <tr>
                	<td colspan="2" class="pcCPspacer"></td>
                </tr>
            	<tr>
                	<td colspan="2"><h2>What happens after your file is successfully imported</h2>
            		<p>A series of tasks are performed by ProductCart: the order status may or may not be changed. The 'Order Shipped' message may or may not be sent. It depends on the information that you add to your import document. Please <a href="http://wiki.earlyimpact.com/productcart/orders_importing_shipping_info#what_happens_after_you_import" target="_blank">see the User Guide</a> for details.</p>
                    </td>
                </tr>
                <tr>
                	<td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr> 
                    <td colspan="2"><h2>Data file location</h2></td>
                </tr>
                <tr>
                <td width="5%" align="right"><input type=radio name="ways" value="1" checked></td>
                <td width="95%">Use the Import Wizard to upload the data file to the server</td>
                </tr>
                <tr>
                    <td align="right"><input type=radio name="ways" value="2"></td>
                    <td width="95%">The data file has been manually uploaded to the server</td>
                </tr>
                <tr>
                	<td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td colspan="2"><input type="submit" name="submit" value="Select" class="submit2"></td>
                </tr>
            </table>
            </form>
        <%end if%>
		</td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->