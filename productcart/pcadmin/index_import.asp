<% pageTitle = "Product Import Wizard - Upload/Locate Data File" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
if request("action")="select" then
	if request("ways")="1" then
	response.redirect "index_import1.asp"
	else
		if request("ways")="2" then
		response.redirect "index_import2.asp"
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
        <td width="5%" align="right"><img border="0" src="images/step1a.gif"></td>
        <td width="95%"><b>Select product data file</b></td>
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
           
    <br /> 
    <% ' START show message, if any %>
        <!--#include file="pcv4_showMessage.asp"-->
    <% 	' END show message %>
    
    <% if (request.querystring("nextstep")=1) then %>
    <form action="step2.asp" method="post" class="pcForms">
		<table class="pcCPcontent">
            <tr>
                <td colspan="2">
                    <input type=radio name="append" value="0" <%if session("append")<>"1" then%> checked<%end if%> onClick="JavaScript:document.getElementById('show_1').style.display='none'"> Import new data into the store database
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <input type=radio name="append" value="1" <%if session("append")<>"1" then
                    else%>checked<%end if%> onClick="JavaScript:document.getElementById('show_1').style.display=''"> Update current data if product SKU is an exact match with an existing SKU
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <table id="show_1" class="pcCPcontent" style="display:none; margin: 10px;">
                    <tr>
                        <th colspan="2">How should category information be updated?</th>
                    </tr>
                    <tr>
                    	<td colspan="2" class="pcCPspacer"></td>
                    </tr>
                    <tr>
                        <td width="7%">
                            <input type="radio" name="movecat" value="1" <%if session("movecat")<>"1" then
                            else%>checked<%end if%> checked></td>
                        <td width="93%">Add the category specified for a product to the list of categories that the product is assigned to</td>
                    </tr>
                    <tr>
                        <td width="7%">
                            <input type="radio" name="movecat" value="2" <%if session("movecat")="2" then%>checked<%end if%>></td>
                        <td width="93%">Move the product from the existing category to the new category specified in the data file</td>
                    </tr>
                    <tr>
                        <td width="7%">
                            <input type="radio" name="movecat" value="3" <%if session("movecat")="3" then%>checked<%end if%>></td>
                        <td width="93%">Ignore any category information included in the data file</td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <p align="right"><input type=submit name=Go value="Go to Step 2 >>" class="submit2"></p>
                </td>
            </tr>
            </table>
            </form>
	<%
	else
	%>
        <form method="post" action="index_import.asp?action=select" class="pcForms">
        <table class="pcCPcontent">
        <tr>
            <td colspan="2">
                <p>Select how ProductCart should locate your data file. You can either upload the file now, or provide a location on your Web server if the file has already been uploaded. Please note that <strong>ONLY</strong> <strong>*.csv</strong> &amp; <strong>*.xls</strong> files are accepted. For more information on what fields can be imported and on how to prepare your *.csv &amp; *.xls files for import, please refer to the ProductCart <a href="http://wiki.earlyimpact.com/productcart/products_import" target="_blank">User Guide</a>.<p>
            </td>
        </tr>
        <tr> 
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr> 
            <th colspan="2">Product data file location</th>
        </tr>
        <tr> 
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td width="5%"><input type=radio name="ways" value="1" checked></td>
            <td width="95%">Use the Import Wizard to upload the data file to the server</td>
        </tr>
        <tr>
            <td><input type=radio name="ways" value="2"></td>
            <td>The data file has been manually uploaded to the server</td>
        </tr>
        <tr> 
            <td colspan="2" class="pcCPspacer"></td>
        </tr>
        <tr>
            <td colspan="2"><input class="submit2" type="submit" name="submit" value="Select"></td>
        </tr>
	</table>
	</form>
<%end if%>
<!--#include file="AdminFooter.asp"-->