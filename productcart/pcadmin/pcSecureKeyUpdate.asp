<% pageTitle = "Generate New Encryption Key" %>
<% section = "" %>
<%PmAdmin=9%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/rc4.asp"--> 
<!--#include file="pcKeyFunctions.asp"--> 
<!--#include file="AdminHeader.asp"-->
<%
dim rs, conntemp, query

call openDb()

if request("Generate")="YES" then
	pcv_NewSecurityKey =gen_pass(17) & gen2_pass(12)

	query="UPDATE pcSecurityKeys SET pcActiveKey = 0;"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)

	query="INSERT INTO pcSecurityKeys (pcSecurityKey, pcActiveKey, pcDateUpdated) VALUES ('"&pcv_NewSecurityKey&"',1,'"&nOW()&"');"
	set rs=Server.CreateObject("ADODB.Recordset")
	set rs=connTemp.execute(query)
	set rs=nothing

	call closeDb()
	response.redirect "pcSecureKeyUpdate.asp?s=1&msg=This key was updated successfully!"
end if

query="SELECT pcDateUpdated FROM pcSecurityKeys WHERE pcActiveKey=1;"
set rs=Server.CreateObject("ADODB.Recordset")
set rs=connTemp.execute(query) 
if rs.eof then
	'no keys have been generated yet
	pcDateUpdated = "0"
else
	pcDateUpdated = rs("pcDateUpdated")
end if
set rs=nothing
call closeDb()

%>
<form method="post" action="pcSecureKeyUpdate.asp" class="pcForms">
    <table class="pcCPcontent">
        <tr>
            <td>
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
              <% 	' END show message %>
                <input type="hidden" name="Generate" value="YES"> 
            </td>
        </tr>
        <% If pcDateUpdated = "0" then %>
            <tr>
                <td class="pcCPspacer"><p>It appears that you have not generated a new encryption key since you installed or upgraded to ProductCart v4. PCI (Protection of Cardholder Information) standards require that you generate a new encryption key at least once a year. You must perform this task annually to maintain PCI compliance, and you can use this feature to do so.</p>
                <p>See the ProductCart documentation for more information about ProductCart and PCI compliance.<a href="http://wiki.earlyimpact.com/productcart/pci-compliance" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Learn more about ProductCart and PCI compliance" title="Learn more about ProductCart and PCI compliance" hspace="5" vspace="5"></a>
                  <ul>
                    <li>Existing data will remain encrypted with the current encryption key</li>
                    <li>New data will be encrypted with the new key that you are about to generate</li>
                    <li>ProductCart will be able to unencrypt both old and new data</li>
                </ul></td>
            </tr>
        <% Else %>     
            <tr>
                <td class="pcCPspacer"><p><strong>You last created a new encryption key on: <%=pcDateUpdated%></strong>.
                <br /><br />
                PCI (Protection of Cardholder Information) standards require that you generate a new encryption key at least once a year. You must perform this task annually to maintain PCI compliance, and you can use this feature to do so.</p>
                <p>See the ProductCart documentation for more information about ProductCart and PCI compliance.<a href="http://wiki.earlyimpact.com/productcart/pci-compliance" target="_blank"><img src="images/pcv3_infoIcon.gif" width="16" height="16" alt="Learn more about ProductCart and PCI compliance" title="Learn more about ProductCart and PCI compliance" hspace="5" vspace="5"></a>
                  <ul>
                    <li>Existing data will remain encrypted with the current encryption key</li>
                    <li>New data will be encrypted with the new key that you are about to generate</li>
                    <li>ProductCart will be able to unencrypt both old and new data</li>
                </ul></td>
            </tr>                    
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
        <% End If %>
        <tr>
            <td>
            <hr />
            <input type="submit" name="submit" value="Generate New Encryption Key" class="submit2">
            </td>
        </tr>           
    </table>
</form>
<!--#include file="AdminFooter.asp"-->