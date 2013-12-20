<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Currency Codes" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">

            <tr>
                <td class="pcCPspacer" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td class="pcCPspacer" colspan="2"><strong>Frequently used</strong>:</td>
            </tr>
            <tr>
                <td class="pcCPspacer" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td valign="bottom" style="border:solid gray .75pt;background:#E0E0E0; padding:0in 5.4pt 0in 5.4pt">
                <p class="MsoNormal"><b>CURRENCY COUNTRY</b></p></td>
                <td valign="bottom" style="border:solid gray .75pt;border-left:none;background:#E0E0E0;padding:0in 5.4pt 0in 5.4pt">
                <p class="MsoNormal"><b>CURRENCY CODE</b></p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">US Dollar (United States)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">USD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Canadian Dollar (Canada)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CAD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pound Sterling (United Kingdom)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GBP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Euro (Europe)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">EUR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Australian Dollar (Australia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AUD</p></td>
            </tr>
            <tr>
                <td class="pcCPspacer" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td class="pcCPspacer" colspan="2"><strong>All currency codes</strong>:</td>
            </tr>
            <tr>
                <td class="pcCPspacer" colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td valign="bottom" style="border:solid gray .75pt;background:#E0E0E0; padding:0in 5.4pt 0in 5.4pt">
                <p class="MsoNormal"><b>CURRENCY COUNTRY</b></p></td>
                <td valign="bottom" style="border:solid gray .75pt;border-left:none;background:#E0E0E0;padding:0in 5.4pt 0in 5.4pt">
                <p class="MsoNormal"><b>CURRENCY CODE</b></p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Afghani (Afghanistan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AFA</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Algerian Dinar (Algeria)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">DZD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Andorran Peseta (Andorra)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ADP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Argentine Peso (Argentina)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ARS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Armenian Dram (Armenia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AMD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Aruban Guilder (Aruba)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AWG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Australian Dollar (Australia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AUD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Azerbaijanian Manat (Azerbaijan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AZM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Bahamian Dollar (Bahamas)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BSD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Bahraini Dinar (Bahrain)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BHD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Baht (Thailand)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">THB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Balboa (Panama)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PAB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Barbados Dollar (Barbados)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BBD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Belarussian Ruble (Belarus)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BYB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Belgian Franc (Belgium)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BEF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Belize Dollar (Belize)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BZD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Bermudian Dollar (Bermuda)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BMD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Bolivar (Venezuela)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">VEB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Boliviano (Bolivia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BOB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Brazilian Real (Brazil)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BRL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Brunei Dollar (Brunei Darussalam)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BND</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Bulgarian Lev (Bulgaria)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BGN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Burundi Franc (Burundi)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BIF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Canadian Dollar (Canada)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CAD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cape Verde Escudo (Cape Verde)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CVE</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cayman Islands Dollar (Cayman Islands)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KYD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cedi (Ghana)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GHC</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CFA Franc BCEAO (Guinea-Bissau)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XOF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CFA Franc BEAC (Central African Republic)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XAF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CFP Franc (New Caledonia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XPF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Chilean Peso (Chile)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CLP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Colombian Peso (Colombia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">COP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Comoro Franc (Comoros)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KMF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Convertible Marks (Bosnia And Herzegovina)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BAM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cordoba Oro (Nicaragua)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NIO</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Costa Rican Colon (Costa Rica)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CRC</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cuban Peso (Cuba)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CUP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Cyprus Pound (Cyprus)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CYP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Czech Koruna (Czech Republic)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CZK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Dalasi (Gambia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GMD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Danish Krone (Denmark)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">DKK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Denar (The Former Yugoslav Republic Of
            Macedonia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MKD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Deutsche Mark (Germany)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">DEM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Dirham (United Arab Emirates)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AED</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Djibouti Franc (Djibouti)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">DJF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Dobra (Sao Tome And Principe)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">STD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Dominican Peso (Dominican Republic)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">DOP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Dong (Vietnam)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">VND</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Drachma (Greece)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GRD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">East Caribbean Dollar (Grenada)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XCD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Egyptian Pound (Egypt)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">EGP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">El Salvador Colon (El Salvador)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SVC</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Ethiopian Birr (Ethiopia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ETB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Euro (Europe)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">EUR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Falkland Islands Pound (Falkland Islands)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">FKP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Fiji Dollar (Fiji)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">FJD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Forint (Hungary)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">HUF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Franc Congolais (The Democratic Republic Of
            Congo)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CDF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">French Franc (France)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">FRF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Gibraltar Pound (Gibraltar)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GIP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Gold</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XAU</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Gourde (Haiti)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">HTG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Guarani (Paraguay)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PYG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Guinea Franc (Guinea)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GNF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Guinea-Bissau Peso (Guinea-Bissau)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GWP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Guyana Dollar (Guyana)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GYD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Hong Kong Dollar (Hong Kong)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">HKD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Hryvnia (Ukraine)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">UAH</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Iceland Krona (Iceland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ISK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Indian Rupee (India)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">INR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Iranian Rial (Islamic Republic Of Iran)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">IRR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Iraqi Dinar (Iraq)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">IQD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Irish Pound (Ireland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">IEP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Italian Lira (Italy)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ITL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Jamaican Dollar (Jamaica)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">JMD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Jordanian Dinar (Jordan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">JOD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kenyan Shilling (Kenya)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KES</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kina (Papua New Guinea)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PGK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kip (Lao People's Democratic Republic)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LAK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kroon (Estonia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">EEK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kuna (Croatia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">HRK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kuwaiti Dinar (Kuwait)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KWD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kwacha (Malawi)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MWK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kwacha (Zambia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ZMK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kwanza Reajustado (Angola)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AOR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Kyat (Myanmar)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MMK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lari (Georgia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GEL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Latvian Lats (Latvia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LVL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lebanese Pound (Lebanon)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LBP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lek (Albania)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ALL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lempira (Honduras)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">HNL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Leone (Sierra Leone)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SLL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Leu (Romania)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ROL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lev (Bulgaria)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BGL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Liberian Dollar (Liberia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LRD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Libyan Dinar (Libyan Arab Jamahiriya)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LYD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lilangeni (Swaziland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SZL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Lithuanian Litas (Lithuania)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LTL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Loti (Lesotho)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LSL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Luxembourg Franc (Luxembourg)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LUF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Malagasy Franc (Madagascar)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MGF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Malaysian Ringgit (Malaysia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MYR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Maltese Lira (Malta)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MTL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Manat (Turkmenistan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TMM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Markka (Finland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">FIM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Mauritius Rupee (Mauritius)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MUR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Metical (Mozambique)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MZM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Mexican Peso (Mexico)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MXN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Mexican Unidad de Inversion (Mexico)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MXV</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Moldovan Leu (Republic Of Moldova)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MDL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Moroccan Dirham (Morocco)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MAD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Mvdol (Bolivia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BOV</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Naira (Nigeria)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NGN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Nakfa (Eritrea)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ERN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Namibia Dollar (Namibia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NAD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Nepalese Rupee (Nepal)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NPR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Netherlands (Netherlands)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ANG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Netherlands Guilder (Netherlands)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NLG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Dinar (Yugoslavia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">YUM</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Israeli Sheqel (Israel)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ILS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Kwanza (Angola)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">AON</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Taiwan Dollar (Province Of China
            Taiwan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TWD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Zaire (Zaire)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ZRN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">New Zealand Dollar (New Zealand)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NZD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Ngultrum (Bhutan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BTN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">North Korean Won (Democratic People's Republic
            Of Korea)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KPW</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Norwegian Krone (Norway)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">NOK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Nuevo Sol (Peru)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PEN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Ouguiya (Mauritania)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MRO</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pa'anga (Tonga)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TOP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pakistan Rupee (Pakistan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PKR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Palladium</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XPD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pataca (Macau)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MOP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Peso Uruguayo (Uruguay)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">UYU</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Philippine Peso (Philippines)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PHP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Platinum</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XPT</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Portuguese Escudo (Portugal)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PTE</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pound Sterling (United Kingdom)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GBP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Pula (Botswana)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BWP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Qatari Rial (Qatar)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">QAR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Quetzal (Guatemala)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">GTQ</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rand (Financial) (Lesotho)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ZAL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rand (South Africa)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ZAR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rial Omani (Oman)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">OMR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Riel (Cambodia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KHR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rufiyaa (Maldives)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MVR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rupiah (Indonesia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">IDR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Russian Ruble (Russian Federation)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">RUB</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Russian Ruble (Russian Federation)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">RUR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Rwanda Franc (Rwanda)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">RWF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Saudi Riyal (Saudi Arabia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SAR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Schilling (Austria)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ATS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Seychelles Rupee (Seychelles)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SCR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Silver</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">XAG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Singapore Dollar (Singapore)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SGD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Slovak Koruna (Slovakia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SKK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Solomon Islands Dollar (Solomon Islands)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SBD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Som (Kyrgyzstan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KGS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Somali Shilling (Somalia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SOS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Spanish Peseta (Spain)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ESP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Sri Lanka Rupee (Sri Lanka)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">LKR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">St Helena Pound (St Helena)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SHP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Sucre (Ecuador)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ECS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Sudanese Dinar (Sudan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SDD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Surinam Guilder (Suriname)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SRG</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Swedish Krona (Sweden)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SEK</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Swiss Franc (Switzerland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CHF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Syrian Pound (Syrian Arab Republic)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SYP</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tajik Ruble (Tajikistan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TJR</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Taka (Bangladesh)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">BDT</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tala (Samoa)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">WST</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tanzanian Shilling (United Republic Of
            Tanzania)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TZS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tenge (Kazakhstan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KZT</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Timor Escudo (East Timor)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TPE</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tolar (Slovenia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">SIT</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Trinidad and Tobago Dollar (Trinidad And
            Tobago)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TTD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tugrik (Mongolia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">MNT</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Tunisian Dinar (Tunisia)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TND</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Turkish Lira (Turkey)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">TRL</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Uganda Shilling (Uganda)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">UGX</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Unidad de Valor Constante (Ecuador)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ECV</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Unidades de fomento (Chile)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CLF</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">US Dollar (Next day) (United States)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">USN</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">US Dollar (Same day) (United States)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">USS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">US Dollar (United States)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">USD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Uzbekistan Sum (Uzbekistan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">UZS</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Vatu (Vanuatu)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">VUV</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Won (Republic Of Korea)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">KRW</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Yemeni Rial (Yemen)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">YER</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Yen (Japan)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">JPY</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Yuan Renminbi (China)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">CNY</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Zimbabwe Dollar (Zimbabwe)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">ZWD</p></td>
            </tr>
            <tr>
            <td valign="bottom" style="border:solid gray .75pt;border-top:none;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">Zloty (Poland)</p></td>
            <td valign="bottom" style="border-top:none;border-left:none;border-bottom:solid gray .75pt; border-right:solid gray .75pt;padding:0in 5.4pt 0in 5.4pt">
            <p class="MsoNormal">PLN</p></td>
            </tr>
        </table>
<!--#include file="AdminFooter.asp"-->