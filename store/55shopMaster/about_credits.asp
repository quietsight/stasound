<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="About ProductCart&reg; - Copyright &amp; Acknowledgments" %>
<% Section="about" %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<% dim mySQL, conntemp, rstemp %>
<!--#include file="AdminHeader.asp"-->                
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>ProductCart&reg; - Copyright&copy; Information</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<td>
		<p>ProductCart&reg; is a registered trademark property of NetSource Commerce, a Florida Corporation.</p>
		<p>ProductCart, its source code, graphical interface, logos, and documentation are property of NetSource Commerce and are protected by US and International copyright laws.</p>
		<p>Copyright&copy; 2001-<%=Year(now)%> <a href="http://www.productcart.com" target="_blank">NetSource Commerce</a>. All rights reserved.</p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Editing Rights</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<p>You have the right to make changes to ProductCart's Control Panel. For example, you may change the look &amp; feel of the Control Panel, including this page. However, you <strong>MAY NOT </strong>remove any of the copyright information included with any ProductCart file.</p>
			<p>For more information about the <a href="about_terms.asp">Terms &amp; Conditions</a> under which you may use ProductCart, please <a href="about_terms.asp">click here</a>.</p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 
		<th>Credits &amp; Acknowledgments</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		<p>The idea and  the  code for the <u>image upload &amp; resize</u> feature was provided by <a href="http://www.ironhammer.com/" target="_blank">Iron Hammer</a>,  Certified ProductCart Developer.</p>
		<p>For information about Certified ProductCart Developers, <a href="http://www.earlyimpact.com/partners.asp" target="_blank">visit</a> the NetSource Commerce Web site. These are companies who have become experts in providing custom ecommerce solutions based on ProductCart.</p>
		</td>
	</tr>
	<tr>        
		<td>
			<p>Elements of ProductCart's original database and limited portions of its original storefront source code were based on open source e-commerce software developed by <a href="http://www.comersus.com/" target="_blank">Comersus</a>&#153;. Today's version of ProductCart contains virtually no traces of such code.</p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="center">
		<form class="pcForms">
		<input type="button" name="back" value="Back" onClick="javascript:history.back()">
		&nbsp;
		<input type="button" name="" value="Start Page" onClick="location.href='menu.asp'">
		</form>
		</td>
	</tr>  
</table>
<!--#include file="AdminFooter.asp"-->