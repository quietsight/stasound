<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "ProductCart Online Help - Support Center" %>
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Troubleshooting Utility</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<p>Is something wrong with your store? For example, are you trying to figure out why the store cannot send e-mails or successfully connect to your payment gateway? Try running the built-in <a href="pcTSUtility.asp">troubleshooting utility</a> to look for possible issues. <a href="pcTSUtility.asp">Run the utility &gt;&gt;</a></p>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Keep your Store Up-To-Date</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<ul class="pcListIcon">
            	<li><a href="http://www.earlyimpact.com/productcart/support/updates/pcvm_listPatches.asp?keyid=<%=scCrypPass%>&vnum=<%=scVersion%>&svnum=<%=scSubVersion%>&sp=<%=scSP%>" target="_blank">Check for Updates &gt;&gt;</a></li>
                <li><a href="http://www.earlyimpact.com/productcart/support/#updates" target="_blank">Recent bug fixes &gt;&gt;</a> (in-between official releases of ProductCart).</li>
            </ul>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Manuals, Forums, Knowledge Base, etc.</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		<p>ProductCart is a very feature-rich application: there are a number of useful resources that will help you quickly get up to speed. Find an answer to a question by <strong>running a search</strong> (this is powered by  Google Site Search):</p>
		<div style="text-align: center; margin: 15px;">
			<div style="background-color: #FFFFCC; border: 1px solid #999999; margin-top: 15px; text-align: center; width: 600px; height: 70px; position: relative;">
				<!-- Google CSE Search Box Begins -->
					<form id="searchbox_015541040505873019426:rx6ixm-0y6i" action="http://www.earlyimpact.com/search.asp" style="margin:0;" target="_blank">
          	<div style="position: absolute; left: 20px; top: 20px;">
						<input type="hidden" name="cx" value="015541040505873019426:rx6ixm-0y6i" />
						<input type="hidden" name="cof" value="FORID:11" />
						<input name="q" type="text" size="60" style="height: 23px;" />
            </div>
            <div style="position: absolute; left: 420px; top: 20px;">
						<input type="submit" name="sa" value="Find a solution" style="border: 1px #ccc solid; background-color: #0066FF; color:#FFF; font-weight: bold; height: 30px;" />
            </div>
					</form>
				<!-- Google CSE Search Box Ends -->
			</div>
		</div>
		<p>Other resources:</p>
		<ul class="pcListIcon">
			<li><a href="http://wiki.earlyimpact.com/" target="_blank">Productcart User Guides</a>: The User Guides are now a WIKI, constantly updated and expanded (get an <a href="http://wiki.earlyimpact.com/feed.php" target="_blank">RSS feed of recent changes</a>).</li>
		  	<li><a href="http://www.earlyimpact.com/kb.html" target="_blank"> Knowledge Base</a>: Answers to many frequently asked questions.</li>
			<li><a href="http://www.earlyimpact.com/forum" target="_blank">Forums</a>: Communicate with other ProductCart users. </li>
			<li><a href="http://wiki.earlyimpact.com/developers/developers" target="_blank">Developer's Corner</a> If you are an experience ASP programmer and are interested in customizing ProductCart, this is the place to go. </li>
			<li><a href="http://www.earlyimpact.com/productcart/support/" target="_blank">Support Center</a>: Visit the support center for additional resources.</li>
			<li><strong>Submit a Support Request: </strong>
              <ul>
                <li>Did you purchase ProductCart from NetSource Commerce? <a href="https://www.earlyimpact.com/eistore/productcart/pc/custpref.asp" target="_blank">Log into your account to submit a ticket &gt;&gt;</a></li>
                <li>Did you purchase from a ProductCart reseller? Please contact them for technical support.</li>
                <li>More information? See our <a href="https://www.earlyimpact.com/productcart/support_updates.asp" target="_blank">Technical Support Policy</a>.</li>
              </ul>
			</li>
		</ul>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Error Finder</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
			<p>By default ProductCart does not show detailed error messages in the storefront. This is both a customer-oriented feature (technical errors may scare customers away) and a security measure (e.g. avoids providing any information on your database to a potential hacker). To obtain more information on an error you can:</p>
			<ul class="pcListIcon">
				<li>Enter the error's <span style="font-weight: bold">Reference ID</span> in the <a href="helpErrorFinder.asp">Error Finder</a>.</li>
				<li>Turn <span style="font-weight: bold">off</span> the <a href="AdminSettings.asp?tab=4">Error Handler</a>. Do this only temporarily.</li>
				<li>Debugging code? If you need <a href="http://wiki.earlyimpact.com/how_to/show_raw_errors" target="_blank">raw errors</a> printed to the screen, see this <a href="http://wiki.earlyimpact.com/how_to/show_raw_errors" target="_blank">WIKI article</a>.</li>
			</ul>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<th>Issues with your ProductCart license</th>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td><p>If you are having issues with your ProductCart license, please contact the company from which you have purchased the license. ProductCart includes a license protection system that links the License Key to the Store URL. See <a href="http://wiki.earlyimpact.com/productcart/licensing" target="_blank">how ProductCart is licensed</a>.</p></td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->