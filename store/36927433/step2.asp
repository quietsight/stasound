<!--#include file="pcSetupHeader.asp"-->
        <form name="form1" action="step2_results.asp" method="post" class="pcForms">
     	 <h1>Step 2: Web server readiness utility</h1>
			<p>ProductCart will now run some tests to see if your Web server is ready. The tests are as follows:</p>
			<ol>
				<li><strong>Folder permissions</strong>. ProductCart will check whether the correct folder permissions have been assigned.</li>
				<li><strong>Database connection string</strong>. ProductCart will test whether it can successfully connect to your database. Please enter your database connection string below (<a href="javascript:win('connectionStrings.asp')">view examples</a>):<br /><br />
        <input name="dbConn" type="text" id="dbConn" size="90">
				<br /><br /></li>
				<li><strong>Available E-mail Components.</strong> ProductCart will check the server and let you know which e-mail components are installed, so you can decide which one to use. You will choose the e-mail component in the &quot;General Settings/E-mail Settings&quot; area of the Control Panel.</li>
				<li><strong>Parent Paths</strong>. Depending on whether parent paths are enabled or disabled on your server, you should use the corresponding version of ProductCart. <a href="http://wiki.earlyimpact.com/productcart/install#parent_paths_enabled_and_disabled" target="_blank">More information</a>.</li>
				<li><strong>XML Parser</strong>. We will check to see whether your server has a compatible version of the MS XML Parser installed. This is used by ProductCart to communicate with third-party servers in a variety of scenarios (shipping module, payment gateways, etc.).</li>
			</ol>
            <hr>
			<p align="center">
			<input name="next" type="submit" value="Run Tests" class="submit2">
			<input name="back" type="button" value="Back" onClick="JavaScript: history.back()" class="ibtnGrey">
            </p>
        </form>
<!--#include file="pcSetupFooter.asp"-->