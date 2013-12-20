<!--#include file="pcSetupHeader.asp"-->
      	<h1>Step 1: Review ProductCart activation checklist</h1>
			<ul>
				<li><strong><a href="http://wiki.earlyimpact.com/productcart/install" target="_blank">File Upload</a></strong>. You have uploaded the unzipped <em>productcart</em> folder 'as is' to your Web server.</li>
				<li><strong><a href="http://wiki.earlyimpact.com/productcart/install-db" target="_blank">MS SQL Database Setup</a></strong> (skip this step if you are using MS Access): You have run the SQL script to use ProductCart with a SQL database. The script is located in the <em>productcart/database</em> folder. <a href="http://wiki.earlyimpact.com/productcart/install-db" target="_blank">More information</a>.</li>
				<li><strong><a href="http://wiki.earlyimpact.com/productcart/install-permissions" target="_blank">Folder Permissions</a></strong>. You have assigned folder permissions as follows: (a) 'Read/Write' permissions for the <em>productcart</em> folder and all of its subfolders</font>; (b)'Read/Write/Delete' permissions for the <em>productcart/includes</em> folder and all of its subfolders.</li>
				<li><strong><a href="http://wiki.earlyimpact.com/productcart/install-db" target="_blank">Database Connection String</a></strong></strong>. You have created a database connection pointing to the right file. Please review the <a href="step1b.asp">ProductCart Security Tips</a> on the next page before creating your database connection string. <a href="javascript:win('connectionStrings.asp')">Click here</a> for <a href="javascript:win('connectionStrings.asp')">examples</a> of supported database connection strings.</li>
				<li><strong>Path to Your Store</strong>. You know the correct path to your store. For example: http://www.mystore.com/ indicates that the <em>productcart</em> folder is located in the root. You will have to enter this information on the Setup form.</li>
				<li><strong><a href="http://wiki.earlyimpact.com/productcart/license" target="_blank">ProductCart License</a></strong>. You have all the necessary license information: E-mail on file (or Partner ID), KeyID, User Name, and temporary Password. You will have to enter this information on the final step of this Wizard.</li>
			</ul>
              <hr>
              <p align="center">
                    <form class="pcForms">
                    <input name="next" type="button" value="Proceed to Security Tips" onClick="location.href='step1b.asp'" class="submit2">
                    <input name="back" type="button" value="Back" onClick="JavaScript: history.back()" class="ibtnGrey">
                    </form>
              </p>
<!--#include file="pcSetupFooter.asp"-->