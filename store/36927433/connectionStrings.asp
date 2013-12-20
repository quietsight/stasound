<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>ProductCart Shopping Cart Software: Setup Wizard</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcSetup.css" />
</head>
<body>
<div id="pcSetup">
	<table>
  	<tr>
    	<td>
				<h1>Examples of Database Connection Strings</h1>
				<ul>
            <li>Check with your Web hosting company for supported connection strings.</li>
						<li>The example paths refer to the hard disk on your Web hosting company's server (e.g. change &quot;c&quot; to the actual name of the drive).</li>
				</ul>
				<hr>     
				<p><strong>DSN-Less connection to a MS SQL database</strong></p>
				<p><span style="background-color:#FFFF99">Provider=sqloledb;Data Source=<font color="#0000FF">SERVER-IP</font>,1433;Initial Catalog=<font color="#0000FF">DB-NAME</font>;User Id=<font color="#0000FF">USER</font>;Password=<font color="#0000FF">PWD</font>;</span></p>
<ul>
				<li>&ldquo;SERVER-IP&rdquo; is the server&rsquo;s IP address</li>
				<li>&ldquo;DB-NAME&rdquo; is the name of the database</li>
				<li>&ldquo;USER&rdquo; and &ldquo;PWD&rdquo; are the user name  and password that grant access to it (use SQL authentication, not Windows authentication).</li>
			</ul>
			<p><strong>DSN Connection to a MS SQL database</strong></p>
			<p><span style="background-color:#FFFF99">DSN=SQLDSN;UID=<font color="#0000FF">USER</font>;PWD=<font color="#0000FF">PWD</font></span></p>
			<ul>
				<li>&ldquo;SQLDSN&rdquo; is the name of the DSN</li>
				<li>&ldquo;USER&rdquo; and &ldquo;PWD&rdquo; are the user name and password that grant access to the database (use SQL authentication, not Windows authentication).</li>
			</ul>
			<hr>
			<p><strong>DSN-Less connection to a MS Access database</strong></p>
			<p><span style="background-color:#FFFF99">DRIVER={Microsoft Access Driver (*.mdb)};DBQ=<font color="#0000FF">c:\anydatabase.mdb</font></span></p>
			<ul>
			<li>Where... &quot;c:\anydatabase.mdb&quot; is the physical path to the database.</li>
			</ul>
			<p><strong>DSN connection to a MS Access database</strong></p>
			<p><span style="background-color:#FFFF99">DSN=<font color="#0000FF">AccessDSN</font></span></p>
			<ul>
<li>Where... &quot;AccessDSN&quot; is the name of the DSN.</li>
</ul></td>
	</tr>
	</table>
</div>
</body>
</html>