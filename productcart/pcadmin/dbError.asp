<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Database Connection Problem</title>
</head>
<body>
<div style="padding:20px; margin:50px; background-color:#e1e1e1; border: solid 1px #CCCCCC; font-family:Verdana, Arial, Helvetica, sans-serif; font-size: 12px; width: 350px;">
  <p>You are receiving this message because <strong>the system is not able to connect to the database</strong>. Please take the following steps:
  <br /><br />
  (1)   If you are using <strong>MS Access</strong>: reset permissions on the database file. The database file must have Read/Write access.
  <br /><br />
  (2) MS Access and MS SQL: Make sure that the <strong>database connection string</strong> is correct. To test it, use the ProductCart Setup Wizard, even if you have already setup and activated the shopping cart. Consult the <a href="http://wiki.earlyimpact.com/productcart/install-db#connecting_to_the_database" target="_blank">ProductCart User Guide</a> for examples of connection strings that you may use with ProductCart. 
  <br /><br />
  (3) If you are using a <strong>DSN connection</strong>, make sure that the name of the connection is correct, and that it is pointing to the right file. To do so, you will have to check the settings on your Web server. Many Web hosting companies allow you to do this through an administration panel. In other cases, you will have to contact the hosting company.</p>
  <p>(4) If you are using a <strong>MS SQL database</strong>, make sure that you can access the SQL server and that the user that is listed in your database connection strings has ownership rights on the database (it must be a database owner or &quot;DBO&quot;). For more information, please contact your Web hosting company.
  </p>
</div>
</body>
</html>