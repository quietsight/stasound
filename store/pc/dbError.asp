<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../includes/emailsettings.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Database Connection Problem</title>
</head>
<body>
<div style="padding:20px; margin:50px; background-color:#e1e1e1; border: solid 1px #CCCCCC; font-family:Verdana, Arial, Helvetica, sans-serif; font-size: 12px; width: 350px;">Dear customer,<br /><br />
We are currently experiencing a database connection problem on our store.<br /><br />We apologize for the inconvenience.<br /><br />Please <a href="mailto:<%=scFrmEmail%>?subject=Database Connection Problem&body=I received an error when visiting your store. The system indicated that there is a database connection problem. Please let me know when it has been resolved.">notify the store administrator</a> so that we may fix the problem as soon as possible and let you know when the store becomes available again.<br /><br />Thank you for your cooperation!
</div>
</body>
</html>