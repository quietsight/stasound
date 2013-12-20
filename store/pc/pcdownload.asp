<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languages.asp"-->
<%
'Get Path Info
pcv_filePath=Request.ServerVariables("PATH_INFO")
do while instr(pcv_filePath,"/")>0
	pcv_filePath=mid(pcv_filePath,instr(pcv_filePath,"/")+1,len(pcv_filePath))
loop

pcv_Query=Request.ServerVariables("QUERY_STRING")

if pcv_Query<>"" then
pcv_filePath=pcv_filePath & "?" & pcv_Query
end if

' verifies if customer is logged, so as not send to login page
if Session("idCustomer")=0 then
	response.redirect "Checkout.asp?cmode=1&redirectUrl="&Server.URLEncode(pcv_filePath)
end if
%>
<html>
<head>
<title><%response.write dictLanguage.Item(Session("language")&"_pcDownload_1")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" rel="stylesheet" href="pcStorefront.css" />
</head>
<body style="margin: 0;">
<div id="pcMain">
<table class="pcMainTable">
	<tr>
    <td width="70%%">
		<h2><%response.write dictLanguage.Item(Session("language")&"_pcDownload_1")%></h2>
		</td>
    <td width="30%" valign="top" align="right">
    <img border="0" src="catalog/yourlogohere.gif">
		</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" valign="top">
			<%
			dim query, connTemp, rsTemp, rs   
			call opendb()
			DownloadID=replace(request("id"),"'","''")
			if DownloadID<>"" then
			 query="select idOrder, idCustomer, idProduct, startDate from DPRequests where RequestSTR='" & DownloadID & "' and IDcustomer=" & Session("idCustomer")
			 set rsTemp=Server.CreateObject("ADODB.Recordset")
			 set rsTemp=connTemp.execute(query)
			 IF rsTemp.eof then
			 %>
			 	<div class="pcErrorMessage"><%response.write dictLanguage.Item(Session("language")&"_pcDownload_2")%></div>
			 <%
			 ELSE
				pIdOrder=rstemp("idOrder")
				pIdCustomer=rstemp("idCustomer")
				pIdProduct=rstemp("idProduct")
				pProcessDate=rstemp("StartDate")
				
				query="select name,lastname,email from Customers where idcustomer=" & pIdCustomer
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=connTemp.execute(query)
				
				pCInfor=rs("name") & " " & rs("lastname") & " - E-mail: " & rs("email")
				query="select sku,description from Products where idproduct=" & pIdProduct
				set rs=connTemp.execute(query)
				psku=rs("sku")
				pName=rs("description")
				%>
			 	<%response.write dictLanguage.Item(Session("language")&"_pcDownload_3")%><%=(scpre + int(pIdOrder))%><br>
				<%response.write dictLanguage.Item(Session("language")&"_pcDownload_4")%><%=pCInfor%><br>
				<%response.write dictLanguage.Item(Session("language")&"_pcDownload_5")%><%=pidproduct%><br>
				<%response.write dictLanguage.Item(Session("language")&"_pcDownload_6")%><b><%=pName%></b><br>
				<%response.write dictLanguage.Item(Session("language")&"_pcDownload_7")%><%=psku%><br>
				<%
				query="select URLExpire, ExpireDays from DProducts where Idproduct=" & pIdProduct
				'response.write query
				set rs=connTemp.execute(query)  
				pURLExpire=rs("URLExpire")
				pExpireDays=rs("ExpireDays")
				myTest=true
				myMsg=""
				if (pURLExpire<>"") and (pURLExpire="1") then
					if date()-(CDate(pprocessDate)+pExpireDays)<0 then
			 			myMsg="<br><b>" & dictLanguage.Item(Session("language")&"_pcDownload_8") & "</b>: " & dictLanguage.Item(Session("language")&"_pcDownload_9") & (CDate(pprocessDate)+pExpireDays)-date() & dictLanguage.Item(Session("language")&"_pcDownload_10")
			 		else
						if date()-(CDate(pprocessDate)+pExpireDays)=0 then
			 				myMsg="<br><b>" & dictLanguage.Item(Session("language")&"_pcDownload_8") & "</b>: " & dictLanguage.Item(Session("language")&"_pcDownload_11")
							else
							myTest=false
							myMsg="<br><b>" & dictLanguage.Item(Session("language")&"_pcDownload_8") & "</b>: " & dictLanguage.Item(Session("language")&"_pcDownload_12")
						end if
					end if
				end if
				
				if myTest=false then%>
				 <div class="pcErrorMessage"><%=myMsg%></div>
				<%else%> 
				 <%if myMsg<>"" then%>
				 <div class="pcErrorMessage"><%=myMsg%></div>
				 <%end if%>
				 <div>
				 <a href="downloadnow.asp?id=<%=DownloadID%>"><%response.write dictLanguage.Item(Session("language")&"_pcDownload_13")%></a>
				 </div>
				<%
				end if
			END IF
		end if
		%>	

    </td>
  </tr>
</table>
</div>
</body>
</html>