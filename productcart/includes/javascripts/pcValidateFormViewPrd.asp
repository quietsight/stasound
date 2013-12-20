<script language="JavaScript">
<!--		
		<%
		pcv_strNumberValidations = xfieldCnt + xOptionsCnt + 1
		
		str = ""
		x = 1		 
		do until x = pcv_strNumberValidations
			if x > 1 then 
			str = str & ", "
			end if
		str = str & "choice" & x & ", " & "option" & x
		x = x + 1
		loop
		%>		
		
		function cdDynamic(<%=str%><%if str<>"" then%>,rtype<%end if%>) 
		{
		
		<% 
		x = 1		 
		do until x = pcv_strNumberValidations
		%>
			if (choice<%=x%>== 0) {
				alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ option<%=x%> + "<%=dictLanguage.Item(Session("language")&"_alert_2")%>.\n");			
				} else {
		<% 
		x = x + 1
		loop
		%>	
				if (rtype==0)
				{
					document.additem.action="<%=pcv_strFormAction%>";
					document.additem.method="POST";
					document.additem.submit();
				}
				else
				{
					return(true);
				}
		
		<%
		x = 1		 
		do until x = pcv_strNumberValidations
		%>				
			}
		<% 
		x = x + 1
		loop
		%>
			if (rtype==1)
			{
				return(false);
			}
		}				


        function CheckRequiredCS(reqstr)
        {
            if (reqstr.length>0)
            {
                var objArray = reqstr.split(",");
                var i = 0;
                while (i < objArray.length)
                {
                    var obj = eval(objArray[i]);
			        if (obj.checked==0) 
			        {
				        alert("<%=dictLanguage.Item(Session("language")&"_alert_1")%>\n"+ obj.value + "<%=dictLanguage.Item(Session("language")&"_alert_2")%>.\n");
				        return false;
			        }     
			        i+=1;       
			    }
			}
			return true;
        }
		
//-->
</script>