<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="XML Tools - Allowed IP Addresses" %>
<% section="layout" %>
<%PmAdmin=19%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<%
Dim rs, connTemp, query
call openDB()

IF request("action")="upd" THEN
	IF request("submit3")<>"" then
		TurnOn=0
	
		query="SELECT pcXIP_id,pcXIP_IPAddr,pcXIP_TurnOn FROM pcXMLIPs;"
		set rs=connTemp.execute(query)
		if not rs.eof then
			TurnOn=rs("pcXIP_TurnOn")
		end if
		set rs=nothing
		
		For i=1 to 10
			tmpIP=""
			if (request("N1_" & i)<>"") AND (request("N2_" & i)<>"") AND (request("N3_" & i)<>"") AND (request("N4_" & i)<>"") then
				tmpIP=cint(request("N1_" & i)) & "." & cint(request("N2_" & i)) & "." & cint(request("N3_" & i)) & "." & cint(request("N4_" & i))
			end if
			if tmpIP<>"" then
				query="DELETE FROM pcXMLIPs WHERE pcXIP_IPAddr LIKE '" & tmpIP & "';"
				set rs=connTemp.execute(query)
				set rs=nothing
				query="INSERT INTO pcXMLIPs (pcXIP_IPAddr,pcXIP_TurnOn) VALUES ('" & tmpIP & "'," & TurnOn & ");"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
		Next
		
		call closedb()
		response.redirect "AdminManageXMLIPs.asp?msg=added"
	END IF
	
	IF request("submit2")<>"" then
		IPCount=request("IPCount")
		if IPCount>"0" then
			For i=1 to IPCount
				tmpID=request("C" & i)
				if tmpID<>"" then
					query="DELETE FROM pcXMLIPs WHERE pcXIP_id=" & tmpID & ";"
					set rs=connTemp.execute(query)
					set rs=nothing
				end if
			Next
		end if
		
		call closedb()
		response.redirect "AdminManageXMLIPs.asp?msg=deleted"
	END IF
	
	IF request("submit1")<>"" then
		TurnOn=0
	
		query="SELECT pcXIP_id,pcXIP_IPAddr,pcXIP_TurnOn FROM pcXMLIPs;"
		set rs=connTemp.execute(query)
		if rs.eof then
			set rs=nothing
			call closedb()
			response.redirect "AdminManageXMLIPs.asp?msg=cantturn"
		end if
		set rs=nothing
		
		TurnOn=request("TurnXML")
		
		if TurnOn<>"" then
			query="UPDATE pcXMLIPs SET pcXIP_TurnOn=" & TurnOn & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
		
		call closedb()
		if TurnOn="1" then
			response.redirect "AdminManageXMLIPs.asp?msg=turnon"
		end if
		if TurnOn="0" then
			response.redirect "AdminManageXMLIPs.asp?msg=turnoff"
		end if
	END IF
	
END IF	

%>
<!--#include file="AdminHeader.asp"-->
<form name="Form1" action="AdminManageXMLIPs.asp?action=upd" method="post" class="pcForms">
<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		Use this feature to further tighten the security of the XML Tools. When you turn on this feature, only IP addresses listed on this page will be able to transact (send/receive XML data feeds) with your store.
        </td>
	</tr>
	<tr>
		<td class="pcCPspacer">
        
			<% If Request.QueryString("msg")<>"" Then %>
                    
                    <%if Request.QueryString("msg")="added" then%>
                        <div class="pcCPmessageSuccess">The IP Addresses have been added successfully!</div>
                    <%end if%>
                    <%if Request.QueryString("msg")="deleted" then%>
                        <div class="pcCPmessageSuccess">The IP Addresses have been removed successfully!</div>
                    <%end if%>
                    <%if Request.QueryString("msg")="turnon" then%>
                        <div class="pcCPmessageSuccess">This feature has been turned on!</div>
                    <%end if%>
                    <%if Request.QueryString("msg")="turnoff" then%>
                        <div class="pcCPmessageSuccess">This feature has been turned off!</div>
                    <%end if%>
                    <%if Request.QueryString("msg")="cantturn" then%>
                        <div class="pcCPmessage">You can only turn this feature on/off when there are IP Addresses listed.</div>
                    <%end if%>
                    </div>
            <% End If %>
	
        </td>
	</tr>
	
	<%
    TurnOn=0
	
	query="SELECT pcXIP_id,pcXIP_IPAddr,pcXIP_TurnOn FROM pcXMLIPs;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		TurnOn=rs("pcXIP_TurnOn")
	end if
	%>
	
    <tr>
        <th colspan="2">Turn On/Off IP Blocking</th>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
        <td colspan="2"><input type="radio" name="TurnXML" value="1" <%if TurnOn=1 then%>checked<%end if%> class="clearBorder">Turn On&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="TurnXML" value="0" <%if TurnOn=0 then%>checked<%end if%> class="clearBorder">Turn Off
        &nbsp;&nbsp;<input type="submit" name="Submit1" value="Update Setting" class="submit2"></td>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
    <tr>
        <th colspan="2">Allowed IP Addresses</th>
    </tr>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>
				<%IPCount=0
				if not rs.eof then%>
					<tr>
						<td colspan="2">IP Address</td>
					</tr>
					<%pcArr=rs.getRows()
					intCount=ubound(pcArr,2)
					For i=0 to intCount
						IPCount=IPCount+1%>
						<tr>
							<td colspan="2"><input type="checkbox" name="C<%=IPCount%>" value="<%=pcArr(0,i)%>" class="clearBorder">&nbsp;<%=pcArr(1,i)%></td>
						</tr>
					<%Next
				else%>
				<tr>
					<td colspan="2">NO IP Addresses are currently allowed.</td>
				</tr>
				<%end if%>
				<tr>
					<td colspan="2">
						<input type="hidden" name="IPCount" value="<%=IPCount%>">
						<%if IPCount>0 then%>
						<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
						<script language="JavaScript">
						function checkAll()
						{
							for (var j = 1; j <= <%=IPCount%>; j++)
							{
								box = eval("document.Form1.C" + j); 
								if (box.checked == false) box.checked = true;
							}
						}

						function uncheckAll()
						{
							for (var j = 1; j <= <%=IPCount%>; j++)
							{
								box = eval("document.Form1.C" + j); 
								if (box.checked == true) box.checked = false;
							}
						}
						</script>
						<%end if%>
					</td>
				</tr>
				<%if IPCount>0 then%>
				<tr>
					<td colspan="2"><input type="submit" name="Submit2" value="Remove Selected"></td>
				</tr>
				<%end if%>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">Add new IP Addresses</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<%For i=1 to 10%>
				<tr>
					<td colspan="2">IP Address &nbsp;&nbsp;&nbsp;<input type="text" name="N1_<%=i%>" size="4" onblur="javascript:TestNumber(this);">.<input type="text" name="N2_<%=i%>" size="4" onblur="javascript:TestNumber(this);">.<input type="text" name="N3_<%=i%>" size="4" onblur="javascript:TestNumber(this);">.<input type="text" name="N4_<%=i%>" size="4" onblur="javascript:TestNumber(this);"></td>
				</tr>
				<%Next%>
				<tr>
					<td colspan="2" class="pcCPspacer"><hr></td>
				</tr>
				<tr>
					<td colspan="2"><input type="submit" name="Submit3" value="Add new" class="submit2">
                    &nbsp;<input type="button" name="Back" value="XML Tools Manager" onclick="location='XMLToolsManager.asp';"></td>
				</tr>
			</table>
			<script>
				function isDigit(s)
				{
					var test=""+s;
					if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
					{
						return(true) ;
					}
					return(false);
				}
	
				function allDigit(s)
				{
					var test=""+s ;
					for (var k=0; k <test.length; k++)
					{
						var c=test.substring(k,k+1);
						if (isDigit(c)==false)
						{
							return (false);
						}
					}
					return (true);
				}
				
				function TestNumber(fname)
				{
					if (fname.value!="")
					{
						if (allDigit(fname.value) == false)
						{
						    alert("Please enter a numeric value for this Field.");
						    fname.focus();
						    return (false);
					    }
					    if ((fname.value<0) || (fname.value>255))
					    {
					    	alert("Please enter a numeric value between 0-255 for this Field.");
						    fname.focus();
						    return (false);
					    }
					}
				}
			</script>
			</form>
<%call closedb()%>
<!--#include file="AdminFooter.asp"-->