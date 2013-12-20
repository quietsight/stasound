<%
if (pcgwTransId<>"" AND isNULL(pcgwTransId)=False) then
%>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">NetSource Commerce Gateway</th>
	</tr>	
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>	
	<tr>
		<td colspan="2">This section includes payment-related tasks that you can perform with the EI Gateway.&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=700')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></td>
	</tr>	
    
	<%
    EIGbtns=0
    'if pcv_PaymentStatus="1" OR pcv_PaymentStatus="2" then %>
    <tr>
        <td colspan="2">
        <div style="padding-bottom:4px;">
            <input type="hidden" name="EIGTransID" value="<%=pcgwTransId%>">
            <input type="hidden" name="EIGTransParentID" value="<%=pcgwTransParentId%>">

            <% if pcv_PaymentStatus<>6 AND pcv_PaymentStatus<>8 AND pcv_PaymentStatus<>1 then
                EIGbtns=1 
                %>					
                <input type="submit" name="SubmitEIG1" value=" Refund "  onClick="javascript: if (confirm('This action will NOT cancel the order, but rather refund the payment via the NetSource Commerce Gateway. You can separately cancel the order. Are you sure you want to continue?')) return true ; else return false ;" class="submit2">&nbsp;&nbsp;					
            <% end if %>
        </div>
        
        
        <%if EIGbtns=1 then%>							
        <%else%>
            <div style="padding-bottom:4px;">						
				<% if pcv_PaymentStatus=1 and porderStatus="2" then %>
                    <a href='batchprocessorders.asp'>Batch process this order</a> to process it &amp; capture funds at the same time.
				<% else %>
                    <em>No payment-related task available for this order.</em>
                <% end if %>
            </div>
        <%end if%>	
        
        
        </td>
    </tr>
    <%' end if %>
    <tr>
        <td colspan="2" class="pcCPspacer"></td>
    </tr>	

<%
end if 
%>

<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'// END: EIG - Display Risk Managment if its available.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>