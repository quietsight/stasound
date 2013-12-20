<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle="Product Reviews - Manage Bad Words"
pageIcon="pcv4_icon_reviews.png"
section="reviews" 
%>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
Dim rs, connTemp, strSQL, pid

call openDb()

IF request("action")="add" then
	query="DELETE FROM pcRevBadWords"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	pcv_Values=request("pcv_values")
	if pcv_Values<>"" then
		pcArray=split(pcv_Values,vbcrlf)
		For k=lbound(pcArray) to ubound(pcArray)
			if trim(pcArray(k))<>"" then
				pcArray(k)=trim(pcArray(k))
				query="INSERT INTO pcRevBadWords (pcRBW_word) VALUES ('" & pcArray(k) & "');"
				set rs=server.CreateObject("ADODB.RecordSet")
				set rs=connTemp.execute(query)
			end if
		Next
	end if
	msg="The list of 'Bad Words' was updated successfully!"
	msgType=1
	set rs=nothing
END IF

query="SELECT pcRBW_word FROM pcRevBadWords"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

pcv_strValues=""
if not rs.eof then
	pcArray=rs.getRows()
	intCount=ubound(pcArray,2)
	for k=0 to intCount
		pcv_strValues=pcv_strValues & pcArray(0,k) & vbcrlf
	next
end if
set rs=nothing

call closedb()
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>


<form method="POST" action="prv_ManageBadWords.asp?action=add" name="checkboxform" class="pcForms">
	
    <table class="pcCPcontent">
        <tr>                    
        	<td>Enter a list of words that are not allowed when customers write product reviews. If a customer includes any of these words in a review, the words are automatically replaced with ****. <span class="pcCPnotes" style="margin-top: 10px;">Enter one word per row.</span></td>
        </tr>
        <tr>                    
            <td class="pcCPspacer"></td>
        </tr>
        <tr>
            <td>
                <textarea rows="15" cols="25" name="pcv_values"><%=pcv_strValues%></textarea>
            </td>
        </tr>
        <tr>                    
            <td class="pcCPspacer"></td>
        </tr>
        <tr>
            <td>
            	<input type="submit" value="Update List" class="submit2">&nbsp;
                <input type="button" value="Back" onClick="javascript: history.back();">
            </td>
        </tr>
        <tr>
            <td>
            <input type="button" value="Product Reviews Settings" onClick="location.href='PrvSettings.asp'">&nbsp;
            <input type="button" value="View Pending Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=1'">&nbsp;
            <input type="button" value="View Live Reviews" onClick="location.href='prv_ManageRevPrds.asp?nav=2'">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->