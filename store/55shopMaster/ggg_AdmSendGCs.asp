<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.buffer=true %>
<% pageTitle="Send Gift Certificate Code" %>
<% Section="products" %>
<%PmAdmin="1*2*3*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/emailsettings.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->
<%
msg=""
if request("action")="post" then
	pemail=request("toemail")
	msubject=request("subject")
	mbody=request("message")
	call sendmail (scCompanyName, scEmail, pemail, msubject, mbody)
	msg="Message has been sent successfully!"
end if
%>
<script>
function validate_email(field,alerttxt)
{
with (field)
  {
  apos=value.indexOf("@");
  dotpos=value.lastIndexOf(".");
  if (apos<1||dotpos-apos<2)
	{alert(alerttxt);return false;}
  else {return true;}
  }
}
	
function Form1_Validator(theForm)
{
	if (theForm.toemail.value=="")
	{
		alert("Please enter an e-mail address.");
	    theForm.toemail.focus();
	    return (false);
	}

	if (validate_email(theForm.toemail,"Please enter a valid e-mail address.")==false)
	{
		theForm.toemail.focus();
		return (false);
	}
	
	if (theForm.subject.value=="")
	{
		alert("Please enter a subject for the message.");
	    theForm.subject.focus();
	    return (false);
	}
	
	if (theForm.message.value=="")
	{
		alert("Please enter a message.");
	    theForm.message.focus();
	    return (false);
	}
  	return (true);
}
</script>
<form name="form1" method="post" action="ggg_AdmSendGCs.asp?action=post" onSubmit="return Form1_Validator(this)" class="pcForms">
	<%if msg<>"" then%>
        <table class="pcCPcontent">
            <tr>
                <td>
                    <div class="pcCPmessageSuccess"><%=msg%></div>
                </td>
            </tr>
        </table>
    <%end if%>
    <table class="pcCPcontent">
        <tr>
            <td>To:</td>
            <td>
                <input name="toemail" type="text" value="" size="50">
            </td>
        </tr>
        <tr>
            <td>Subject:</td>
            <td>
                <input name="subject" type="text" value="Here is your Gift Certificate" size="50">
            </td>
        </tr>
        <tr>
            <td valign="top">Message:</td>
			<%
            tmp_msg="To redeem a Gift Certificate, enter the gift certificate code on the order verification page and click on the 'Recalculate' button." & vbcrlf & vbcrlf
            if request("mode")="all" and session("adm_generated_gcs")<>"" then
                tmpArr=split(session("adm_generated_gcs"),"***")
                For i=lbound(tmpArr) to ubound(tmpArr)
                    if tmpArr(i)<>"" then
                        tmp_msg=tmp_msg & "Gift Certificate #" & cint(i+1) & ": " & tmpArr(i) & vbcrlf
                    end if
                Next
                session("adm_generated_gcs")=""
            else
                if request("GcCode")<>"" then
                    tmp_msg=tmp_msg & "Gift Certificate: " & request("GcCode")
                end if
            end if%>
            <td>
                <textarea name="message" cols="50" rows="15"><%=tmp_msg%></textarea>
            </td>
        </tr>
        <tr>
            <td>&nbsp;</td> 
            <td>
                <br>
                <input type="submit" name="Submit" value=" Send message " class="submit2">
                &nbsp;<input type="button" name="Button" value=" Back " onClick="location='<%if request("mode")="all" then%>ggg_AdmManageGCs.asp<%else%>ggg_AdmSrcGCb.asp<%end if%>';">
            </td>
        </tr>
    </table>
</form>
<!--#include file="AdminFooter.asp"-->