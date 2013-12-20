<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Recently Reviewed Products" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% 
Dim connTemp,rs,query
' PRV41 start
Dim pcMessage, pcIntRecentRevCount, pcStrPageStyle, pcStrRecentRevDesc, pcIntRevDays, pcIntRecentRevNFS, pcIntRecentRevInStock, pcIntShowSKU, pcIntShowSmallImg, pcintReviewsPerProduct
' PRV41 end

if request("action")="upd" then
	pcIntRecentRevCount=request("pcIntRecentRevCount")
		if not validNum(pcIntRecentRevCount) or pcIntRecentRevCount=0  then
			pcIntRecentRevCount=14
		end If
    ' PRV41 begin
	pcIntReviewsPerProduct=request("pcintReviewsPerProduct")
		if not validNum(pcintReviewsPerProduct) or pcintReviewsPerProduct=0  then
			pcintReviewsPerProduct=3
		end if

	' PRV41 end
	pcStrPageStyle=request("pcStrPageStyle")
		if pcStrPageStyle="" then
			pcStrPageStyle=LCase(bType)
		end if
	pcStrRecentRevDesc=replace(request("pcStrRecentRevDesc"),"'","''")
	pcIntRevDays=request("pcIntRevDays")
		if not validNum(pcIntRevDays) or pcIntRevDays=0 then
			pcIntRevDays=30
		end if
	pcIntRecentRevNFS=request("pcIntRecentRevNFS")
		if pcIntRecentRevNFS="" then
			pcIntRecentRevNFS="-1"
		end if
	pcIntRecentRevInStock=request("pcIntRecentRevInStock")
		if pcIntRecentRevInStock="" then
			pcIntRecentRevInStock="-1"
		end if
	pcIntShowSKU=request("pcIntShowSKU")
		if pcIntShowSKU="" then
			pcIntShowSKU="-1"
		end if
	pcIntShowSmallImg=request("pcIntShowSmallImg")
		if pcIntShowSmallImg="" then
			pcIntShowSmallImg="-1"
		end if
	
	call opendb()

	query="SELECT pcRR_RecentRevCount FROM pcRecentRevSettings;"
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		dim rsUpdObj
		' PRV41 begin
		query="UPDATE pcRecentRevSettings SET pcRR_RecentRevCount=" & pcIntRecentRevCount & ",pcRR_Style='" & pcStrPageStyle & "',pcRR_PageDesc='" & pcStrRecentRevDesc & "',pcRR_RevDays=" & pcIntRevDays & ",pcRR_NotForSale=" & pcIntRecentRevNFS & ",pcRR_OutOfStock=" & pcIntRecentRevInStock & ",pcRR_SKU=" & pcIntShowSKU & ",pcRR_ShowImg=" & pcIntShowSmallImg & ", pcRR_ReviewsPerProduct=" & pcintReviewsPerProduct & ";"
		' PRV41 end
		set rsUpdObj=Server.CreateObject("ADODB.RecordSet")
		set rsUpdObj=connTemp.execute(query)
		set rsUpdObj=nothing
	else
		dim rsInsObj
		' PRV41 begin
		query="INSERT INTO pcRecentRevSettings (pcRR_RecentRevCount, pcRR_Style, pcRR_PageDesc, pcRR_RevDays, pcRR_NotForSale, pcRR_OutOfStock, pcRR_SKU, pcRR_ShowImg, pcRR_ReviewsPerProduct) VALUES (" & pcIntRecentRevCount & ", '" & pcStrPageStyle & "', '" & pcStrRecentRevDesc & "', " & pcIntRevDays & ", " & pcIntRecentRevNFS & ", " & pcIntRecentRevInStock & ", " & pcIntShowSKU & ", " & pcIntShowSmallImg & "," & pcintReviewsPerProduct & ");"
		' PRV41 end
		set rsInsObj=Server.CreateObject("ADODB.RecordSet")
		set rsInsObj=connTemp.execute(query)
		set rsInsObj=nothing
	end if
	
	set rs=nothing
	call closedb()
	
	msg="Recently Reviewed Products Page Settings were updated successfully!"
	msgType=1
end if

call opendb()

query="SELECT pcRR_RecentRevCount, pcRR_Style, pcRR_PageDesc, pcRR_RevDays, pcRR_NotForSale, pcRR_OutOfStock, pcRR_SKU, pcRR_ShowImg FROM pcRecentRevSettings;"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntRecentRevCount=rs("pcRR_RecentRevCount")
		if pcIntRecentRevCount = "" or not validNum(pcIntRecentRevCount) then
			pcIntRecentRevCount = 14
		end If
	' PRV41 begin
	pcintReviewsPerProduct=rs("pcRR_ReviewsPerProduct")
		if pcintReviewsPerProduct = "" or not validNum(pcintReviewsPerProduct) then
			pcintReviewsPerProduct = 3
		end if

	' PRV41 end
	pcStrPageStyle=rs("pcRR_Style")
	pcStrRecentRevDesc=rs("pcRR_PageDesc")
	pcIntRevDays=rs("pcRR_RevDays")
		if pcIntRevDays = "" or not validNum(pcIntRevDays) then
			pcIntRevDays = 30
		end if
	pcIntRecentRevNFS=rs("pcRR_NotForSale")
	pcIntRecentRevInStock=rs("pcRR_OutOfStock")
	pcIntShowSKU=rs("pcRR_SKU")
	pcIntShowSmallImg=rs("pcRR_ShowImg")
end if

set rs=nothing
call closedb()
%>

<script language="JavaScript">
<!--
	
function isDigit(s)
{
var test=""+s;
if(test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
}

function isDigit1(s)
{
var test=""+s;
if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
}
	
function allDigit(s,stype)
{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (stype==1)
		{
			if (isDigit1(c)==false)
			{
				return (false);
			}
		}
		else
		{
			if (isDigit(c)==false)
			{
				return (false);
			}		
		}
	}
	return (true);
}
	
function TestField(theField,stype)
{
	if (theField.value == "")
	  	{
	    alert("Please enter a value for this field!");
	    theField.focus();
	    return (false);
		}
		else
		{
			if (allDigit(theField.value,stype) == false)
			{
				if (stype==1)
				{
					alert("Please enter a positive integer value for this field.");
				}
				else
				{
			    	alert("Please enter a numeric value ##.## for this field.");
			    }
		    theField.focus();
		    return (false);
		    }
	    }
}

function Form1_Validator(theForm)
{
  if (TestField(theForm.pcIntRecentRevCount,1)==false)
  {
  	return(false);
  }
  // PRV41 begin
  if (TestField(theForm.pcIntReviewsPerProduct,1)==false)
  {
  	return(false);
  }

  // PRV41 end
  if (TestField(theForm.pcIntRevDays,1)==false)
  {
  	return(false);
  }
return (true);
}
	function newWindow(file,window) {
		msgWindow=open(file,window,'resizable=no,width=400,height=500');
		if (msgWindow.opener == null) msgWindow.opener = self;
	}
// -->
</script>
	 
<form action="manageRecentlyReviewed.asp?action=upd" method="post" name="manageRecentRev" class="pcForms" onsubmit="return Form1_Validator(this)">
<% 'PRV41 start %>
<input type="hidden" name="pcStrPageStyle" value="l">
<% 'PRV41 end %>

	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<td colspan="2">This page controls the way <a href="../pc/showRecentlyReviewed.asp" target="_blank">Recently Reviewed Products</a> are shown in the storefront.</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Display Settings</th>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr valign="top">
			<td width="25%">
			Page Description:
				<br />
				<br />
				<span class="pcCPnotes">If empty, no page description is shown at the top of the page</span>
			</td>
			<td width="75%">
				<textarea name="pcStrRecentRevDesc" cols="50" rows="6"><%=pcStrRecentRevDesc%></textarea>
				&nbsp;
				<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=pcStrRecentRevDesc&iform=manageRecentRev','window2')">
				<br />
			</td>
		</tr>	
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show the product part number (SKU)? <input type="radio" name="pcIntShowSKU" value="-1" <%If pcIntShowSKU<>"0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntShowSKU" value="0" <%If pcIntShowSKU="0" then%> checked<% end if%> class="clearBorder">No
			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show the small product thumbnail? <input type="radio" name="pcIntShowSmallImg" value="-1" <%If pcIntShowSmallImg<>"0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntShowSmallImg" value="0" <%If pcIntShowSmallImg="0" then%> checked<% end if%> class="clearBorder">No
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="2">Products to Show</th>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">
			Number of products to show:&nbsp;
			<input type="text" name="pcIntRecentRevCount" size="5" value="<%=pcIntRecentRevCount%>"></td>
		</tr>
		<% 'PRV41 start %>
		<tr> 
			<td colspan="2">
			Number of reviews to show per product:&nbsp;
			<input type="text" name="pcIntReviewsPerProduct" size="5" value="<%=pcintReviewsPerProduct%>"></td>
		</tr>
		<% 'PRV41 end %>
		<tr> 
			<td colspan="2">
			Only include products for which a review has been written in the last <input type="text" name="pcIntRevDays" size="5" value="<%=pcIntRevDays%>"> days</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show &quot;Not for Sale&quot; products? <input type="radio" name="pcIntRecentRevNFS" value="0" <%If pcIntRecentRevNFS="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntRecentRevNFS" value="-1" <%If pcIntRecentRevNFS<>"0" then%> checked<% end if%> class="clearBorder">No
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=202')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show &quot;Out of Stock&quot; products? <input type="radio" name="pcIntRecentRevInStock" value="0" <%If pcIntRecentRevInStock="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntRecentRevInStock" value="-1" <%If pcIntRecentRevInStock<>"0" then%> checked<% end if%> class="clearBorder">No
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=201')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>

		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
			<input type="submit" name="modify" value="Update Settings" class="submit2">
            &nbsp;
            <input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->