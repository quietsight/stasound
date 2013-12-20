<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Best Sellers" %>
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
Dim pcMessage, pcIntBestSellCount, pcStrPageStyle, pcStrBestSellDesc, pcIntNSold, pcIntBestSellNFS, pcIntBestSellInStock, pcIntShowSKU, pcIntShowSmallImg

if request("action")="upd" then
	pcIntBestSellCount=request("pcIntBestSellCount")
		if not validNum(pcIntBestSellCount) or pcIntBestSellCount=0  then
			pcIntBestSellCount=14
		end if
	pcStrPageStyle=request("pcStrPageStyle")
		if pcStrPageStyle="" then
			pcStrPageStyle=LCase(bType)
		end if
	pcStrBestSellDesc=replace(request("pcStrBestSellDesc"),"'","''")
	pcIntNSold=request("pcIntNSold")
		if not validNum(pcIntNSold) or pcIntNSold=0 then
			pcIntNSold=2
		end if
	pcIntBestSellNFS=request("pcIntBestSellNFS")
		if pcIntBestSellNFS="" then
			pcIntBestSellNFS="-1"
		end if
	pcIntBestSellInStock=request("pcIntBestSellInStock")
		if pcIntBestSellInStock="" then
			pcIntBestSellInStock="-1"
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

	query="SELECT pcBSS_BestSellCount FROM pcBestSellerSettings;"
	set rs=Server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		dim rsUpdObj
		query="UPDATE pcBestSellerSettings SET pcBSS_BestSellCount=" & pcIntBestSellCount & ",pcBSS_Style='" & pcStrPageStyle & "',pcBSS_PageDesc='" & pcStrBestSellDesc & "',pcBSS_NSold=" & pcIntNSold & ",pcBSS_NotForSale=" & pcIntBestSellNFS & ",pcBSS_OutOfStock=" & pcIntBestSellInStock & ",pcBSS_SKU=" & pcIntShowSKU & ",pcBSS_ShowImg=" & pcIntShowSmallImg & ";"
		set rsUpdObj=Server.CreateObject("ADODB.RecordSet")
		set rsUpdObj=connTemp.execute(query)
		set rsUpdObj=nothing
	else
		dim rsInsObj
		query="INSERT INTO pcBestSellerSettings (pcBSS_BestSellCount, pcBSS_Style, pcBSS_PageDesc, pcBSS_NSold, pcBSS_NotForSale, pcBSS_OutOfStock, pcBSS_SKU, pcBSS_ShowImg) VALUES (" & pcIntBestSellCount & ", '" & pcStrPageStyle & "', '" & pcStrBestSellDesc & "', " & pcIntNSold & ", " & pcIntBestSellNFS & ", " & pcIntBestSellInStock & ", " & pcIntShowSKU & ", " & pcIntShowSmallImg & ");"
		set rsInsObj=Server.CreateObject("ADODB.RecordSet")
		set rsInsObj=connTemp.execute(query)
		set rsInsObj=nothing
	end if
	
	set rs=nothing
	call closedb()
	
	msg="Best Sellers Page Settings were updated successfully!"
	msgType=1
end if

call opendb()

query="SELECT pcBSS_BestSellCount, pcBSS_Style, pcBSS_PageDesc, pcBSS_NSold, pcBSS_NotForSale, pcBSS_OutOfStock, pcBSS_SKU, pcBSS_ShowImg FROM pcBestSellerSettings;"
set rs=Server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntBestSellCount=rs("pcBSS_BestSellCount")
		if pcIntBestSellCount = "" or not validNum(pcIntBestSellCount) then
			pcIntBestSellCount = 14
		end if
	pcStrPageStyle=rs("pcBSS_Style")
	pcStrBestSellDesc=rs("pcBSS_PageDesc")
	pcIntNSold=rs("pcBSS_NSold")
		if pcIntNSold = "" or not validNum(pcIntNSold) then
			pcIntNSold = 2
		end if
	pcIntBestSellNFS=rs("pcBSS_NotForSale")
	pcIntBestSellInStock=rs("pcBSS_OutOfStock")
	pcIntShowSKU=rs("pcBSS_SKU")
	pcIntShowSmallImg=rs("pcBSS_ShowImg")
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
  if (TestField(theForm.pcIntBestSellCount,1)==false)
  {
  	return(false);
  }
  if (TestField(theForm.pcIntNSold,1)==false)
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
	 
<form action="manageBestSellers.asp?action=upd" method="post" name="manageBestSellers" class="pcForms" onsubmit="return Form1_Validator(this)">
	<table class="pcCPcontent">
        <tr>
            <td colspan="3" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<td colspan="2">This page controls the way <a href="../pc/showbestsellers.asp" target="_blank">Best Sellers</a> are shown in the storefront.</td>
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
				<textarea name="pcStrBestSellDesc" cols="50" rows="6"><%=pcStrBestSellDesc%></textarea>
				&nbsp;
				<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=pcStrBestSellDesc&iform=manageBestSellers','window2')">
				<br />
			</td>
		</tr>	
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2">Choose a display option for how products are displayed, if any:
				<select name="pcStrPageStyle">
					<option value="" <%if pcStrPageStyle="" then%>selected<%end if%>>Default</option>
					<option value="h" <%if pcStrPageStyle="h" then%>selected<%end if%>>Horizontally</option>
					<option value="p" <%if pcStrPageStyle="p" then%>selected<%end if%>>Vertically</option>
					<option value="l" <%if pcStrPageStyle="l" then%>selected<%end if%>>In a list</option>
					<option value="m" <%if pcStrPageStyle="m" then%>selected<%end if%>>In a list (multiple Add to Cart)</option>
				</select>
				&nbsp;<a href="JavaScript:;" onClick="JavaScript:window.open('images/pcv3_displayOptions.gif','','scrollbars=yes,status=no,width=640,height=700')"><img src="images/pcv3_infoIcon.gif" border="0"></a>
			</td>
		</tr>
		<tr> 
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td colspan="2"> 
			The following two settings only apply when products are shown in a list.
			</td>
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
			<input type="text" name="pcIntBestSellCount" size="5" value="<%=pcIntBestSellCount%>"></td>
		</tr>
		<tr> 
			<td colspan="2">
			Only include products of which you have sold at least <input type="text" name="pcIntNSold" size="5" value="<%=pcIntNSold%>"> units</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show &quot;Not for Sale&quot; products? <input type="radio" name="pcIntBestSellNFS" value="0" <%If pcIntBestSellNFS="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntBestSellNFS" value="-1" <%If pcIntBestSellNFS<>"0" then%> checked<% end if%> class="clearBorder">No
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=202')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show &quot;Out of Stock&quot; products? <input type="radio" name="pcIntBestSellInStock" value="0" <%If pcIntBestSellInStock="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntBestSellInStock" value="-1" <%If pcIntBestSellInStock<>"0" then%> checked<% end if%> class="clearBorder">No
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