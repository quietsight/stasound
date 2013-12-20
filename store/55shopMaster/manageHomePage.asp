<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Home Page" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% Dim connTemp,rs,query

dim pcMessage, pcStrHPStyle, pcStrHPDesc, pcIntHPFirst, pcIntHPShowSKU, pcIntHPShowImg, pcIntHPFeaturedCount
dim pcIntHPSpcCount, pcIntHPSpcOrder, pcIntHPNewCount, pcIntHPSNewOrder, pcIntHPBestCount, pcIntHPBestOrder

if request("action")="upd" then
	pcIntHPFeaturedCount=request("pcIntHPFeaturedCount")
	pcStrHPStyle=request("pcStrHPStyle")
	pcStrHPDesc=replace(request("pcStrHPDesc"),"'","''")
	pcIntHPFirst=request("pcIntHPFirst")
	if pcIntHPFirst = "" then
		pcIntHPFirst = 0
	else
		pcIntHPFirst = -1
	end if
	pcIntHPShowSKU=request("pcIntHPShowSKU")
		if pcIntHPShowSKU = "" then
			pcIntHPShowSKU = 0
		end if
	pcIntHPShowImg=request("pcIntHPShowImg")
		if pcIntHPShowImg = "" then
			pcIntHPShowImg = 0
		end if
	pcIntHPSpcCount=request("pcIntHPSpcCount")
		if not validNum(pcIntHPSpcCount) then
			pcIntHPSpcCount = 0
		end if
	pcIntHPSpcOrder=request("pcIntHPSpcOrder")
		if not validNum(pcIntHPSpcOrder) then
			pcIntHPSpcOrder = 0
		end if
	pcIntHPNewCount=request("pcIntHPNewCount")
		if not validNum(pcIntHPNewCount) then
			pcIntHPNewCount = 0
		end if
	pcIntHPNewOrder=request("pcIntHPNewOrder")
		if not validNum(pcIntHPNewOrder) then
			pcIntHPNewOrder = 0
		end if
	pcIntHPBestCount=request("pcIntHPBestCount")
		if not validNum(pcIntHPBestCount) then
			pcIntHPBestCount = 0
		end if
	pcIntHPBestOrder=request("pcIntHPBestOrder")
		if not validNum(pcIntHPBestOrder) then
			pcIntHPBestOrder = 0
		end if
	
	call opendb()

	query="SELECT pcHPS_FeaturedCount FROM pcHomePageSettings;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		dim rsUpdObj
		query="UPDATE pcHomePageSettings SET pcHPS_FeaturedCount=" & pcIntHPFeaturedCount & ",pcHPS_Style='" & pcStrHPStyle & "'," &_
		    "pcHPS_PageDesc='" & pcStrHPDesc & "',pcHPS_ShowSKU=" & pcIntHPShowSKU & ",pcHPS_ShowImg=" & pcIntHPShowImg & "," &_
		    "pcHPS_First=" & pcIntHPFirst & ",pcHPS_SpcCount=" & pcIntHPSpcCount & ",pcHPS_SpcOrder=" & pcIntHPSpcOrder & "," &_
		    "pcHPS_NewCount=" & pcIntHPNewCount & ",pcHPS_NewOrder=" & pcIntHPNewOrder & "," &_
		    "pcHPS_BestCount=" & pcIntHPBestCount & ",pcHPS_BestOrder=" & pcIntHPBestOrder & ";"
		set rsUpdObj=server.CreateObject("ADODB.RecordSet")
		set rsUpdObj=connTemp.execute(query)
		set rsUpdObj=nothing
	else
		dim rsInsObj
		query="INSERT INTO pcHomePageSettings (pcHPS_FeaturedCount,pcHPS_Style,pcHPS_PageDesc,pcHPS_First,pcHPS_ShowSKU,pcHPS_ShowImg," &_
		    "pcHPS_SpcCount,pcHPS_SpcOrder,pcHPS_NewCount,pcHPS_NewOrder,pcHPS_BestCount,pcHPS_BestOrder) " &_
		    "VALUES (" & pcIntHPFeaturedCount & ",'" & pcStrHPStyle & "','" & pcStrHPDesc &"'," & pcIntHPFirst &_
		    "," & pcIntHPShowSKU & "," & pcIntHPShowImg & "," & pcIntHPSpcCount & "," & pcIntHPSpcOrder  &_
		    "," & pcIntHPNewCount & "," & pcIntHPNewOrder  &_
		    "," & pcIntHPBestCount & "," & pcIntHPBestOrder & ");"
		set rsInsObj=server.CreateObject("ADODB.RecordSet")
		set rsInsObj=connTemp.execute(query)
		set rsInsObj=nothing
	end if
	
	set rs=nothing
	call closedb()
	msg="Home Page Settings were updated successfully!"
	msgType=1
end if

call opendb()

query=  "SELECT pcHPS_FeaturedCount,pcHPS_Style,pcHPS_PageDesc,pcHPS_First,pcHPS_ShowSKU,pcHPS_ShowImg," &_
        "pcHPS_SpcCount,pcHPS_SpcOrder,pcHPS_NewCount,pcHPS_NewOrder,pcHPS_BestCount,pcHPS_BestOrder FROM pcHomePageSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntHPFeaturedCount=rs("pcHPS_FeaturedCount")
	if pcIntHPFeaturedCount = "" or not validNum(pcIntHPFeaturedCount) then
		pcIntHPFeaturedCount = 14
	end if
	pcStrHPStyle=rs("pcHPS_Style")	
	pcStrHPDesc=replace(rs("pcHPS_PageDesc"),"''","'")
	pcIntHPFirst=rs("pcHPS_First")
	pcIntHPShowSKU=rs("pcHPS_ShowSKU")
	pcIntHPShowImg=rs("pcHPS_ShowImg")
	pcIntHPSpcCount=rs("pcHPS_SpcCount")
	if pcIntHPSpcCount = "" or not validNum(pcIntHPSpcCount) then
		pcIntHPSpcCount = 4
	end if
	pcIntHPSpcOrder=rs("pcHPS_SpcOrder")
	pcIntHPNewCount=rs("pcHPS_NewCount")
	if pcIntHPNewCount = "" or not validNum(pcIntHPNewCount) then
		pcIntHPNewCount = 4
	end if
	pcIntHPNewOrder=rs("pcHPS_NewOrder")
	pcIntHPBestCount=rs("pcHPS_BestCount")
	if pcIntHPBestCount = "" or not validNum(pcIntHPBestCount) then
		pcIntHPBestCount = 4
	end if
	pcIntHPBestOrder=rs("pcHPS_BestOrder")
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
  if (TestField(theForm.pcIntHPFeaturedCount,1)==false)
  {
  	return(false);
  }
  if (TestField(theForm.pcIntHPSpcCount,1)==false)
  {
  	return(false);
  }
  if (TestField(theForm.pcIntHPNewCount,1)==false)
  {
  	return(false);
  }
  if (TestField(theForm.pcIntHPBestCount,1)==false)
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

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form action="manageHomePage.asp?action=upd" method="post" name="manageHomePage" class="pcForms" onsubmit="return Form1_Validator(this)">
	<table class="pcCPcontent">
		<tr>
			<td colspan="2">This page controls ProductCart's <a href="../pc/home.asp" target="_blank">default home page</a>. The page shows a list of <a href="AdminFeatures.asp">featured products</a>, <a href="manageSpecials.asp">specials</a>, <a href="manageBestSellers.asp">best sellers</a>, <a href="manageNewArrivals.asp">new arrivals</a>, and other information, based on your settings.</td>
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
				<span class="pcCPnotes">If empty, no page description is shown at the top of the page</span>			</td>
			<td width="75%">
				<textarea name="pcStrHPDesc" cols="50" rows="6"><%=pcStrHPDesc%></textarea>
				&nbsp;
				<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=pcStrHPDesc&iform=manageHomePage','window2')">
				<br />
			</td>
		</tr>	
		<tr> 
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2">Choose a display option for how <span style="font-weight: bold">featured products</span> are displayed, if any:
				<select name="pcStrHPStyle">
					<option value="" <%if pcStrHPStyle="" then%>selected<%end if%>>Default</option>
					<option value="h" <%if pcStrHPStyle="h" then%>selected<%end if%>>Horizontally</option>
					<option value="p" <%if pcStrHPStyle="p" then%>selected<%end if%>>Vertically</option>
					<option value="l" <%if pcStrHPStyle="l" then%>selected<%end if%>>In a list</option>
					<option value="m" <%if pcStrHPStyle="m" then%>selected<%end if%>>In a list (multiple Add to Cart)</option>
				</select>
				&nbsp;<a href="JavaScript:;" onClick="JavaScript:window.open('images/pcv3_displayOptions.gif','','scrollbars=yes,status=no,width=640,height=700')"><img src="images/pcv3_infoIcon.gif" border="0"></a>
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<input type="checkbox" name="pcIntHPFirst" <%if pcIntHPFirst<>0 then%>checked<%end if%> value="0" class="clearBorder"> Highlight the first featured product. &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=203')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>
		<tr> 
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2"> 
			The following two settings apply to how <span style="font-weight: bold">best sellers</span>, <span style="font-weight: bold">new arrivals</span>, and <span style="font-weight: bold">specials</span> are shown on the page.			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show the product part number (SKU)? <input type="radio" name="pcIntHPShowSKU" value="-1" <%If pcIntHPShowSKU="-1" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntHPShowSKU" value="0" <%If pcIntHPShowSKU="0" then%> checked<% end if%> class="clearBorder">No			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show the small product thumbnail? <input type="radio" name="pcIntHPShowImg" value="-1" <%If pcIntHPShowImg="-1" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntHPShowImg" value="0" <%If pcIntHPShowImg="0" then%> checked<% end if%> class="clearBorder">No	
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
			<td colspan="2">For any of the product sections listed below, enter 0 to hide that section. For example, if you don't want to show &quot;specials&quot; on the home page, enter 0 in the corresponding input field. Use the &quot;Order&quot; field to set the order in which the sections should be shown.</td>
		</tr>
		<tr>
			<td colspan="2">
				<table class="pcCPcontent">
					<tr>
						<td>&nbsp;</td>
						<td>Number of items to show</td>
						<td>Order</td>
					</tr>
					<tr>
						<td><a href="AdminFeatures.asp">Featured products</a></td>
						<td><input type="text" name="pcIntHPFeaturedCount" size="4" value="<%=pcIntHPFeaturedCount%>"></td>
						<td>Featured products are shown above the other sections</td>
					</tr>
					<tr>
						<td><a href="manageSpecials.asp">Specials</a></td>
						<td><input type="text" name="pcIntHPSpcCount" size="4" value="<%=pcIntHPSpcCount%>"></td>
						<td><input type="text" name="pcIntHPSpcOrder" size="4" value="<%=pcIntHPSpcOrder%>"></td>
					</tr>
					<tr>
						<td><a href="manageNewArrivals.asp">New Arrivals</a></td>
						<td><input type="text" name="pcIntHPNewCount" size="4" value="<%=pcIntHPNewCount%>"></td>
						<td><input type="text" name="pcIntHPNewOrder" size="4" value="<%=pcIntHPNewOrder%>"></td>
					</tr>
					<tr>
						<td><a href="manageBestSellers.asp">Best Sellers</a></td>
						<td><input type="text" name="pcIntHPBestCount" size="4" value="<%=pcIntHPBestCount%>"></td>
						<td><input type="text" name="pcIntHPBestOrder" size="4" value="<%=pcIntHPBestOrder%>"></td>
					</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> 
			<input type="submit" name="modify" value="Update Settings" class="submit2">
            &nbsp;<input type="button" name="back" value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->