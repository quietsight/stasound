<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage New Arrivals" %>
<% section="specials" %>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="AdminHeader.asp"-->

<% 
Dim connTemp,rs,query
Dim pcMessage, pcIntNewArrCount, pcStrPageStyle, pcStrNewArrDesc, pcIntNDays, pcIntNewArrNFS, pcIntNewArrInStock, pcShowSKU, pcShowSmallImg

if request("action")="upd" then
	pcIntNewArrCount=request("pcIntNewArrCount")
	pcStrPageStyle=request("pcStrPageStyle")
	pcStrNewArrDesc=replace(request("pcStrNewArrDesc"),"'","''")
	pcIntNDays=request("pcIntNDays")
	pcIntNewArrNFS=request("pcIntNewArrNFS")
	pcIntNewArrInStock=request("pcIntNewArrInStock")
	pcIntShowSKU=request("pcIntShowSKU")
	pcIntShowSmallImg=request("pcIntShowSmallImg")
	
	call opendb()

	query="SELECT pcNAS_NewArrCount FROM pcNewArrivalsSettings;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		dim rsUpdObj
		query="UPDATE pcNewArrivalsSettings SET pcNAS_NewArrCount=" & pcIntNewArrCount & ",pcNAS_Style='" & pcStrPageStyle & "',pcNAS_PageDesc='" & pcStrNewArrDesc & "',pcNAS_NDays=" & pcIntNDays & ",pcNAS_NotForSale=" & pcIntNewArrNFS & ",pcNAS_OutOfStock=" & pcIntNewArrInStock & ",pcNAS_SKU=" & pcIntShowSKU & ",pcNAS_ShowImg=" & pcIntShowSmallImg & ";"
		set rsUpdObj=server.CreateObject("ADODB.RecordSet")
		set rsUpdObj=connTemp.execute(query)
		set rsUpdObj=nothing
	else
		dim rsInsObj
		query="INSERT INTO pcNewArrivalsSettings (pcNAS_NewArrCount,pcNAS_Style,pcNAS_PageDesc,pcNAS_NDays,pcNAS_NotForSale,pcNAS_OutOfStock,pcNAS_SKU,pcNAS_ShowImg) VALUES (" & pcIntNewArrCount & ",'" & pcStrPageStyle & "','" & pcStrNewArrDesc & "'," & pcIntNDays & "," & pcIntNewArrNFS & "," & pcIntNewArrInStock & "," & pcIntShowSKU & "," & pcIntShowSmallImg & ");"
		set rsInsObj=server.CreateObject("ADODB.RecordSet")
		set rsInsObj=connTemp.execute(query)
		set rsInsObj=nothing
	end if
	
	set rs=nothing
	call closedb()
	msg="New Arrivals Page Settings were updated successfully!"
	msgtype=1
end if

call opendb()

query="SELECT pcNAS_NewArrCount, pcNAS_Style,pcNAS_PageDesc, pcNAS_NDays, pcNAS_NotForSale, pcNAS_OutOfStock, pcNAS_SKU, pcNAS_ShowImg FROM pcNewArrivalsSettings;"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=connTemp.execute(query)

if not rs.eof then
	pcIntNewArrCount=rs("pcNAS_NewArrCount")
		if pcIntNewArrCount = "" or not validNum(pcIntNewArrCount) then
			pcIntNewArrCount = 14
		end if
	pcStrPageStyle=rs("pcNAS_Style")
	pcStrNewArrDesc=rs("pcNAS_PageDesc")
	pcIntNDays=rs("pcNAS_NDays")
		if pcIntNDays = "" or not validNum(pcIntNDays) then
			pcIntNDays = 30
		end if
	pcIntNewArrNFS=rs("pcNAS_NotForSale")
	pcIntNewArrInStock=rs("pcNAS_OutOfStock")
	pcIntShowSKU=rs("pcNAS_SKU")
	pcIntShowSmallImg=rs("pcNAS_ShowImg")
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
  if (TestField(theForm.pcIntNewArrCount,1)==false)
  {
  	return(false);
  }
  if (TestField(theForm.pcIntNDays,1)==false)
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
	 
<form action="manageNewArrivals.asp?action=upd" method="post" name="manageNewArrivals" class="pcForms" onsubmit="return Form1_Validator(this)">
	<table class="pcCPcontent">
        <tr>
            <td colspan="2" class="pcCPspacer">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>
            </td>
        </tr>
		<tr>
			<td colspan="2">This page controls the way <a href="../pc/shownewarrivals.asp" target="_blank">New Arrivals</a> are shown in the storefront.</td>
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
				<textarea name="pcStrNewArrDesc" cols="50" rows="6"><%=pcStrNewArrDesc%></textarea>
				&nbsp;
				<input type="button" value="Use HTML Editor" onClick="newWindow('pop_HtmlEditor.asp?fi=pcStrNewArrDesc&iform=manageNewArrivals','window2')">
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
			<input type="text" name="pcIntNewArrCount" size="5" value="<%=pcIntNewArrCount%>"></td>
		</tr>

		<tr> 
			<td colspan="2">
			Only include products that have been added in the last <input type=text name="pcIntNDays" size="5" value="<%=pcIntNDays%>"> days</td>
		</tr>
			<td colspan="2"> 
			Show &quot;Not for Sale&quot; products? <input type="radio" name="pcIntNewArrNFS" value="0" <%If pcIntNewArrNFS="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntNewArrNFS" value="-1" <%If pcIntNewArrNFS<>"0" then%> checked<% end if%> class="clearBorder">No
			&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=202')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>
			</td>
		</tr>
		<tr> 
			<td colspan="2"> 
			Show &quot;Out of Stock&quot; products? <input type="radio" name="pcIntNewArrInStock" value="0" <%If pcIntNewArrInStock="0" then%> checked<% end if%> class="clearBorder">Yes <input type="radio" name="pcIntNewArrInStock" value="-1" <%If pcIntNewArrInStock<>"0" then%> checked<% end if%> class="clearBorder">No
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