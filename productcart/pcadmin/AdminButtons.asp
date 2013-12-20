<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Edit Store Buttons" %>
<% Section="layout" %>
<%PmAdmin=1%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp"-->

<% dim query, conntemp, rstemp
on error resume next
call openDb()
query="SELECT recalculate, continueshop, checkout, submit, morebtn, viewcartbtn, checkoutbtn, addtocart, addtowl, register, cancel, remove, add2, login, login_checkout, back, register_checkout"
'BTO ADDON-S
if scBTO=1 then
	query=query&", customize, [reconfigure], resetdefault, savequote,revorder,submitquote,pcLO_requestQuote"
end if
'BTO ADDON-E
query=query&", ID,pcLO_placeOrder,pcLO_checkoutWR,pcLO_processShip,pcLO_finalShip,pcLO_backtoOrder,pcLO_Previous,pcLO_Next,CreRegistry,DelRegistry,AddToRegistry,UpdRegistry,SendMsgs,RetRegistry,pcLO_Update, pcLO_Savecart FROM layout WHERE (((ID)=2));"

set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	response.write "Error in AdminButtons: "&Err.Description
		set rstemp=nothing
		call closeDb()
end If

precalculate=rstemp("recalculate")
pcontinueshop=rstemp("continueshop")
pcheckout=rstemp("checkout")
psubmit=rstemp("submit")
pmorebtn=rstemp("morebtn")
pviewcartbtn=rstemp("viewcartbtn")
pcheckoutbtn=rstemp("checkoutbtn")
paddtocart=rstemp("addtocart")
paddtowl=rstemp("addtowl")
pregister=rstemp("register")
pcancel=rstemp("cancel")
premove=rstemp("remove")
padd2=rstemp("add2")
plogin=rstemp("login")
plogin_checkout=rstemp("login_checkout")
pback=rstemp("back")
pregister_checkout=rstemp("register_checkout")
'BTO ADDON-S
If scBTO=1 then
	pcustomize=rstemp("customize")
	preconfigure=rstemp("reconfigure")
	presetdefault=rstemp("resetdefault")
	psavequote=rstemp("savequote")
	prevorder=rstemp("revorder")
	psubmitquote=rstemp("submitquote")
	pcv_requestQuote=rstemp("pcLO_requestQuote")
End If
'BTO ADDON-E
pcv_placeOrder=rstemp("pcLO_placeOrder")
pcv_checkoutWR=rstemp("pcLO_checkoutWR")
pcv_processShip=rstemp("pcLO_processShip")
pcv_finalShip=rstemp("pcLO_finalShip")
pcv_backtoOrder=rstemp("pcLO_backtoOrder")
pcv_previous=rstemp("pcLO_Previous")
pcv_next=rstemp("pcLO_Next")

'GGG Add-on start

	pcrereg=rstemp("CreRegistry")
	pdelreg=rstemp("DelRegistry")
	paddreg=rstemp("AddToRegistry")
	pupdreg=rstemp("UpdRegistry")
	psendmsgs=rstemp("SendMsgs")
	pretreg=rstemp("RetRegistry")

'GGG Add-on end

yellowupd=rstemp("pcLO_Update")
pcv_strSaveCart=rstemp("pcLO_Savecart")

set rstemp=nothing
call closedb()

%>
<!--#include file="AdminHeader.asp"-->
<form method="post" enctype="multipart/form-data" action="buttonupl.asp" class="pcForms">
	<table class="pcCPcontent">
	<tr>
		<td colspan="3" class="pcCPspacer">
			<!--#include file="pcv4_showMessage.asp"-->
		</td>
	</tr>
	<tr>
		<th colspan="2">Browse and Upload New Buttons&nbsp;<a href="JavaScript:win('helpOnline.asp?ref=435')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a></th>
		<th width="31%" align="center">Current Buttons</th>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td width="25%">Small Add to Cart:</td>
		<td width="44%">
			<input class=ibtng type="file" name="add2" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=padd2%>"></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Add To Cart:</td>
		<td width="44%">
			<input class=ibtng type="file" name="addtocart" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=paddtocart%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">View Cart:</td>
		<td width="44%">
			<input class=ibtng type="file" name="viewcartbtn" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pviewcartbtn%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Tell-a-friend:</td>
		<td width="44%">
			<input class=ibtng type="file" name="checkoutbtn" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcheckoutbtn%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Wish List:</td>
		<td width="44%">
			<input class=ibtng type="file" name="addtowl" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=paddtowl%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Checkout:</td>
		<td width="44%">
			<input class=ibtng type="file" name="checkout" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcheckout%>"></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Cancel: </td>
		<td width="44%">
			<input class=ibtng type="file" name="cancel" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcancel%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Back:</td>
		<td width="44%">
			<input class=ibtng type="file" name="back" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pback%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Continue Shopping:</td>
		<td width="44%">
			<input class=ibtng type="file" name="continueshop" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcontinueshop%>"></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Small More Info: </td>
		<td width="44%">
			<input class=ibtng type="file" name="morebtn" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pmorebtn%>"></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Login:</td>
		<td width="44%">
			<input class=ibtng type="file" name="login" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=plogin%>"></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Recalculate:</td>
		<td width="44%">
			<input class=ibtng type="file" name="recalculate" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=precalculate%>" ></div>
		</td>
	</tr>
	<tr>
		<td>Register:</td>
		<td>
			<input class=ibtng type="file" name="register" size="30">
		</td>
		<td>
			<div align="center"><img src="../pc/<%=pregister%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Remove from Cart: </td>
		<td width="44%">
			<input class=ibtng type="file" name="remove" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=premove%>" ></div>
		</td>
	</tr>

	<% 'BTO ADDON-S
	If scBTO=1 then %>
	<tr>
		<td width="25%">Customize:</td>
		<td width="44%">
			<input class=ibtng type="file" name="customize" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcustomize%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Reconfigure:</td>
		<td width="44%">
			<input class=ibtng type="file" name="reconfigure" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=preconfigure%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Reset to Default:</td>
		<td width="44%">
			<input class=ibtng type="file" name="resetdefault" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=presetdefault%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Save Quote:</td>
		<td width="44%">
			<input class=ibtng type="file" name="savequote" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=psavequote%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Review &amp; Order:</td>
		<td width="44%">
			<input class=ibtng type="file" name="revorder" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=prevorder%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Submit Quote:</td>
		<td width="44%">
			<input class=ibtng type="file" name="submitquote" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=psubmitquote%>" ></div>
		</td>
	</tr>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Request a Quote:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_requestQuote" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_requestQuote%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<% End If
	'BTO ADDON-E %>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Place Order:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_placeOrder" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_placeOrder%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<!--<tr>
		<td width="25%">Checkout Without Registering:</td>
		<td width="44%">-->
			<input class=ibtng type="hidden" name="pcv_checkoutWR" size="30">
	<!--</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_checkoutWR%>"></div>
		</td>
	</tr>-->
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Process Shipment:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_processShip" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_processShip%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Finalize Shipment:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_finalShip" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_finalShip%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Back to Order Details:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_backtoOrder" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_backtoOrder%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Previous:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_previous" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_previous%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'New Button for ProductCart v3%>
	<tr>
		<td width="25%">Next:</td>
		<td width="44%">
			<input class=ibtng type="file" name="pcv_next" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_next%>"></div>
		</td>
	</tr>
	<%'End of New Button for ProductCart v3%>
	<%'GGG Add-on start%>
	<tr>
		<td width="25%">Create New Registry:</td>
		<td width="44%">
			<input class=ibtng type="file" name="crereg" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcrereg%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Delete Registry:</td>
		<td width="44%">
			<input class=ibtng type="file" name="delreg" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pdelreg%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Add To Registry:</td>
		<td width="44%">
			<input class=ibtng type="file" name="addreg" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=paddreg%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Update Registry:</td>
		<td width="44%">
			<input class=ibtng type="file" name="updreg" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pupdreg%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Send Messages:</td>
		<td width="44%">
			<input class=ibtng type="file" name="sendmsgs" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=psendmsgs%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%">Return to Registry:</td>
		<td width="44%">
			<input class=ibtng type="file" name="retreg" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pretreg%>" ></div>
		</td>
	</tr>
	<%'GGG Add-on end%>
	<tr>
		<td width="25%" nowrap>Continue/Next Step/Update:<br><span class="pcSmallText">Used on One Page Checkout</span></td>
		<td width="44%">
			<input class=ibtng type="file" name="yellowupd" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=yellowupd%>" ></div>
		</td>
	</tr>
	<tr>
		<td width="25%" nowrap>Save Cart:<br><span class="pcSmallText">Used on View Cart</span></td>
		<td width="44%">
			<input class=ibtng type="file" name="savecart" size="30">
		</td>
		<td width="31%">
			<div align="center"><img src="../pc/<%=pcv_strSaveCart%>" ></div>
		</td>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td colspan="3" align="center">
			<input name="submit" type="submit" class="submit2" value="Update">
			&nbsp;
			<input name="default" type="button" onClick="document.location.href='setBtnDefault.asp'" value="Set back to default settings">
			&nbsp;
			<input type="button" name="Button" value="Back" onClick="javascript:history.back()">
		</td>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->