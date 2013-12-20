<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=7%>
<% 
pageTitle="Manage Gift Registry"
pageIcon="pcv4_icon_gift.png"
section="mngAcc" 
%>
<!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->  
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/currencyformatinc.asp"--> 
<!--#include file="../includes/languages.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="../includes/ErrorHandler.asp"-->
<%

pcv_IdCustomer=getUserInput(request("idCustomer"),0)
gIDEvent=getUserInput(request("IDEvent"),0)

if pcv_IdCustomer="" or gIDEvent="" then
	response.redirect "menu.asp"
end if

dim conntemp, rs, rstemp
call openDb()

if validNum(gIDEvent) then

	query="select pcEv_IDEvent,pcEv_Name,pcEv_Date,pcEv_Type,pcEv_IncGcs from pcEvents where pcEv_IDCustomer=" & pcv_IdCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		call closeDb()
		response.redirect "ggg_manageGRs.asp?idcustomer=" & pcv_IdCustomer
	else
		geName=rstemp("pcEv_Name")
		geDate=rstemp("pcEv_Date")
		if gedate<>"" then
			if scDateFrmt="DD/MM/YY" then
				gedate=(day(gedate)&"/"&month(gedate)&"/"&year(gedate))
			else
				gedate=(month(gedate)&"/"&day(gedate)&"/"&year(gedate))
			end if
		end if
		gType=rstemp("pcEv_Type")
		if gType<>"" then
		else
			gType=""
		end if
		geincGc=rstemp("pcEv_IncGcs")
		if geincGc<>"" then
		else
			geincGc="0"
		end if
		if request("addgc")="1" then
			if geincgc="1" then
				query="select IDProduct from Products where pcprod_GC=1 and removed=0 and active<>0"
				set rstemp=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				do while not rstemp.eof
					IDProduct=rstemp("IDProduct")
					query="select pcEP_IDProduct from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_IDProduct=" & IDProduct & " and pcEP_GC=1"
					set rs1=conntemp.execute(query)
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs1=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					if rs1.eof then
						query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_GC) values (" & gIDEvent & "," & IDProduct & ",1)"
						set rs1=conntemp.execute(query)
						if err.number<>0 then
							call LogErrorToDatabase()
							set rs1=nothing
							call closedb()
							response.redirect "techErr.asp?err="&pcStrCustRefID
						end if
					end if
					set rs1=nothing
					rstemp.MoveNext
				loop
				set rs=nothing
			end if
		end if
	end if
	set rstemp=nothing
end if

if (request("action")="update") and (request("submit1")<>"") then
	Count=getUserInput(request("Count"),0)
	
	For k=1 to Count
		geID=getUserInput(request("geID" & k),0)
		gQty=getUserInput(request("gQty" & k),0)
		gHQty=getUserInput(request("gHQty" & k),0)
		if gQty="" then
			gQty="0"
		end if
		if gHQty="" then
			gHQty="0"
		end if
		gesel=getUserInput(request("sel" & k),0)
		if gesel="" then
			gesel="0"
		end if
		IF (geID<>"") and (gesel="1") THEN
		'Update
			query="update pcEvProducts set pcEP_Qty=" & gQty & ",pcEP_HQty=" & gHQty & " where pcEP_IDEvent=" & gIDEvent & " and pcEP_ID=" & geID
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rstemp=nothing
		END IF
	Next
	msg="Products were updated successfully!"
	msgType=1
end if

if (request("action")="update") and (request("submit2")<>"") then
	Count=getUserInput(request("Count"),0)
	
	For k=1 to Count
		geID=getUserInput(request("geID" & k),0)
		gesel=getUserInput(request("sel" & k),0)
		if gesel="" then
			gesel="0"
		end if
		IF (geID<>"") and (gesel="1") THEN
			query="delete from pcEvProducts where pcEP_IDEvent=" & gIDEvent & " and pcEP_ID=" & geID
			set rstemp=connTemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rstemp=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			set rstemp=nothing
		END IF
	Next
	msg="Products were deleted successfully!"
	msgType=1
end if

%>
<!--#include file="Adminheader.asp"--> 

<%
query="SELECT products.idProduct,products.sku,products.description,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC,pcEvProducts.pcEP_OptionsArray FROM products,pcEvProducts WHERE pcEvProducts.pcEP_IDEvent=" & gIDEvent & " AND products.idproduct=pcEvProducts.pcEP_IDProduct AND products.removed=0 AND products.active=-1 ORDER BY products.Description ASC,pcEvProducts.pcEP_GC ASC"
set rstemp=Server.CreateObject("ADODB.Recordset")
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
%>
<script language="JavaScript">
<!--
	function newWindow2(file,window)
	{
		catWindow=open(file,window,'resizable=no,width=480,height=360,scrollbars=1');
		if (catWindow.opener == null) catWindow.opener = self;
		checkwin();
	}
	
	function checkwin()
	{
		if (catWindow.closed)
		{
			location="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer%>";
		}
		else
		{
			setTimeout('checkwin()',500);
		}
	}

//-->
</script>

<form method="post" name="Form1" action="ggg_GRDetails.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
<table class="pcCPcontent">
    <tr>
        <td colspan="6" class="pcCPspacer">
            <% ' START show message, if any %>
                <!--#include file="pcv4_showMessage.asp"-->
            <% 	' END show message %>
        </td>
    </tr>
    <tr> 
        <td colspan="6">
            Event Name: <b><%=geName%></b>
            <br><%response.write dictLanguage.Item(Session("language")&"_GRDetails_1b")%><%=geDate%>
            <%if gType<>"" then%>
                <br>Event Type: <%=gType%>
            <%end if%>
            <br>
            <br>
        </td>
    </tr>
	
<%IF rstemp.eof then%>
<tr>
	<td colspan="6">
		<br>
		<div class="pcCPmessage">The customer does not currently have any products in this registry.</div>
		<br>
		<input type=button name="back1" value=" Back " class="ibtnGrey" onclick="javascript:location='ggg_manageGRs.asp?idcustomer=<%=pcv_IdCustomer%>';">
		<br>
	</td>
</tr>
<%ELSE%>
<tr>
	<th nowrap>SKU</th>
	<th nowrap width="90%">Product Name</th>
	<th nowrap align="right">Items</th>
	<th nowrap align="right">Has</th>
	<th nowrap>&nbsp;</th>
</tr>
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<%
Count=0
do while not rstemp.eof
	pidProduct=rstemp("idProduct")
	gsku=rstemp("sku")
	gname=rstemp("description")
	geID=rstemp("pcEP_ID")
	gQty=rstemp("pcEP_Qty")
	if gQty<>"" then
	else
		gQty="0"
	end if
	gHQty=rstemp("pcEP_HQty")
	if gHQty<>"" then
	else
		gHQty="0"
	end if
	gGC=rstemp("pcEP_GC")
	pcv_Opts=rstemp("pcEP_OptionsArray")
	if gGC<>"1" then
		Count=Count+1%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td nowrap valign="top">
				<%=gsku%>
			</td>
			<td nowrap valign="top">
				<a href="FindProductType.asp?id=<%=pidProduct%>" target="_blank"><%=gname%></a>
				<%
				'// Display Options Link, If there are options
				
				'// CHECK FOR OPTIONS
				' SELECT DATA SET
				' TABLES: products, pcProductsOptions, optionsgroups, ptions_optionsGroups
				query = "SELECT DISTINCT optionsGroups.OptionGroupDesc, pcProductsOptions.pcProdOpt_ID, pcProductsOptions.idOptionGroup, pcProductsOptions.pcProdOpt_Required, pcProductsOptions.pcProdOpt_Order "
				query = query & "FROM products "
				query = query & "INNER JOIN ( "
				query = query & "pcProductsOptions INNER JOIN ( "
				query = query & "optionsgroups "
				query = query & "INNER JOIN options_optionsGroups "
				query = query & "ON optionsgroups.idOptionGroup = options_optionsGroups.idOptionGroup "
				query = query & ") ON optionsGroups.idOptionGroup = pcProductsOptions.idOptionGroup "
				query = query & ") ON products.idProduct = pcProductsOptions.idProduct "
				query = query & "WHERE products.idProduct=" & pidProduct &" "
				query = query & "AND options_optionsGroups.idProduct=" & pidProduct &" "
				query = query & "ORDER BY pcProductsOptions.pcProdOpt_Order;"
				set rsCheckOptions=server.createobject("adodb.recordset")
				set rsCheckOptions=conntemp.execute(query)	
			
				If NOT rsCheckOptions.eof Then
					pcv_intOptionsExist = 1
				Else
					pcv_intOptionsExist = 2
				End If	
				set rsCheckOptions = nothing	
				If pcv_intOptionsExist = 1 Then
				%>
					(<a href="#" onClick="newWindow2('ggg_options_popup.asp?IDEvent=<%=gIDEvent%>&idOptionArray=<%=pcv_Opts%>&idproduct=<%=pidProduct%>&epID=<%=geID%>','window2')">Add/Edit Options</a>)
				<%
				End If
				%>
				<%'Show Product Options
				if pcv_Opts<>"" then
				%>
				<br>
				<%
				tmp_Arr=split(pcv_Opts,"|")
				for i=lbound(tmp_Arr) to ubound(tmp_Arr)
				if tmp_Arr(i)<>"" then
					query="SELECT optionsGroups.OptionGroupDesc FROM optionsGroups INNER JOIN options_optionsGroups ON optionsGroups.idOptionGroup=options_optionsGroups.idOptionGroup WHERE idoptoptgrp=" & tmp_Arr(i)
					set rs=connTemp.execute(query)
					pcv_OptGrpName=""
					if not rs.eof then
						pcv_OptGrpName=rs("OptionGroupDesc")
					end if
					set rs=nothing
					query="SELECT options.optionDescrip FROM options INNER JOIN options_optionsGroups ON options.idOption=options_optionsGroups.idOption WHERE idoptoptgrp=" & tmp_Arr(i)
					set rs=connTemp.execute(query)
					pcv_OptName=""
					if not rs.eof then
						pcv_OptName=rs("optionDescrip")
					end if
					set rs=nothing
					if (pcv_OptGrpName<>"") and (pcv_OptName<>"") then%>
						<%=pcv_OptGrpName%>: <%=pcv_OptName%><br>
					<%end if%>
				<%end if
				next
				end if%>
			</td>
			<td nowrap align="right" valign="top">
				<input name="gQty<%=Count%>" value="<%=gQty%>" size="3" style="float: right; text-align:right">
			</td>
			<td nowrap align="right" valign="top">
				<input name="gHQty<%=Count%>" value="<%=gHQty%>" size="3" style="float: right; text-align:right">
			</td>
			<td width="20%" align="center" valign="top">
				<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
				<input type=checkbox name="sel<%=Count%>" value="1" class="clearBorder">
			</td>
		</tr>
	<%
	end if
	rstemp.MoveNext
loop
set rstemp=nothing

query="select products.sku,products.description,pcEvProducts.pcEP_ID,pcEvProducts.pcEP_Qty,pcEvProducts.pcEP_HQty,pcEvProducts.pcEP_GC from products,pcEvProducts where pcEvProducts.pcEP_IDEvent=" & gIDEvent & " and products.idproduct=pcEvProducts.pcEP_IDProduct and pcEvProducts.pcEP_GC=1 and products.removed=0 order by products.Description asc,pcEvProducts.pcEP_GC asc"
set rstemp=connTemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rstemp=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if

IF NOT rstemp.eof then%>
<tr>
	<td colspan="5" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<th colspan="5">Gift Certificates</th>
</tr>
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<%do while not rstemp.eof
	gsku=rstemp("sku")
	gname=rstemp("description")
	geID=rstemp("pcEP_ID")
	gQty=rstemp("pcEP_Qty")
	if gQty<>"" then
	else
		gQty="0"
	end if
	gHQty=rstemp("pcEP_HQty")
	if gHQty<>"" then
	else
		gHQty="0"
	end if
	gGC=rstemp("pcEP_GC")
	if ((gGC="1") and (geincgc="1")) or (clng(gHQty)>0) then
		Count=Count+1%>
		<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
			<td nowrap>
				<%=gsku%>
			</td>
			<td nowrap>
				<%=gname%>
			</td>
			<td nowrap align="right">
			</td>
			<td nowrap align="right">
				<input name="gHQty<%=Count%>" value="<%=gHQty%>" size="3" style="float: right; text-align:right">
				<input type="hidden" name="gQty<%=Count%>" value="<%=gQty%>">
			</td>
			<td width="20%" align="center">
				<input type=hidden name="geID<%=Count%>" value="<%=geID%>">
				<input type=checkbox name="sel<%=Count%>" value="1" class="clearBorder">
			</td>
		</tr>
	<%
	end if
	rstemp.MoveNext
loop
END IF 'Have GCs
set rstemp=nothing

if geincgc="1" then
query="select IDProduct from Products where pcprod_GC=1 and removed=0 and active<>0"
set rstemp=connTemp.execute(query)
if not rstemp.eof then%>
<tr>
	<td colspan="5" class="pcCPspacer"></td>
</tr>
<tr>
	<td colspan="5">
		<a href="ggg_GRDetails.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer%>&addgc=1">Reset Gift Certificate product list</a>
	</td>
</tr>
<%end if
set rstemp=nothing
end if%>
<%if Count>0 then%>
<tr>
	<td colspan="5" class="pcSpacer">&nbsp;</td>
</tr>
<tr>
	<td colspan="5" align="right" class="cpLinksList">
		<a href="javascript:checkAll();">Check All</a>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a>
	</td>
</tr>
<script language="JavaScript">
<!--
function checkAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.Form1.sel" + j); 
if (box.checked == false) box.checked = true;
   }
}

function uncheckAll() {
for (var j = 1; j <= <%=count%>; j++) {
box = eval("document.Form1.sel" + j); 
if (box.checked == true) box.checked = false;
   }
}

//-->
</script>
<%end if%>
<tr>
	<td colspan="5" class="pcSpacer">&nbsp;</td>
</tr>
<tr> 
	<td colspan="5"> 
		<input type="submit" class="submit2" name="submit1" value=" Update Registry ">&nbsp;
		<input type="submit" class="submit2" name="submit2" value=" Remove Selected Products ">&nbsp;
		<input type="button" name="modify" value=" Modify Details " onclick="javascript:location='ggg_EditGR.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer%>';">&nbsp;
		<input type="button" name="addnew" value=" Add New Products " onclick="javascript:location='ggg_addtoGR.asp?IDEvent=<%=gIDEvent%>&idcustomer=<%=pcv_IdCustomer%>';">&nbsp;
		<input type="button" name="back" value=" Back " onclick="javascript:location='ggg_manageGRs.asp?idcustomer=<%=pcv_IdCustomer%>';">
		<input type="hidden" name="IDEvent" value="<%=gIDEvent%>">
		<input type="hidden" name="Count" value="<%=Count%>">
		<input type="hidden" name="idcustomer" value="<%=pcv_IdCustomer%>">
	</td>
</tr>
<%END IF 'Have products%>
</table>
</form>

<script language="JavaScript">
<!--
function isDigit(s)
{
	var test=""+s;
	if(test=="-"||test=="+"||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
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
	
function Form1_Validator(theForm)
{
	<%For k=1 to Count%>
		if ((theForm.gQty<%=k%>.value != "") && (theForm.gQty<%=k%>.value != "0"))
	  	{
			if (allDigit(theForm.gQty<%=k%>.value) == false)
			{
				alert("Please enter a valid number for this field.");
				theForm.gQty<%=k%>.focus();
			    return (false);
			}
		}
		if ((theForm.gHQty<%=k%>.value != "") && (theForm.gHQty<%=k%>.value != "0"))
	  	{
			if (allDigit(theForm.gHQty<%=k%>.value) == false)
			{
				alert("Please enter a valid number for this field.");
				theForm.gHQty<%=k%>.focus();
			    return (false);
			}
		}
	<%Next%>

	return (true);
}
//-->
</script>
<%call closedb()%>
<!--#include file="Adminfooter.asp"-->