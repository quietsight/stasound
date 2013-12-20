<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%><!--#include file="adminv.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->
<!--#include file="../includes/ppdstatus.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/languagesCP.asp" -->
<!--#include file="../includes/SQLFormat.txt"-->
<!--#include file="inc_UpdateDates.asp" -->

<%
on error resume next
Server.ScriptTimeout = 5400

dim f, query, conntemp, rstemp, rstemp1,TopRecord(100), IDcustom(2), Customcontent(2)
Dim PrdWithoutOpts, CheckCount
Dim ErrorsReport

PrdWithoutOpts=0
CheckCount=0

ErrorsReport=""
TempProducts=""
TempProducts=session("TempProducts") 
ErrorsReport=session("ErrorsReport")

'Import Field Names
skuN="SKU"
ig1N="Image1_General"
ig2N="Image2_General"
ig3N="Image3_General"
ig4N="Image4_General"
ig5N="Image5_General"
ig6N="Image6_General"
id1N="Image1_Detail"
id2N="Image2_Detail"
id3N="Image3_Detail"
id4N="Image4_Detail"
id5N="Image5_Detail"
id6N="Image6_Detail"

skuid=-1
ig1id=-1
ig2id=-1
ig3id=-1
ig4id=-1
ig5id=-1
ig6id=-1
id1id=-1
id2id=-1
id3id=-1
id4id=-1
id5id=-1
id6id=-1

psku=""
pig1=""
pig2=""
pig3=""
pig4=""
pig5=""
pig6=""
pid1=""
pid2=""
pid3=""
pid4=""
pid5=""
pid6=""

iPageSize=10000
iPageCurrent=session("iPageCurrent") 
if iPagecurrent="" then 
	iPageCurrent=1 
end if 

imType=session("ii_imType")
if PPD="1" then
	FileXLS = "/"&scPcFolder&"/pc/catalog/" & session("importfile")
else
	FileXLS = "../pc/catalog/" & session("importfile")
end if

Set cnnExcel = Server.CreateObject("ADODB.Connection")
cnnExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(FileXLS) & ";Extended Properties=Excel 8.0;"
Set rsExcel = Server.CreateObject("ADODB.Recordset")
	
rsExcel.CacheSize=iPageSize 
rsExcel.PageSize=iPageSize  

'/*rsExcel.open "SELECT * FROM IMPORT;", cnnExcel
'/*Altered by Sheri
rsExcel.open "SELECT * FROM IMPORT;", cnnExcel , adOpenStatic, adLockReadOnly, adCmdText

dim iPageCount 
iPageCount=rsExcel.PageCount 

If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount) 
If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1) 
'rsExcel.MoveFirst 
rsExcel.AbsolutePage=iPageCurrent 

' counting variable for our recordset 
dim count 

	
if Err.number<>0 then
	response.Clear()
	session("importfile")=""
	response.redirect "msg.asp?message=30"
end if

TotalXLSlines=session("TotalXLSlines")
if TotalXLSlines="" then
	TotalXLSlines=0
end if
ImportedRecords=session("ImportedRecords")
if ImportedRecords="" then
	ImportedRecords=0
end if
iCols = rsExcel.Fields.Count
FCSku=0
FCount=0
for i=0 to iCols-1
	if trim(rsExcel.Fields.Item(I).Name)<>"" then
		Select Case Ucase(trim(rsExcel.Fields.Item(I).Name))
			Case Ucase(skuN):
				skuID=i
				FCSku=FCSku+1
			Case Ucase(ig1N):
				ig1ID=i
				FCount=FCount+1
			Case Ucase(ig2N):
				ig2ID=i
				FCount=FCount+1
			Case Ucase(ig3N):
				ig3ID=i
				FCount=FCount+1
			Case Ucase(ig4N):
				ig4ID=i
				FCount=FCount+1
			Case Ucase(ig5N):
				ig5ID=i
				FCount=FCount+1
			Case Ucase(ig6N):
				ig6ID=i
				FCount=FCount+1
			Case Ucase(id1N):
				id1ID=i
				FCount=FCount+1
			Case Ucase(id2N):
				id2ID=i
				FCount=FCount+1
			Case Ucase(id3N):
				id3ID=i
				FCount=FCount+1
			Case Ucase(id4N):
				id4ID=i
				FCount=FCount+1
			Case Ucase(id5N):
				id5ID=i
				FCount=FCount+1
			Case Ucase(id6N):
				id6ID=i
				FCount=FCount+1
		End Select
	end if
next

if (FCSku=0) OR ((FCSku=1) AND (FCount<1)) then
	response.Clear()
	session("importfile")=""
	response.redirect "msg.asp?message=29"
end if
   	
if rsExcel.EOF then
	response.Clear()
	session("importfile")=""
	response.redirect "msg.asp?message=32"
end if

call openDb()	

SKUError=0
' set count equal to zero 
count=0 

do while not rsExcel.eof and count < rsExcel.pageSize '/*Altered by Sheri
'/*Do While not rsExcel.EOF
	
	RecordError=false
	TotalXLSlines=TotalXLSlines+1
	
	'// Get record data from XLS file
	if RecordError=False then
		psku=""
		pig1=""
		pig2=""
		pig3=""
		pig4=""
		pig5=""
		pig6=""
		pid1=""
		pid2=""
		pid3=""
		pid4=""
		pid5=""
		pid6=""

		psku=trim(rsExcel.Fields.Item(int(skuid)).Value)
		
		if Left(psku,1)="'" then
			psku=mid(psku,2,len(psku))
			if Left(psku,1)="'" then
				psku=mid(psku,2,len(psku))
			end if
		end if
		
		if psku<>"" then
			psku=replace(psku,"'","''")
			psku=replace(psku,"""","&quot;")
		end if
		
		if ig1ID<>-1 then
			pig1=trim(rsExcel.Fields.Item(int(ig1ID)).Value)
		end if
		if ig2ID<>-1 then
			pig2=trim(rsExcel.Fields.Item(int(ig2ID)).Value)
		end if
		if ig3ID<>-1 then
			pig3=trim(rsExcel.Fields.Item(int(ig3ID)).Value)
		end if
		if ig4ID<>-1 then
			pig4=trim(rsExcel.Fields.Item(int(ig4ID)).Value)
		end if
		if ig5ID<>-1 then
			pig5=trim(rsExcel.Fields.Item(int(ig5ID)).Value)
		end if
		if ig6ID<>-1 then
			pig6=trim(rsExcel.Fields.Item(int(ig6ID)).Value)
		end if
		
		if id1ID<>-1 then
			pid1=trim(rsExcel.Fields.Item(int(id1ID)).Value)
		end if
		if id2ID<>-1 then
			pid2=trim(rsExcel.Fields.Item(int(id2ID)).Value)
		end if
		if id3ID<>-1 then
			pid3=trim(rsExcel.Fields.Item(int(id3ID)).Value)
		end if
		if id4ID<>-1 then
			pid4=trim(rsExcel.Fields.Item(int(id4ID)).Value)
		end if
		if id5ID<>-1 then
			pid5=trim(rsExcel.Fields.Item(int(id5ID)).Value)
		end if
		if id6ID<>-1 then
			pid6=trim(rsExcel.Fields.Item(int(id6ID)).Value)
		end if
		
		
		'//Check data
		if psku<>"" then
		else
			ErrorsReport=ErrorsReport & "<tr align=""left""><td>" & "Record " & TotalXLSlines & ": does not include a Product SKU." & "</td></tr>" & vbcrlf
			RecordError=true
		end if
		if pig1 & pig2 & pig3 & pig4 & pig5 & pig6 & pid1 & pid2 & pid3 & pid4 & pid5 & pid6<>"" then
		else
			ErrorsReport=ErrorsReport & "<tr align=""left""><td>" & "Record " & TotalXLSlines & ": does not have a image file name to import." & "</td></tr>" & vbcrlf
			RecordError=true
		end if
	end if
		
	'// Import data to PC database
	if RecordError=false then
	
		query="SELECT idproduct FROM Products WHERE sku like '" & psku & "';"
		set rs=connTemp.execute(query)
		IF not rs.eof THEN
			pPrdID=rs("idProduct")
			set rs=nothing
						
			if imType="1" then
				query="DELETE FROM pcProductsImages WHERE idProduct=" & pPrdID & ";"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig1<>"" OR pid1<>"" then
				if pid1<>"" AND pig1="" then
					pig1=pid1
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig1 & "','" & pid1 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig2<>"" OR pid2<>"" then
				if pid2<>"" AND pig2="" then
					pig2=pid2
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig2 & "','" & pid2 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig3<>"" OR pid3<>"" then
				if pid3<>"" AND pig3="" then
					pig3=pid3
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig3 & "','" & pid3 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig4<>"" OR pid4<>"" then
				if pid4<>"" AND pig4="" then
					pig4=pid4
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig4 & "','" & pid4 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig5<>"" OR pid5<>"" then
				if pid5<>"" AND pig5="" then
					pig5=pid5
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig5 & "','" & pid5 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if
			
			if pig6<>"" OR pid6<>"" then
				if pid6<>"" AND pig6="" then
					pig6=pid6
				end if
				query="INSERT INTO pcProductsImages (idProduct,pcProdImage_Url,pcProdImage_LargeUrl) VALUES (" & pPrdID & ",'" & pig6 & "','" & pid6 & "');"
				set rs=connTemp.execute(query)
				set rs=nothing
			end if

		ELSE
		'Does not have SKU
			set rs=nothing
			SKUError=1
			ErrorsReport=ErrorsReport & "<tr align=""left""><td>" & "Record " & TotalXLSlines & ": This SKU is not in the database." & "</td></tr>" & vbcrlf
			RecordError=true
		END IF
	end if
			
	if RecordError=false then
		ImportedRecords=ImportedRecords+1
	end if

	count=count + 1 
	rsExcel.MoveNext
	
Loop
rsexcel.Close	
set rsexcel=nothing
cnnExcel.close
set cnnExcel=nothing

call closeDB()
	
session("TempProducts")=TempProducts 
session("ErrorsReport")=ErrorsReport 
session("iPageCurrent")=iPageCurrent+1 
session("TotalXLSlines")=TotalXLSlines 
session("ImportedRecords")=ImportedRecords 

If Cint(iPageCurrent) < Cint(iPageCount) Then
	response.redirect "iistep3.asp"
else
	session("importfile")=""
end if
	
if SKUError=1 then
	ErrorsReport="<tr align=""left""><td>One of the records you are importing does not currently exist in the database. The Import feature is strictly for modifying existing product information. Please correct the error and try again." & "</td></tr>" & vbcrlf&vbcrlf &ErrorsReport
	session("ErrorsReport")=ErrorsReport
end if

pageTitle = pageTitle & " PRODUCT ADDITIONAL IMAGES IMPORT WIZARD - Review Import Results"
section = "products" %>
<!--#include file="AdminHeader.asp"-->
<script type="text/javascript" language="javascript" src="../includes/spry/SpryDOMUtils.js"></script>
<style type="text/css">
<!--
.grayBG {
	background-color: #F5F5F5;
}
-->
</style>
<table class="pcCPcontent">
<tr>
	<td valign="top">
        
	<%if ImportedRecords>0 then%>
		<div class="pcCPmessageSuccess">A total of <b><%=ImportedRecords%></b> records were imported successfully!</div>
	<%end if%>
		
	<%if TotalXLSlines-ImportedRecords>0 then%>
		<div class="pcCPmessage">A total of <b><font color="#FF0000"><%=TotalXLSlines-ImportedRecords%></font></b> records <u>could NOT</u> be imported. See the Error(s) Report section below for details</div>
	<%end if%>

	<% if ErrorsReport<>"" then%> 
	<br /><br />
	<table class="pcCPcontent">
	<tr> 
		<td> 
			<table border="0" cellspacing="0" width="100%" cellpadding="2">
				<tr>
					<th>
						Error(s) Report
					</th>
				</tr>
                <tr>
                    <td align="center">
                        <div style="width: 98%; height: 150px; overflow: scroll; border: 1px dotted #E1E1E1;">
                            <table id="noheaderodd" style="font-family: Arial; font-size: 9px; width: 100%;">
                                <%=ErrorsReport%>
                            </table>
                        </div>
                    </td>
                </tr>
			</table> 
			<script type="text/javascript" language="javascript">
				Spry.$$("table#noheaderodd tr:nth-child(odd)").addClassName("grayBG");            
            </script> 
		</td>
	</tr>
	</table>
	<%end if%>
  	<br /><br />
	<p align="center">
	<input type=button name=mainmenu value="Back to Main menu" onClick="location='menu.asp';" class="ibtnGrey">
	</p>
	</td>
</tr>
</table>
<% If session("importfile")="" Then
	session("ii_imType")=""
	session("ErrorsReport")=""
	session("iPageCurrent")=""
	session("TotalXLSlines")=""
	session("ImportedRecords")=""
end if %>
<!--#include file="AdminFooter.asp"-->