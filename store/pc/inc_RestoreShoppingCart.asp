<!--#include file="pcCheckPricingCats.asp"-->
<%
pcCartIndex=Session("pcCartIndex")
if IsNull(pcCartIndex) then
	pcCartArray=Session("pcCartSession")
end if
ppcCartIndex=Session("pcCartIndex")

HasSavedCart=0
HasSavedPrds=0
IDSC=0
tmpGUID=getUserInput(Request.Cookies("SavedCartGUID"),0)

if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
	tmpIDCust=session("IDCustomer")
else
	tmpIDCust=0
end if
call opendb()
query="SELECT IDCustomer FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	if clng(tmpIDCust)<>clng(rsQ("IDCustomer")) then
		tmpGUID=""
	end if
end if
set rsQ=nothing
call closedb()

'// Check if this feature is disabled
if (scRestoreCart=0 or isNull(scRestoreCart) or scRestoreCart="") AND (HaveToRestore<>"yes") then
 tmpGUID=""
end if

'// v4.5: Disable feature if Admin Order
'if session("pcAdminOrder")=1 then 
' tmpGUID=""
'end if

IF tmpGUID<>"" THEN	
	connTemp=""
	call opendb()
	if session("IDCustomer")<>"" AND session("IDCustomer")<>"0" then
		tmpIDCust=session("IDCustomer")
	else
		tmpIDCust=0
	end if
	query="SELECT SavedCartID FROM pcSavedCarts WHERE SavedCartGUID like '" &  tmpGUID & "';"
	set rsQ=connTemp.execute(query)
	if not rsQ.eof then
		IDSC=rsQ("SavedCartID")
		query="SELECT SCArray0, SCArray1, SCArray2, SCArray3, SCArray4, SCArray5, SCArray6, SCArray7, SCArray8, SCArray9, SCArray10, SCArray11, SCArray12, SCArray13, SCArray14, SCArray15, SCArray16, SCArray17, SCArray18, SCArray19, SCArray20, SCArray21, SCArray22, SCArray23, SCArray24, SCArray25, SCArray26,SCArray27, SCArray28, SCArray29,SCArray30, SCArray31, SCArray32,SCArray33, SCArray34, SCArray35, SCArray36, SCArray37, SCArray38, SCArray39, SCArray40, SCArray41, SCArray42, SCArray43, SCArray44, SCArray45 FROM pcSavedCartArray WHERE SavedCartID=" & IDSC & ";"
		set rsQ=connTemp.execute(query)
		if not rsQ.eof then
			tmpCartArr=rsQ.getRows()
			set rsQ=nothing
			tmpIntCount=ubound(tmpCartArr,2)
			HasSavedCart=1
		else
			query="DELETE FROM pcSavedCarts WHERE SavedCartID=" & IDSC & ";"
			set rsQ=connTemp.execute(query)
			set rsQ=nothing
			Response.Cookies("SavedCartGUID")=""
		end if
		set rsQ=nothing
	else
		Response.Cookies("SavedCartGUID")=""
	end if
	set rsQ=nothing
	
	IF HasSavedCart=1 THEN
		For k=0 to tmpIntCount
			tmpPrdID=tmpCartArr(0,k)
			pquantity=tmpCartArr(2,k)
			
			pcv_OrdHaveOutStock=0
			
			'// START v4.1 - Not For Sale override
				if NotForSaleOverride(session("customerCategory"))=1 then
					queryNFSO=""
				else
					queryNFSO=" AND products.formQuantity=0"
				end if
			'// END v4.1

			query="SELECT products.serviceSpec,products.stock,products.noStock,products.pcprod_minimumqty,products.pcprod_qtyvalidate,products.pcProd_BackOrder,products.Description FROM Products WHERE products.idproduct=" & tmpPrdID & " AND products.removed=0 AND products.active<>0" & queryNFSO & ";"
			set rsQ=connTemp.execute(query)
			IF not rsQ.eof THEN	
				
				pserviceSpec=rsQ("serviceSpec")
				pStock=rsQ("stock")
				pNoStock=rsQ("noStock")
				pcv_minqty=rsQ("pcprod_minimumqty")
				pcv_qtyvalid=rsQ("pcprod_qtyvalidate")
				pcv_BackOrder=rsQ("pcProd_BackOrder")
				tmpCartArr(1,k)=rsQ("Description")
				set rsQ=nothing
				
				if PStock<pcv_minqty then
					pStock=0
				else
					if (PStock<pquantity) and (pStock>pcv_minqty) then
						pcv_minqty1=pcv_minqty
						if pcv_minqty1=0 then
							pcv_minqty1=1
						end if
						pStock=Fix(pStock/pcv_minqty1)*pcv_minqty1
					end if
				end if
				
				IF (scOutofStockPurchase=-1 AND CLng(pStock)<1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND CLng(pStock)<1 AND pNoStock=0 AND pcv_BackOrder=0) THEN
					pcv_OrdHaveOutStock=1
				END IF
				
				if pcv_OrdHaveOutStock=0 then
					if pStock=0 then
						if pcv_minqty>"0" then
							PStock=pcv_minqty
						else
							pStock=1
						end if
					end if
			
					IF (scOutofStockPurchase=-1 AND pserviceSpec=0 AND pNoStock=0 AND pcv_BackOrder=0) OR (pserviceSpec<>0 AND scOutofStockPurchase=-1 AND iBTOOutofstockpurchase=-1 AND pNoStock=0 AND pcv_BackOrder=0) THEN	
						if pStock<pquantity then
							pquantity=pStock
						end if
					END IF
				end if
			ELSE
				pcv_OrdHaveOutStock=1
			END IF
			
			set rsQ=nothing
			
			IF pcv_OrdHaveOutStock=0 THEN
				HasSavedPrds=1
				pcCartIndex=pcCartIndex+1
				for x=0 to 45
					If x=38 AND tmpCartArr(x,k)=""  Then
						pcCartArray(pcCartIndex,x)=0
					Else
					pcCartArray(pcCartIndex,x)=tmpCartArr(x,k)
					End If
				next
			END IF
		Next
		IF HasSavedPrds=1 THEN
			pcCartArray(1,18)=0
			Session("pcCartSession")=pcCartArray
			Session("pcCartIndex")=pcCartIndex
			ppcCartIndex=pcCartIndex
			session("NeedToShowRSCMsg")="1"
			NeedReCalculate=1
			%>
			<!--#include file="pcReCalPricesLogin.asp"-->
			<%
		END IF
	ELSE
		Response.Cookies("SavedCartGUID")=""
	END IF
	call closedb()
END IF
%>