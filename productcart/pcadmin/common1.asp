<%
		InvalidGrp1=0
		InvalidGrp2=0
		InvalidGrp3=0
		InvalidGrp4=0
		InvalidGrp5=0
		
		surcharge1=0
		surcharge2=0
		
		prdnote=""
		eimag=0
		playout=""
		hidesku=0

		'//Google Shopping		
		goCat=""
		goGen=""
		goAge=""
		goColor=""
		goSize=""
		goPat=""
		goMat=""

		psku=""
		pname=""
		pdesc=""
		sdesc=""
		pptype=""
		poprice=""
		plprice=""
		pwprice=""
		pweight=""
		unitslb=0
		pstock=""
		
		mt_title=""
		mt_desc=""
		mt_key=""
		
		pcategory=""
		SCATDesc=""
		LCATDesc=""
		pcsimage=""
		pclimage=""
		ppcategory=""
		
		pcategory1=""
		SCATDesc1=""
		LCATDesc1=""
		pcsimage1=""
		pclimage1=""
		ppcategory1=""
		
		pcategory2=""
		SCATDesc2=""
		LCATDesc2=""
		pcsimage2=""
		pclimage2=""
		ppcategory2=""
		
		ptimage=""
		pgimage=""
		pdimage=""
		pactive=""
		psaving=""
		pspecial=""
		pfeatured=0
		prwp=""
		pntax=""
		pnship=""
		pnforsale=""
		pnforsalecopy=""
		distock=0
		dishiptext=0
		OverSize=""
		brandname=""
		brandlogo=""
		
		fileurl=""
		urlexpire=0
		expiredays=0
		license=0
		localLG=""
		remoteLG=""
		LFN1=""
		LFN2=""
		LFN3=""
		LFN4=""
		LFN5=""
		Addcopy=""
		MQty=0
		VQty=0
		
		'BTO
		prd_hidebtoprice=0
		prd_hideconf=0
		prd_dispurchase=0
		prd_skipdetails=0
		
		prd_Opt1=""
		prd_Attr1=""
		prd_Opt1Req=0
		prd_Opt1Order=1
		prd_Opt2=""
		prd_Attr2=""
		prd_Opt2Req=0
		prd_Opt2Order=2
		prd_Opt3=""
		prd_Attr3=""
		prd_Opt3Req=0
		prd_Opt3Order=3
		prd_Opt4=""
		prd_Attr4=""
		prd_Opt4Req=0
		prd_Opt4Order=4
		prd_Opt5=""
		prd_Attr5=""
		prd_Opt5Req=0
		prd_Opt5Order=5

		prd_downprd=0
		prd_giftcert=0
		prd_giftexp=0
		prd_giftelect=0
		prd_giftgen=0
		prd_giftexpdate=""
		prd_giftexpdays=0
		prd_giftcustgenfile=""
		
		'Start SDBA
		prd_Cost=0
		prd_BackOrder=0
		prd_ShipNDays=0
		prd_NotifyStock=0
		prd_ReorderLevel=0
		prd_IsDropShipped=0
		prd_Supplier=0
		prd_DropShipper=0
		prd_IsDropShipper=0
		'End SDBA

		psku=trim(CSVRecord(skuid))
		
		if Left(psku,1)="'" then
			psku=mid(psku,2,len(psku))
			if Left(psku,1)="'" then
				psku=mid(psku,2,len(psku))
			end if
		end if
		
		if surcharge1ID<>-1 then
			surcharge1=trim(CSVRecord(surcharge1ID))
			if (not (surcharge1<>"")) or (isNumeric(surcharge1)=false) then
				surcharge1="0"
			end if
		end if
		
		if surcharge2ID<>-1 then
			surcharge2=trim(CSVRecord(surcharge2ID))
			if (not (surcharge2<>"")) or (isNumeric(surcharge2)=false) then
				surcharge2="0"
			end if
		end if
		
		if goCatid<>-1 then
			goCat=trim(CSVRecord(goCatid))
		end if
		
		if goGenid<>-1 then
			goGen=trim(CSVRecord(goGenid))
		end if
		
		if goAgeid<>-1 then
			goAge=trim(CSVRecord(goAgeid))
		end if
		
		if goColorid<>-1 then
			goColor=trim(CSVRecord(goColorid))
		end if
		
		if goSizeid<>-1 then
			goSize=trim(CSVRecord(goSizeid))
		end if
		
		if goPatid<>-1 then
			goPat=trim(CSVRecord(goPatid))
		end if
		
		if goMatid<>-1 then
			goMat=trim(CSVRecord(goMatid))
		end if
		
		if prdnoteID<>-1 then
			prdnote=trim(CSVRecord(prdnoteID))
		end if
		
		if playoutID<>-1 then
			playout=trim(CSVRecord(playoutID))
			if playout<>"" then
				Select Case playout
					Case "l","c","o":
					Case Else: playout=""
				End Select
			end if
		end if
		
		if eimagID<>-1 then
			eimag=trim(CSVRecord(eimagID))
			if (not (eimag<>"")) or (isNumeric(eimag)=false) then
				eimag="0"
			end if
		end if
		
		if hideskuID<>-1 then
			hidesku=trim(CSVRecord(hideskuID))
			if (not (hidesku<>"")) or (isNumeric(hidesku)=false) then
				hidesku="0"
			end if
		end if

		'Start SDBA
		if prd_CostID<>-1 then
			prd_Cost=trim(CSVRecord(prd_CostID))
			if (not (prd_Cost<>"")) or (isNumeric(prd_Cost)=false) then
				prd_Cost="0"
			end if
		end if
		
		if prd_BackOrderID<>-1 then
			prd_BackOrder=trim(CSVRecord(prd_BackOrderID))
			if (not (prd_BackOrder<>"")) or (isNumeric(prd_BackOrder)=false) then
				prd_BackOrder="0"
			end if
		end if
		
		if prd_ShipNDaysID<>-1 then
			prd_ShipNDays=trim(CSVRecord(prd_ShipNDaysID))
			if (not (prd_ShipNDays<>"")) or (isNumeric(prd_ShipNDays)=false) then
				prd_ShipNDays="0"
			end if
		end if
		
		if prd_NotifyStockID<>-1 then
			prd_NotifyStock=trim(CSVRecord(prd_NotifyStockID))
			if (not (prd_NotifyStock<>"")) or (isNumeric(prd_NotifyStock)=false) then
				prd_NotifyStock="0"
			end if
		end if
		
		if prd_ReorderLevelID<>-1 then
			prd_ReorderLevel=trim(CSVRecord(prd_ReorderLevelID))
			if (not (prd_ReorderLevel<>"")) or (isNumeric(prd_ReorderLevel)=false) then
				prd_ReorderLevel="0"
			end if
		end if
		
		if prd_IsDropShippedID<>-1 then
			prd_IsDropShipped=trim(CSVRecord(prd_IsDropShippedID))
			if (not (prd_IsDropShipped<>"")) or (isNumeric(prd_IsDropShipped)=false) then
				prd_IsDropShipped="0"
			end if
		end if
		
		if prd_SupplierID<>-1 then
			prd_Supplier=trim(CSVRecord(prd_SupplierID))
			if (not (prd_Supplier<>"")) or (isNumeric(prd_Supplier)=false) then
				prd_Supplier="0"
			end if
		end if
		
		if prd_DropShipperID<>-1 then
			prd_DropShipper=trim(CSVRecord(prd_DropShipperID))
			if (not (prd_DropShipper<>"")) or (isNumeric(prd_DropShipper)=false) then
				prd_DropShipper="0"
			end if
		end if
		
		if prd_IsDropShipperID<>-1 then
			prd_IsDropShipper=trim(CSVRecord(prd_IsDropShipperID))
			if (not (prd_IsDropShipper<>"")) or (isNumeric(prd_IsDropShipper)=false) then
				prd_IsDropShipper="0"
			end if
		end if
		'End SDBA
		
		if mt_titleID<>-1 then
			mt_title=trim(CSVRecord(mt_titleID))
		end if
		
		if mt_descID<>-1 then
			mt_desc=trim(CSVRecord(mt_descID))
		end if
		
		if mt_keyID<>-1 then
			mt_key=trim(CSVRecord(mt_keyID))
		end if
		
		if Opt1ID<>-1 then
		prd_Opt1=trim(CSVRecord(Opt1id))
		end if
		
		if Opt2ID<>-1 then
		prd_Opt2=trim(CSVRecord(Opt2id))
		end if
		
		if Attr1ID<>-1 then
		prd_Attr1=trim(CSVRecord(Attr1id))
		end if
		
		if prd_Attr1<>"" then
		else
		InvalidGrp1=1
		end if
		
		if Attr2ID<>-1 then
		prd_Attr2=trim(CSVRecord(Attr2id))
		end if
		
		if prd_Attr2<>"" then
		else
		InvalidGrp2=1
		end if
		
		if Opt1ReqID<>-1 then
		prd_Opt1Req=trim(CSVRecord(Opt1Reqid))
		if prd_Opt1Req<>"" then
		else
		InvalidGrp1=1
		prd_Opt1Req="0"
		end if
		if IsNumeric(prd_Opt1Req)=false then
		InvalidGrp1=1
		prd_Opt1Req="0"
		end if
		if prd_Opt1Req>"1" then
		InvalidGrp1=1
		prd_Opt1Req="1"
		end if
		end if
		
		if Opt2ReqID<>-1 then
		prd_Opt2Req=trim(CSVRecord(Opt2Reqid))
		if prd_Opt2Req<>"" then
		else
		InvalidGrp2=1
		prd_Opt2Req="0"
		end if
		if IsNumeric(prd_Opt2Req)=false then
		InvalidGrp2=1
		prd_Opt2Req="0"
		end if
		if prd_Opt2Req>"1" then
		InvalidGrp2=1
		prd_Opt2Req="1"
		end if
		end if
		
		if Opt1OrderID<>-1 then
			prd_Opt1Order=trim(CSVRecord(Opt1Orderid))
			if prd_Opt1Order<>"" then
			else
				InvalidGrp1=1
				prd_Opt1Order="1"
			end if
			if IsNumeric(prd_Opt1Order)=false then
				InvalidGrp1=1
				prd_Opt1Order="1"
			end if
			if prd_Opt1Order>"1" then
				prd_Opt1Order="1"
			end if
		end if
		
		if Opt2OrderID<>-1 then
			prd_Opt2Order=trim(CSVRecord(Opt2Orderid))
			if prd_Opt2Order<>"" then
			else
				InvalidGrp2=1
				prd_Opt2Order="2"
			end if
			if IsNumeric(prd_Opt2Order)=false then
				InvalidGrp2=1
				prd_Opt2Order="2"
			end if
			if prd_Opt2Order>"2" then
				prd_Opt2Order="2"
			end if
		end if
		
		if Opt3ID<>-1 then
			prd_Opt3=trim(CSVRecord(Opt3id))
		end if
		
		if Attr3ID<>-1 then
			prd_Attr3=trim(CSVRecord(Attr3id))
		end if
		
		if prd_Attr3<>"" then
		else
			InvalidGrp3=1
		end if

		if Opt3ReqID<>-1 then
			prd_Opt3Req=trim(CSVRecord(Opt3Reqid))
			if prd_Opt3Req<>"" then
			else
				InvalidGrp3=1
				prd_Opt3Req="0"
			end if
			if IsNumeric(prd_Opt3Req)=false then
				InvalidGrp3=1
				prd_Opt3Req="0"
			end if
			if prd_Opt3Req>"1" then
				InvalidGrp3=1
				prd_Opt3Req="1"
			end if
		end if
		
		if Opt3OrderID<>-1 then
			prd_Opt3Order=trim(CSVRecord(Opt3Orderid))
			if prd_Opt3Order<>"" then
			else
				InvalidGrp3=1
				prd_Opt3Order="3"
			end if
			if IsNumeric(prd_Opt3Order)=false then
				InvalidGrp3=1
				prd_Opt3Order="3"
			end if
			if prd_Opt3Order>"3" then
				prd_Opt3Order="3"
			end if
		end if
		
		if Opt4ID<>-1 then
			prd_Opt4=trim(CSVRecord(Opt4id))
		end if
		
		if Attr4ID<>-1 then
			prd_Attr4=trim(CSVRecord(Attr4id))
		end if
		
		if prd_Attr4<>"" then
		else
			InvalidGrp4=1
		end if

		if Opt4ReqID<>-1 then
			prd_Opt4Req=trim(CSVRecord(Opt4Reqid))
			if prd_Opt4Req<>"" then
			else
				InvalidGrp4=1
				prd_Opt4Req="0"
			end if
			if IsNumeric(prd_Opt4Req)=false then
				InvalidGrp4=1
				prd_Opt4Req="0"
			end if
			if prd_Opt4Req>"1" then
				InvalidGrp4=1
				prd_Opt4Req="1"
			end if
		end if
		
		if Opt4OrderID<>-1 then
			prd_Opt4Order=trim(CSVRecord(Opt4Orderid))
			if prd_Opt4Order<>"" then
			else
				InvalidGrp4=1
				prd_Opt4Order="4"
			end if
			if IsNumeric(prd_Opt4Order)=false then
				InvalidGrp4=1
				prd_Opt4Order="4"
			end if
			if prd_Opt4Order>"4" then
				prd_Opt4Order="4"
			end if
		end if
		
		if Opt5ID<>-1 then
			prd_Opt5=trim(CSVRecord(Opt5id))
		end if
		
		if Attr5ID<>-1 then
			prd_Attr5=trim(CSVRecord(Attr5id))
		end if
		
		if prd_Attr5<>"" then
		else
			InvalidGrp5=1
		end if

		if Opt5ReqID<>-1 then
			prd_Opt5Req=trim(CSVRecord(Opt5Reqid))
			if prd_Opt5Req<>"" then
			else
				InvalidGrp5=1
				prd_Opt5Req="0"
			end if
			if IsNumeric(prd_Opt5Req)=false then
				InvalidGrp5=1
				prd_Opt5Req="0"
			end if
			if prd_Opt5Req>"1" then
				InvalidGrp5=1
				prd_Opt5Req="1"
			end if
		end if
		
		if Opt5OrderID<>-1 then
			prd_Opt5Order=trim(CSVRecord(Opt5Orderid))
			if prd_Opt5Order<>"" then
			else
				InvalidGrp5=1
				prd_Opt5Order="5"
			end if
			if IsNumeric(prd_Opt5Order)=false then
				InvalidGrp5=1
				prd_Opt5Order="5"
			end if
			if prd_Opt5Order>"5" then
				prd_Opt5Order="5"
			end if
		end if
		
		if downprdID<>-1 then
			prd_downprd=trim(CSVRecord(downprdID))
			if prd_downprd<>"" then
			else
				prd_downprd="0"
			end if
			if IsNumeric(prd_downprd)=false then
				prd_downprd="0"
			end if
			if prd_downprd>"1" then
				prd_downprd="1"
			end if
		end if
		
		if giftcertID<>-1 then
			prd_giftcert=trim(CSVRecord(giftcertID))
			if prd_giftcert<>"" then
			else
				prd_giftcert="0"
			end if
			if IsNumeric(prd_giftcert)=false then
				prd_giftcert="0"
			end if
			if prd_giftcert>"1" then
				prd_giftcert="1"
			end if
		end if
		
		if giftexpID<>-1 then
			prd_giftexp=trim(CSVRecord(giftexpID))
			if prd_giftexp<>"" then
			else
				prd_giftexp="0"
			end if
			if IsNumeric(prd_giftexp)=false then
				prd_giftexp="0"
			end if
			if prd_giftexp>"1" then
				prd_giftexp="1"
			end if
		end if
		
		if giftelectID<>-1 then
			prd_giftelect=trim(CSVRecord(giftelectID))
			if prd_giftelect<>"" then
			else
				prd_giftelect="0"
			end if
			if IsNumeric(prd_giftelect)=false then
				prd_giftelect="0"
			end if
			if prd_giftelect>"1" then
				prd_giftelect="1"
			end if
		end if
		
		if giftgenID<>-1 then
			prd_giftgen=trim(CSVRecord(giftgenID))
			if prd_giftgen<>"" then
			else
				prd_giftgen="0"
			end if
			if IsNumeric(prd_giftgen)=false then
				prd_giftgen="0"
			end if
			if prd_giftgen>"1" then
				prd_giftgen="1"
			end if
		end if
		
		if giftexpdateID<>-1 then
			prd_giftexpdate=trim(CSVRecord(giftexpdateID))
		end if
		
		if giftexpdaysID<>-1 then
			prd_giftexpdays=trim(CSVRecord(giftexpdaysID))
			if prd_giftexpdays<>"" then
			else
				prd_giftexpdays="0"
			end if
			if IsNumeric(prd_giftexpdays)=false then
				prd_giftexpdays="0"
			end if
			if prd_giftexpdays>"1" then
				prd_giftexpdays="1"
			end if
		end if
		
		if giftcustgenfileID<>-1 then
			prd_giftcustgenfile=trim(CSVRecord(giftcustgenfileID))
		end if
		
		'BTO
		if hidebtopriceid<>-1 then
			prd_hidebtoprice=trim(CSVRecord(hidebtopriceid))
			if IsNull(prd_hidebtoprice) or prd_hidebtoprice="" then
				prd_hidebtoprice="0"
			end if
			if IsNumeric(prd_hidebtoprice)=false then
				prd_hidebtoprice="0"
			end if
		end if
		
		if hideconfid<>-1 then
			prd_hideconf=trim(CSVRecord(hideconfid))
			if IsNull(prd_hideconf) or prd_hideconf="" then
				prd_hideconf="0"
			end if
			if IsNumeric(prd_hideconf)=false then
				prd_hideconf="0"
			end if
		end if
		
		if dispurchaseid<>-1 then
			prd_dispurchase=trim(CSVRecord(dispurchaseid))
			if IsNull(prd_dispurchase) or prd_dispurchase="" then
				prd_dispurchase="0"
			end if
			if IsNumeric(prd_dispurchase)=false then
				prd_dispurchase="0"
			end if
		end if
		
		if skipdetailsid<>-1 then
			prd_skipdetails=trim(CSVRecord(skipdetailsid))
			if IsNull(prd_skipdetails) or prd_skipdetails="" then
				prd_skipdetails="0"
			end if
			if IsNumeric(prd_skipdetails)=false then
				prd_skipdetails="0"
			end if
		end if
		
		if nameid<>-1 then
		pname=trim(CSVRecord(nameid))
		end if
		
		
		if nameid<>-1 then
		pname=trim(CSVRecord(nameid))
		end if
		if descid<>-1 then
		pdesc=trim(CSVRecord(descid))
		end if
		
		if sdescid<>-1 then
		sdesc=trim(CSVRecord(sdescid))
		end if
		
		if sdesc<>"" then
		sdesc=replace(sdesc,chr(34),"**DD**")
		end if
		
		if opriceid<>-1 then
		poprice=trim(CSVRecord(opriceid))
		end if

		if ptypeid>-1 then
		 pptype=trim(CSVRecord(ptypeid))
		end if
		
		if brandnameid>-1 then
		 brandname=trim(CSVRecord(brandnameid))
		end if
		
		if brandlogoid>-1 then
		 brandlogo=trim(CSVRecord(brandlogoid))
		end if
		
		if fileurlid>-1 then
		 fileurl=trim(CSVRecord(fileurlid))
		end if
		
		if urlexpireid>-1 then
		 urlexpire=trim(CSVRecord(urlexpireid))
		end if
		
		if expiredaysid>-1 then
		 expiredays=trim(CSVRecord(expiredaysid))
		end if
		
		if licenseid>-1 then
		 license=trim(CSVRecord(licenseid))
		end if
		
		if localLGid>-1 then
		 localLG=trim(CSVRecord(localLGid))
		end if
		
		if RemoteLGid>-1 then
		 RemoteLG=trim(CSVRecord(RemoteLGid))
		end if

		if LFN1id>-1 then
		 LFN1=trim(CSVRecord(LFN1id))
		end if
		
		if LFN2id>-1 then
		 LFN2=trim(CSVRecord(LFN2id))
		end if
		
		if LFN3id>-1 then
		 LFN3=trim(CSVRecord(LFN3id))
		end if		
		
		if LFN4id>-1 then
		 LFN4=trim(CSVRecord(LFN4id))
		end if		

		if LFN5id>-1 then
		 LFN5=trim(CSVRecord(LFN5id))
		end if

		if AddCopyid>-1 then
		 AddCopy=trim(CSVRecord(AddCopyid))
		end if
				
		if lpriceid>-1 then
		 plprice=trim(CSVRecord(lpriceid))
		end if
		if wpriceid>-1 then
		 pwprice=trim(CSVRecord(wpriceid))
		end if
		if weightid>-1 then
		 pweight=trim(CSVRecord(weightid))
		end if
		
		if unitslbID>-1 then
		 unitslb=trim(CSVRecord(unitslbID))
		end if
		
		if stockid>-1 then
		 pstock=trim(CSVRecord(stockid))
		end if
		
		if categoryid>-1 then
		 pcategory=trim(CSVRecord(categoryid))
		end if
		
		if SCATDescid>-1 then
		SCATDesc=trim(CSVRecord(SCATDescid))
		end if
		
		if LCATDescid>-1 then
		LCATDesc=trim(CSVRecord(LCATDescid))
		end if
		
		if csimageid>-1 then
		 pcsimage=trim(CSVRecord(csimageid))
		 if pcsimage<>"" then
		 if Left(pcsimage,1)="/" then
		 pcsimage=mid(pcsimage,2,len(pcsimage))
		 end if
		 end if
		end if
		
		if climageid>-1 then
		 pclimage=trim(CSVRecord(climageid))
		 if pclimage<>"" then
		 if Left(pclimage,1)="/" then
		 pclimage=mid(pclimage,2,len(pclimage))
		 end if
		 end if
		end if
		
		if pcategoryid>-1 then
		 ppcategory=trim(CSVRecord(pcategoryid))
		end if
		
		if category1id>-1 then
		 pcategory1=trim(CSVRecord(category1id))
		end if
		
		if SCATDesc1id>-1 then
		SCATDesc1=trim(CSVRecord(SCATDesc1id))
		end if
		
		if LCATDesc1id>-1 then
		LCATDesc1=trim(CSVRecord(LCATDesc1id))
		end if
		
		if csimage1id>-1 then
			pcsimage1=trim(CSVRecord(csimage1id))
			if pcsimage1<>"" then
			if Left(pcsimage1,1)="/" then
				pcsimage1=mid(pcsimage1,2,len(pcsimage1))
			end if
			end if
		end if
		
		if climage1id>-1 then
			pclimage1=trim(CSVRecord(climage1id))
			if pclimage1<>"" then
			if Left(pclimage1,1)="/" then
				pclimage1=mid(pclimage1,2,len(pclimage1))
			end if
			end if
		end if
		
		if pcategory1id>-1 then
		 ppcategory1=trim(CSVRecord(pcategory1id))
		end if
		
		if category2id>-1 then
		 pcategory2=trim(CSVRecord(category2id))
		end if
		
		if SCATDesc2id>-1 then
			SCATDesc2=trim(CSVRecord(SCATDesc2id))
		end if
		
		if LCATDesc2id>-1 then
			LCATDesc2=trim(CSVRecord(LCATDesc2id))
		end if
		
		if csimage2id>-1 then
			pcsimage2=trim(CSVRecord(csimage2id))
			if pcsimage2<>"" then
			if Left(pcsimage2,1)="/" then
				pcsimage2=mid(pcsimage2,2,len(pcsimage2))
			end if
			end if
		end if
		
		if climage2id>-1 then
			pclimage2=trim(CSVRecord(climage2id))
			if pclimage2<>"" then
			if Left(pclimage2,1)="/" then
				pclimage2=mid(pclimage2,2,len(pclimage2))
			end if
			end if
		end if
		
		if pcategory2id>-1 then
		 ppcategory2=trim(CSVRecord(pcategory2id))
		end if
						
		if timageid>-1 then
		 ptimage=trim(CSVRecord(timageid))
		 if ptimage<>"" then
		 if Left(ptimage,1)="/" then
		 ptimage=mid(ptimage,2,len(ptimage))
		 end if
		 end if
		end if
		if gimageid>-1 then
		 pgimage=trim(CSVRecord(gimageid))
		 if pgimage<>"" then
		 if Left(pgimage,1)="/" then
		 pgimage=mid(pgimage,2,len(pgimage))
		 end if
		 end if
		end if
		if dimageid>-1 then
		 pdimage=trim(CSVRecord(dimageid))
		 if pdimage<>"" then
		 if Left(pdimage,1)="/" then
		 pdimage=mid(pdimage,2,len(pdimage))
		 end if
		 end if
		end if
		if activeid>-1 then
		 pactive=trim(CSVRecord(activeid))
		end if
		if savingid>-1 then
		 psaving=trim(CSVRecord(savingid))
		end if
		if specialid>-1 then
		 pspecial=trim(CSVRecord(specialid))
		end if
		
		if featuredid>-1 then
		 pfeatured=trim(CSVRecord(featuredid))
		end if
		
		if rwpid>-1 then
		 prwp=trim(CSVRecord(rwpid))
		end if
		
		if ntaxid>-1 then
		 pntax=trim(CSVRecord(ntaxid))
		end if
		if nshipid>-1 then
		 pnship=trim(CSVRecord(nshipid))
		end if
		if nforsaleid>-1 then
		 pnforsale=trim(CSVRecord(nforsaleid))
		end if
		if nforsalecopyid>-1 then
		pnforsalecopy=trim(CSVRecord(nforsalecopyid))
		end if
		if distockid>-1 then
		 distock=trim(CSVRecord(distockid))
		end if
		if dishiptextid>-1 then
		 dishiptext=trim(CSVRecord(dishiptextid))
		end if
		if customfieldsid(0)>-1 then
		customfields(0)=trim(CSVRecord(customfieldsid(0)))
		end if
		if customfieldsid(1)>-1 then
		customfields(1)=trim(CSVRecord(customfieldsid(1)))
		end if
		if customfieldsid(2)>-1 then
		customfields(2)=trim(CSVRecord(customfieldsid(2)))
		end if
		
		if MQtyID>-1 then
		MQty=trim(CSVRecord(MQtyid))
		end if
		
		if MQty<>"" then
		else
		MQty="0"
		end if
		
		if VQtyID>-1 then
		VQty=trim(CSVRecord(VQtyID))
		end if
		
		if VQty<>"" then
		else
		VQty="0"
		end if
		
		if OverSizeID>-1 then
		OverSize=trim(CSVRecord(OverSizeid))
		end if
		
		if OverSize<>"" then
			OSArray=split(OverSize,"x")
			if Ubound(OSArray)<2 then
				OverSize="NO"
			else
				mTest=0
				For v=lbound(OSArray) to ubound(OSArray)
					if IsNumeric(OSArray(v))=false then
						mTest=1
						exit for
					end if
				Next
				if mTest=1 then
					OverSize="NO"
				else
					OverSize=replace(OverSize,"x","||")
					For m=Ubound(OSArray)+1 to 4
						OverSize=OverSize & "||0"
					Next
				end if
			end if	
		else
			OverSize="NO"
		end if
		
		if not IsNumeric(unitslb) then
			unitslb=0
		end if		

		if pdesc="" then
		pdesc="no information"
		end if
		if poprice="" then 
		poprice="0"
		end if
		if plprice="" then
		plprice="0"
		end if
		if pwprice="" then
		pwprice="0"
		end if
		if pweight="" then
		pweight="0"
		end if
		if pstock="" then
		pstock="0"
		end if
		if pgimage="" then
		pgimage="no_image.gif"
		end if

		if pactive="" then
		pactive="-1"
		end if
		if psaving="" then
		psaving="-1"
		end if
		if pspecial="" then
		pspecial="0"
		end if
		if pfeatured="" then
			pfeatured="0"
		end if		
		if prwp="" then
		prwp="0"
		end if
		if pntax="" then
		pntax="0"
		end if
		if pnship="" then
		pnship="0"
		end if
		if pnforsale="" then
		pnforsale="0"
		end if
		if pnforsalecopy="" then
		pnforsalecopy="no"
		end if
		
		if distock="" then
		distock="0"
		end if
		
		if dishiptext="" then
		dishiptext="0"
		end if
		
		if urlexpire<>"" then
		else
		urlexpire="0"
		end if
		if expiredays<>"" then
		else
		expiredays="0"
		end if

		if license<>"" then
		else
		license="0"
		end if		
		
		if psku="" then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include a Product SKU." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if session("append")<>"1" then
		if pname="" then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include a Product Name." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if pdesc="" then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include a Product Description." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		end if
		if isNumeric(poprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The Online Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(plprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The List Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pwprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The Wholesale Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pweight)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The Weight is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pstock)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The Stock level is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		
		IF PPTYPE<>"" THEN
		if ucase(pptype)="DP" then
		
		if session("append")<>"1" then
			if fileurl<>"" then
			else
				ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include a Downloadable File Location." & "</td></tr>" & vbcrlf
				RecordError=true
			end if
		end if
		
		if isNumeric(urlexpire)=false then
			ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The Make Download URL expire field must be a number." & "</td></tr>" & vbcrlf
			RecordError=true
		else
			if cint(urlexpire)=1 then
				if isNumeric(expiredays)=false then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The URL expiration days value is not a number." & "</td></tr>" & vbcrlf
					RecordError=true
				else
					if Cint(expiredays)=0 then
						ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The URL expiration days value must be greater than zero." & "</td></tr>" & vbcrlf
						RecordError=true
					end if
				end if
			end if
		end if
		
		if isNumeric(license)=false then
			ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": The License Generator is not a number." & "</td></tr>" & vbcrlf
			RecordError=true
		else
			if Cint(license)=1 then
				if LocalLG & RemoteLG<>"" then
					if (LocalLG<>"") and (RemoteLG<>"") then
						ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": contains both a Local License Generator and a Remote License Generator. Only one of the two should be included." & "</td></tr>" & vbcrlf
						RecordError=true
					end if
				else
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include a License Generator." & "</td></tr>" & vbcrlf
					RecordError=true
				end if
		
				if LFN1&LFN2&LFN3&LFN4&LFN5<>"" then
				else
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalCSVlines & ": does not include any License Field Label." & "</td></tr>" & vbcrlf
					RecordError=true
				end if
		
			end if
		
		end if
		
		end if
		END IF
		
		if scDecSign="," then
			poprice=replace(poprice,".","")
			plprice=replace(plprice,".","")
			pwprice=replace(pwprice,".","")
			
			poprice=replace(poprice,",",".")
			plprice=replace(plprice,",",".")
			pwprice=replace(pwprice,",",".")
		else
			poprice=replace(poprice,",","")
			plprice=replace(plprice,",","")
			pwprice=replace(pwprice,",","")
		end if
		
		poprice=replace(poprice,scCurSign,"")
		plprice=replace(plprice,scCurSign,"")
		pwprice=replace(pwprice,scCurSign,"")
		
		if psku<>"" then
			psku=replace(psku,"""","&quot;")
		end if
%>