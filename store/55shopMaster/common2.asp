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

		psku=trim(rsExcel.Fields.Item(int(skuid)).Value)
		
		if Left(psku,1)="'" then
			psku=mid(psku,2,len(psku))
			if Left(psku,1)="'" then
				psku=mid(psku,2,len(psku))
			end if
		end if
		
		if surcharge1ID<>-1 then
			surcharge1=trim(rsExcel.Fields.Item(int(surcharge1ID)).Value)
			if (not (surcharge1<>"")) or (isNumeric(surcharge1)=false) then
				surcharge1="0"
			end if
		end if
		
		if surcharge2ID<>-1 then
			surcharge2=trim(rsExcel.Fields.Item(int(surcharge2ID)).Value)
			if (not (surcharge2<>"")) or (isNumeric(surcharge2)=false) then
				surcharge2="0"
			end if
		end if
		
		if goCatid<>-1 then
			goCat=trim(rsExcel.Fields.Item(int(goCatid)).Value)
			if goCat<>"" then
				goCat=replace(goCat,"'","''")
			end if
		end if
		
		if goGenid<>-1 then
			goGen=trim(rsExcel.Fields.Item(int(goGenid)).Value)
			if goGen<>"" then
				goGen=replace(goGen,"'","''")
			end if
		end if
		
		if goAgeid<>-1 then
			goAge=trim(rsExcel.Fields.Item(int(goAgeid)).Value)
			if goAge<>"" then
				goAge=replace(goAge,"'","''")
			end if
		end if
		
		if goColorid<>-1 then
			goColor=trim(rsExcel.Fields.Item(int(goColorid)).Value)
			if goColor<>"" then
				goColor=replace(goColor,"'","''")
			end if
		end if
		
		if goSizeid<>-1 then
			goSize=trim(rsExcel.Fields.Item(int(goSizeid)).Value)
			if goSize<>"" then
				goSize=replace(goSize,"'","''")
			end if
		end if
		
		if goPatid<>-1 then
			goPat=trim(rsExcel.Fields.Item(int(goPatid)).Value)
			if goPat<>"" then
				goPat=replace(goPat,"'","''")
			end if
		end if
		
		if goMatid<>-1 then
			goMat=trim(rsExcel.Fields.Item(int(goMatid)).Value)
			if goMat<>"" then
				goMat=replace(goMat,"'","''")
			end if
		end if

		if prdnoteID<>-1 then
			prdnote=trim(rsExcel.Fields.Item(int(prdnoteID)).Value)
			if prdnote<>"" then
				prdnote=replace(prdnote,"'","''")
			end if
		end if
		
		if playoutID<>-1 then
			playout=trim(rsExcel.Fields.Item(int(playoutID)).Value)
			if playout<>"" then
				playout=replace(playout,"'","''")
				Select Case playout
					Case "l","c","o":
					Case Else: playout=""
				End Select
			end if
		end if
		
		if eimagID<>-1 then
			eimag=trim(rsExcel.Fields.Item(int(eimagID)).Value)
			if (not (eimag<>"")) or (isNumeric(eimag)=false) then
				eimag="0"
			end if
		end if
		
		if hideskuID<>-1 then
			hidesku=trim(rsExcel.Fields.Item(int(hideskuID)).Value)
			if (not (hidesku<>"")) or (isNumeric(hidesku)=false) then
				hidesku="0"
			end if
		end if
		
		'Start SDBA
		if prd_CostID<>-1 then
			prd_Cost=trim(rsExcel.Fields.Item(int(prd_CostID)).Value)
			if (not (prd_Cost<>"")) or (isNumeric(prd_Cost)=false) then
				prd_Cost="0"
			end if
		end if
		
		if prd_BackOrderID<>-1 then
			prd_BackOrder=trim(rsExcel.Fields.Item(int(prd_BackOrderID)).Value)
			if (not (prd_BackOrder<>"")) or (isNumeric(prd_BackOrder)=false) then
				prd_BackOrder="0"
			end if
		end if
		
		if prd_ShipNDaysID<>-1 then
			prd_ShipNDays=trim(rsExcel.Fields.Item(int(prd_ShipNDaysID)).Value)
			if (not (prd_ShipNDays<>"")) or (isNumeric(prd_ShipNDays)=false) then
				prd_ShipNDays="0"
			end if
		end if
		
		if prd_NotifyStockID<>-1 then
			prd_NotifyStock=trim(rsExcel.Fields.Item(int(prd_NotifyStockID)).Value)
			if (not (prd_NotifyStock<>"")) or (isNumeric(prd_NotifyStock)=false) then
				prd_NotifyStock="0"
			end if
		end if
		
		if prd_ReorderLevelID<>-1 then
			prd_ReorderLevel=trim(rsExcel.Fields.Item(int(prd_ReorderLevelID)).Value)
			if (not (prd_ReorderLevel<>"")) or (isNumeric(prd_ReorderLevel)=false) then
				prd_ReorderLevel="0"
			end if
		end if
		
		if prd_IsDropShippedID<>-1 then
			prd_IsDropShipped=trim(rsExcel.Fields.Item(int(prd_IsDropShippedID)).Value)
			if (not (prd_IsDropShipped<>"")) or (isNumeric(prd_IsDropShipped)=false) then
				prd_IsDropShipped="0"
			end if
		end if
		
		if prd_SupplierID<>-1 then
			prd_Supplier=trim(rsExcel.Fields.Item(int(prd_SupplierID)).Value)
			if (not (prd_Supplier<>"")) or (isNumeric(prd_Supplier)=false) then
				prd_Supplier="0"
			end if
		end if
		
		if prd_DropShipperID<>-1 then
			prd_DropShipper=trim(rsExcel.Fields.Item(int(prd_DropShipperID)).Value)
			if (not (prd_DropShipper<>"")) or (isNumeric(prd_DropShipper)=false) then
				prd_DropShipper="0"
			end if
		end if
		
		if prd_IsDropShipperID<>-1 then
			prd_IsDropShipper=trim(rsExcel.Fields.Item(int(prd_IsDropShipperID)).Value)
			if (not (prd_IsDropShipper<>"")) or (isNumeric(prd_IsDropShipper)=false) then
				prd_IsDropShipper="0"
			end if
		end if
		'End SDBA
		
		if mt_titleID<>-1 then
			mt_title=trim(rsExcel.Fields.Item(int(mt_titleID)).Value)
			if mt_title<>"" then
				mt_title=replace(mt_title,"'","''")
			end if
		end if
		
		if mt_descID<>-1 then
			mt_desc=trim(rsExcel.Fields.Item(int(mt_descID)).Value)
			if mt_desc<>"" then
				mt_desc=replace(mt_desc,"'","''")
			end if
		end if
		
		if mt_keyID<>-1 then
			mt_key=trim(rsExcel.Fields.Item(int(mt_keyID)).Value)
			if mt_key<>"" then
				mt_key=replace(mt_key,"'","''")
			end if
		end if
		
		if Opt1ID<>-1 then
		prd_Opt1=trim(rsExcel.Fields.Item(int(Opt1ID)).Value)
		end if
		
		if prd_Opt1<>"" then
		prd_Opt1=replace(prd_Opt1,"'","''")
		end if
		
		if Opt2ID<>-1 then
		prd_Opt2=trim(rsExcel.Fields.Item(int(Opt2ID)).Value)
		end if
		
		if prd_Opt2<>"" then
		prd_Opt2=replace(prd_Opt2,"'","''")
		end if
		
		if Attr1ID<>-1 then
		prd_Attr1=trim(rsExcel.Fields.Item(int(Attr1ID)).Value)
		end if
		
		if prd_Attr1<>"" then
		prd_Attr1=replace(prd_Attr1,"'","''")
		else
		InvalidGrp1=1
		end if
		
		if Attr2ID<>-1 then
		prd_Attr2=trim(rsExcel.Fields.Item(int(Attr2ID)).Value)
		end if
		
		if prd_Attr2<>"" then
		prd_Attr2=replace(prd_Attr2,"'","''")
		else
		InvalidGrp2=1
		end if
		
		if Opt1ReqID<>-1 then
		prd_Opt1Req=trim(rsExcel.Fields.Item(int(Opt1ReqID)).Value)
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
		prd_Opt2Req=trim(rsExcel.Fields.Item(int(Opt2ReqID)).Value)
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
			prd_Opt1Order=trim(rsExcel.Fields.Item(int(Opt1Orderid)).Value)
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
			prd_Opt2Order=trim(rsExcel.Fields.Item(int(Opt2Orderid)).Value)
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
			prd_Opt3=trim(rsExcel.Fields.Item(int(Opt3id)).Value)
		end if
		
		if prd_Opt3<>"" then
			prd_Opt3=replace(prd_Opt3,"'","''")
		end if

		if Attr3ID<>-1 then
			prd_Attr3=trim(rsExcel.Fields.Item(int(Attr3id)).Value)
		end if
		
		if prd_Attr3<>"" then
			prd_Attr3=replace(prd_Attr3,"'","''")
		else
			InvalidGrp3=1
		end if

		if Opt3ReqID<>-1 then
			prd_Opt3Req=trim(rsExcel.Fields.Item(int(Opt3Reqid)).Value)
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
			prd_Opt3Order=trim(rsExcel.Fields.Item(int(Opt3Orderid)).Value)
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
			prd_Opt4=trim(rsExcel.Fields.Item(int(Opt4id)).Value)
		end if
		
		if prd_Opt4<>"" then
			prd_Opt4=replace(prd_Opt4,"'","''")
		end if

		if Attr4ID<>-1 then
			prd_Attr4=trim(rsExcel.Fields.Item(int(Attr4id)).Value)
		end if
		
		if prd_Attr4<>"" then
			prd_Attr4=replace(prd_Attr4,"'","''")
		else
			InvalidGrp4=1
		end if

		if Opt4ReqID<>-1 then
			prd_Opt4Req=trim(rsExcel.Fields.Item(int(Opt4Reqid)).Value)
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
			prd_Opt4Order=trim(rsExcel.Fields.Item(int(Opt4Orderid)).Value)
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
			prd_Opt5=trim(rsExcel.Fields.Item(int(Opt5id)).Value)
		end if
		
		if prd_Opt5<>"" then
			prd_Opt5=replace(prd_Opt5,"'","''")
		end if

		if Attr5ID<>-1 then
			prd_Attr5=trim(rsExcel.Fields.Item(int(Attr5id)).Value)
		end if
		
		if prd_Attr5<>"" then
			prd_Attr5=replace(prd_Attr5,"'","''")
		else
			InvalidGrp5=1
		end if

		if Opt5ReqID<>-1 then
			prd_Opt5Req=trim(rsExcel.Fields.Item(int(Opt5Reqid)).Value)
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
			prd_Opt5Order=trim(rsExcel.Fields.Item(int(Opt5Orderid)).Value)
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
			prd_downprd=trim(rsExcel.Fields.Item(int(downprdID)).Value)
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
			prd_giftcert=trim(rsExcel.Fields.Item(int(giftcertID)).Value)
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
			prd_giftexp=trim(rsExcel.Fields.Item(int(giftexpID)).Value)
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
			prd_giftelect=trim(rsExcel.Fields.Item(int(giftelectID)).Value)
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
			prd_giftgen=trim(rsExcel.Fields.Item(int(giftgenID)).Value)
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
			prd_giftexpdate=trim(rsExcel.Fields.Item(int(giftexpdateID)).Value)
			if prd_giftexpdate<>"" then
				prd_giftexpdate=replace(prd_giftexpdate,"'","''")
			end if
		end if
		
		if giftexpdaysID<>-1 then
			prd_giftexpdays=trim(rsExcel.Fields.Item(int(giftexpdaysID)).Value)
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
			prd_giftcustgenfile=trim(rsExcel.Fields.Item(int(giftcustgenfileID)).Value)
			if prd_giftcustgenfile<>"" then
				prd_giftcustgenfile=replace(prd_giftcustgenfile,"'","''")
			end if
		end if
		
		'BTO
		if hidebtopriceid<>-1 then
			prd_hidebtoprice=trim(rsExcel.Fields.Item(int(hidebtopriceid)).Value)
			if IsNull(prd_hidebtoprice) or prd_hidebtoprice="" then
				prd_hidebtoprice="0"
			end if
			if IsNumeric(prd_hidebtoprice)=false then
				prd_hidebtoprice="0"
			end if
		end if
		
		if hideconfid<>-1 then
			prd_hideconf=trim(rsExcel.Fields.Item(int(hideconfid)).Value)
			if IsNull(prd_hideconf) or prd_hideconf="" then
				prd_hideconf="0"
			end if
			if IsNumeric(prd_hideconf)=false then
				prd_hideconf="0"
			end if
		end if
		
		if dispurchaseid<>-1 then
			prd_dispurchase=trim(rsExcel.Fields.Item(int(dispurchaseid)).Value)
			if IsNull(prd_dispurchase) or prd_dispurchase="" then
				prd_dispurchase="0"
			end if
			if IsNumeric(prd_dispurchase)=false then
				prd_dispurchase="0"
			end if
		end if
		
		if skipdetailsid<>-1 then
			prd_skipdetails=trim(rsExcel.Fields.Item(int(skipdetailsid)).Value)
			if IsNull(prd_skipdetails) or prd_skipdetails="" then
				prd_skipdetails="0"
			end if
			if IsNumeric(prd_skipdetails)=false then
				prd_skipdetails="0"
			end if
		end if
		
		if nameid<>-1 then
		pname=trim(rsExcel.Fields.Item(int(nameid)).Value)
		end if
		if descid<>-1 then
		pdesc=trim(rsExcel.Fields.Item(int(descid)).Value)
		end if
		
		if sdescid<>-1 then
		sdesc=trim(rsExcel.Fields.Item(int(sdescid)).Value)
		end if
		if sdesc<>"" then
		sdesc=replace(sdesc,chr(34),"**DD**")
		end if
		
		if opriceid<>-1 then
		poprice=trim(rsExcel.Fields.Item(int(opriceid)).Value)
		end if

		if ptypeid>-1 then
		 pptype=trim(rsExcel.Fields.Item(int(ptypeid)).Value)
		end if
		
		if brandnameid>-1 then
		 brandname=trim(rsExcel.Fields.Item(int(brandnameid)).Value)
		end if
		
		if brandlogoid>-1 then
		 brandlogo=trim(rsExcel.Fields.Item(int(brandlogoid)).Value)
		end if
		
		if fileurlid>-1 then
		 fileurl=trim(rsExcel.Fields.Item(int(fileurlid)).Value)
		end if
		
		if urlexpireid>-1 then
		 urlexpire=trim(rsExcel.Fields.Item(int(urlexpireid)).Value)
		end if
		
		if expiredaysid>-1 then
		 expiredays=trim(rsExcel.Fields.Item(int(expiredaysid)).Value)
		end if
		
		if licenseid>-1 then
		 license=trim(rsExcel.Fields.Item(int(licenseid)).Value)
		end if
		
		if localLGid>-1 then
		 localLG=trim(rsExcel.Fields.Item(int(localLGid)).Value)
		end if
		
		if RemoteLGid>-1 then
		 RemoteLG=trim(rsExcel.Fields.Item(int(RemoteLGid)).Value)
		end if

		if LFN1id>-1 then
		 LFN1=trim(rsExcel.Fields.Item(int(LFN1id)).Value)
		end if
		
		if LFN2id>-1 then
		 LFN2=trim(rsExcel.Fields.Item(int(LFN2id)).Value)
		end if
		
		if LFN3id>-1 then
		 LFN3=trim(rsExcel.Fields.Item(int(LFN3id)).Value)
		end if		
		
		if LFN4id>-1 then
		 LFN4=trim(rsExcel.Fields.Item(int(LFN4id)).Value)
		end if		

		if LFN5id>-1 then
		 LFN5=trim(rsExcel.Fields.Item(int(LFN5id)).Value)
		end if

		if AddCopyid>-1 then
		 AddCopy=trim(rsExcel.Fields.Item(int(AddCopyid)).Value)
		end if


		if lpriceid>-1 then
		 plprice=trim(rsExcel.Fields.Item(int(lpriceid)).Value)
		end if
		if wpriceid>-1 then
		 pwprice=trim(rsExcel.Fields.Item(int(wpriceid)).Value)
		end if
		if weightid>-1 then
		 pweight=trim(rsExcel.Fields.Item(int(weightid)).Value)
		end if
		if unitslbID>-1 then
		 unitslb=trim(rsExcel.Fields.Item(int(unitslbID)).Value)
		end if
		if stockid>-1 then
		 pstock=trim(rsExcel.Fields.Item(int(stockid)).Value)
		end if
		
		if categoryid>-1 then
		 pcategory=trim(rsExcel.Fields.Item(int(categoryid)).Value)
		end if
		
		if SCATDescid>-1 then
		 SCATDesc=trim(rsExcel.Fields.Item(int(SCATDescid)).Value)
		end if
		
		if LCATDescid>-1 then
		 LCATDesc=trim(rsExcel.Fields.Item(int(LCATDescid)).Value)
		end if
		
		if csimageid>-1 then
		if rsExcel.Fields.Item(int(csimageid)).Value<>"" then
		 pcsimage=trim(rsExcel.Fields.Item(int(csimageid)).Value)
		 if Left(pcsimage,1)="/" then
		 pcsimage=mid(pcsimage,2,len(pcsimage))
		 end if
		 end if
		end if
		
		if climageid>-1 then
		if rsExcel.Fields.Item(int(climageid)).Value<>"" then
		 pclimage=trim(rsExcel.Fields.Item(int(climageid)).Value)
		 if Left(pclimage,1)="/" then
		 pclimage=mid(pclimage,2,len(pclimage))
		 end if
		 end if
		end if
		
		if pcategoryid>-1 then
		 ppcategory=trim(rsExcel.Fields.Item(int(pcategoryid)).Value)
		end if
		
		if category1id>-1 then
		 pcategory1=trim(rsExcel.Fields.Item(int(category1id)).Value)
		end if
		
		if SCATDesc1id>-1 then
		SCATDesc1=trim(rsExcel.Fields.Item(int(SCATDesc1id)).Value)
		end if
		
		if LCATDesc1id>-1 then
		LCATDesc1=trim(rsExcel.Fields.Item(int(LCATDesc1id)).Value)
		end if
		
		if csimage1id>-1 then
			pcsimage1=trim(rsExcel.Fields.Item(int(csimage1id)).Value)
			if pcsimage1<>"" then
			if Left(pcsimage1,1)="/" then
				pcsimage1=mid(pcsimage1,2,len(pcsimage1))
			end if
			end if
		end if
		
		if climage1id>-1 then
			pclimage1=trim(rsExcel.Fields.Item(int(climage1id)).Value)
			if pclimage1<>"" then
			if Left(pclimage1,1)="/" then
				pclimage1=mid(pclimage1,2,len(pclimage1))
			end if
			end if
		end if
		
		if pcategory1id>-1 then
		 ppcategory1=trim(rsExcel.Fields.Item(int(pcategory1id)).Value)
		end if
		
		if category2id>-1 then
		 pcategory2=trim(rsExcel.Fields.Item(int(category2id)).Value)
		end if
		
		if SCATDesc2id>-1 then
		SCATDesc2=trim(rsExcel.Fields.Item(int(SCATDesc2id)).Value)
		end if
		
		if LCATDesc2id>-1 then
		LCATDesc2=trim(rsExcel.Fields.Item(int(LCATDesc2id)).Value)
		end if
		
		if csimage2id>-1 then
			pcsimage2=trim(rsExcel.Fields.Item(int(csimage2id)).Value)
			if pcsimage2<>"" then
			if Left(pcsimage2,1)="/" then
				pcsimage2=mid(pcsimage2,2,len(pcsimage2))
			end if
			end if
		end if
		
		if climage2id>-1 then
			pclimage2=trim(rsExcel.Fields.Item(int(climage2id)).Value)
			if pclimage2<>"" then
			if Left(pclimage2,1)="/" then
				pclimage2=mid(pclimage2,2,len(pclimage2))
			end if
			end if
		end if
		
		if pcategory2id>-1 then
		 ppcategory2=trim(rsExcel.Fields.Item(int(pcategory2id)).Value)
		end if		
		
		if timageid>-1 then
		if rsExcel.Fields.Item(int(timageid)).Value<>"" then
		 ptimage=trim(rsExcel.Fields.Item(int(timageid)).Value)
		 if Left(ptimage,1)="/" then
		 ptimage=mid(ptimage,2,len(ptimage))
		 end if
		end if
		end if
		if gimageid>-1 then
		if rsExcel.Fields.Item(int(gimageid)).Value<>"" then
		 pgimage=trim(rsExcel.Fields.Item(int(gimageid)).Value)
		 if Left(pgimage,1)="/" then
		 pgimage=mid(pgimage,2,len(pgimage))
		 end if
		end if
		end if
		if dimageid>-1 then
		if rsExcel.Fields.Item(int(dimageid)).Value<>"" then
		 pdimage=trim(rsExcel.Fields.Item(int(dimageid)).Value)
		 if Left(pdimage,1)="/" then
		 pdimage=mid(pdimage,2,len(pdimage))
		 end if
		end if
		end if
		if activeid>-1 then
		 pactive=trim(rsExcel.Fields.Item(int(activeid)).Value)
		end if
		if savingid>-1 then
		 psaving=trim(rsExcel.Fields.Item(int(savingid)).Value)
		end if
		if specialid>-1 then
		 pspecial=trim(rsExcel.Fields.Item(int(specialid)).Value)
		end if
		
		if featuredid>-1 then
			pfeatured=trim(rsExcel.Fields.Item(int(featuredid)).Value)
		end if
		
		if rwpid>-1 then
		 prwp=trim(rsExcel.Fields.Item(int(rwpid)).Value)
		end if
		if ntaxid>-1 then
		 pntax=trim(rsExcel.Fields.Item(int(ntaxid)).Value)
		end if
		if nshipid>-1 then
		 pnship=trim(rsExcel.Fields.Item(int(nshipid)).Value)
		end if
		if nforsaleid>-1 then
		 pnforsale=trim(rsExcel.Fields.Item(int(nforsaleid)).Value)
		end if
		if nforsalecopyid>-1 then
		pnforsalecopy=trim(rsExcel.Fields.Item(int(nforsalecopyid)).Value)
		end if
		if distockid>-1 then
		distock=trim(rsExcel.Fields.Item(int(distockid)).Value)
		end if
		if dishiptextid>-1 then
		dishiptext=trim(rsExcel.Fields.Item(int(dishiptextid)).Value)
		end if
		if customfieldsid(0)>-1 then
		customfields(0)=trim(rsExcel.Fields.Item(int(customfieldsid(0))).Value)
		end if
		if customfieldsid(1)>-1 then
		customfields(1)=trim(rsExcel.Fields.Item(int(customfieldsid(1))).Value)
		end if
		if customfieldsid(2)>-1 then
		customfields(2)=trim(rsExcel.Fields.Item(int(customfieldsid(2))).Value)
		end if
		
		if MQtyID>-1 then
		 MQty=trim(rsExcel.Fields.Item(int(MQtyID)).Value)
		end if
		
		if MQty<>"" then
		else
		MQty="0"
		end if
		
		if VQtyID>-1 then
		 VQty=trim(rsExcel.Fields.Item(int(VQtyID)).Value)
		end if
		
		if VQty<>"" then
		else
		VQty="0"
		end if
		
		if OverSizeID>-1 then
		OverSize=trim(rsExcel.Fields.Item(int(OverSizeID)).Value)
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

		if pdesc<>"" then
		else
		pdesc="no information"
		end if
		if poprice<>"" then 
		else
		poprice="0"
		end if
		if plprice<>"" then
		else
		plprice="0"
		end if
		if pwprice<>"" then
		else
		pwprice="0"
		end if
		if pweight<>"" then
		else
		pweight="0"
		end if
		if pstock<>"" then
		else
		pstock="0"
		end if
		if pgimage<>"" then
		else
		pgimage="no_image.gif"
		end if

		if pactive<>"" then
		else
		pactive="-1"
		end if
		if psaving<>"" then
		else
		psaving="-1"
		end if
		if pspecial<>"" then
		else
		pspecial="0"
		end if
		if pfeatured<>"" then
		else
			pfeatured="0"
		end if
		if prwp<>"" then
		else
		prwp="0"
		end if
		if pntax<>"" then
		else
		pntax="0"
		end if
		if pnship<>"" then
		else
		pnship="0"
		end if
		if pnforsale<>"" then
		else
		pnforsale="0"
		end if
		if pnforsalecopy<>"" then
		else
		pnforsalecopy="no"
		end if
		if distock<>"" then
		else
		distock="0"
		end if
		if dishiptext<>"" then
		else
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

		
		if psku<>"" then
		else
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include a Product SKU." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if session("append")<>"1" then
		if pname<>"" then
		else
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include a Product Name." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if pdesc<>"" then
		else
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include a Product Description." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		end if
		if isNumeric(poprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The Online Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(plprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The List Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pwprice)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The Wholesale Price is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pweight)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The Weight is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		if isNumeric(pstock)=false then
		ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The Stock level is not a number." & "</td></tr>" & vbcrlf
		RecordError=true
		end if
		
		IF PPTYPE<>"" THEN
		if ucase(pptype)="DP" then
		
		if session("append")<>"1" then
			if fileurl<>"" then
			else
				ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include a Downloadable File Location." & "</td></tr>" & vbcrlf
				RecordError=true
			end if
		end if
		
		if isNumeric(urlexpire)=false then
			ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The Make Download URL Expire field must be a number." & "</td></tr>" & vbcrlf
			RecordError=true
		else
			if cint(urlexpire)=1 then
				if isNumeric(expiredays)=false then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The URL Expiration Days value is not a number." & "</td></tr>" & vbcrlf
					RecordError=true
				else
					if Cint(expiredays)=0 then
						ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The URL Expiration Days value must be greater than zero." & "</td></tr>" & vbcrlf
						RecordError=true
					end if
				end if
			end if
		end if
		
		if isNumeric(license)=false then
			ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The License Generator is not a number." & "</td></tr>" & vbcrlf
			RecordError=true
		else
			if Cint(license)=1 then
				if LocalLG & RemoteLG<>"" then
					if (LocalLG<>"") and (RemoteLG<>"") then
						ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": includes both a Local License Generator and a Remote License Generator. Only one of the two should be specified." & "</td></tr>" & vbcrlf
						RecordError=true
					end if
				else
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include a License Generator." & "</td></tr>" & vbcrlf
					RecordError=true
				end if
		
				if LFN1&LFN2&LFN3&LFN4&LFN5<>"" then
				else
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": does not include any License Field Label." & "</td></tr>" & vbcrlf
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
		
		if pname<>"" then
		pname=replace(pname,"'","''")
		end if
		
		if pdesc<>"" then
		pdesc=replace(pdesc,"'","''")
		end if
		
		if sdesc<>"" then
		sdesc=replace(sdesc,"'","''")
		end if
					
		if brandname<>"" then
		brandname=replace(brandname,"'","''")
		end if

		if pnforsalecopy<>"" then
		pnforsalecopy=replace(pnforsalecopy,"'","''")
		end if

		if psku<>"" then
		psku=replace(psku,"'","''")
		psku=replace(psku,"""","&quot;")
		end if
		
		if pcategory<>"" then
		pcategory=replace(pcategory,"'","''")
		end if
		
		if SCATDesc<>"" then
		SCATDesc=replace(SCATDesc,"'","''")
		end if
		
		if LCATDesc<>"" then
		LCATDesc=replace(LCATDesc,"'","''")
		end if
		
		if ppcategory<>"" then
		ppcategory=replace(ppcategory,"'","''")
		end if
				
		if pcategory1<>"" then
		pcategory1=replace(pcategory1,"'","''")
		end if
		
		if SCATDesc1<>"" then
		SCATDesc1=replace(SCATDesc1,"'","''")
		end if
		
		if LCATDesc1<>"" then
		LCATDesc1=replace(LCATDesc1,"'","''")
		end if
		
		if ppcategory1<>"" then
		ppcategory1=replace(ppcategory1,"'","''")
		end if		
		
		if pcategory2<>"" then
		pcategory2=replace(pcategory2,"'","''")
		end if
		
		if SCATDesc2<>"" then
		SCATDesc2=replace(SCATDesc2,"'","''")
		end if
		
		if LCATDesc2<>"" then
		LCATDesc2=replace(LCATDesc2,"'","''")
		end if
		
		if ppcategory2<>"" then
		ppcategory2=replace(ppcategory2,"'","''")
		end if
		
		if fileurl<>"" then
		fileurl=replace(fileurl,"'","''")
		end if			

		if localLG<>"" then
		localLG=replace(localLG,"'","''")
		end if

		if RemoteLG<>"" then
		RemoteLG=replace(RemoteLG,"'","''")
		end if		

		if LFN1<>"" then
		LFN1=replace(LFN1,"'","''")
		end if		

		if LFN2<>"" then
		LFN2=replace(LFN2,"'","''")
		end if
		
		if LFN3<>"" then
		LFN3=replace(LFN3,"'","''")
		end if
		
		if LFN4<>"" then
		LFN4=replace(LFN4,"'","''")
		end if
		
		if LFN5<>"" then
		LFN5=replace(LFN5,"'","''")
		end if

		if AddCopy<>"" then
		AddCopy=replace(AddCopy,"'","''")
		end if
		
		poprice=replace(poprice,scCurSign,"")
		plprice=replace(plprice,scCurSign,"")
		pwprice=replace(pwprice,scCurSign,"")
%>