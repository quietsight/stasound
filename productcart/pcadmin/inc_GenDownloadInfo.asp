<%
'// Call License Generator for Standard & BTO Products
	Sub CreateDownloadInfo(pIDProduct,pQuantity)
		Dim query,rsQ,pSku,pLicense,pLocalLG,pRemoteLG,k,dd
			query="select sku,License,LocalLG,RemoteLG from Products,DProducts where products.idproduct=" & pIdproduct & " and DProducts.idproduct=Products.idproduct and products.downloadable=1"
			set rsQ=server.CreateObject("ADODB.RecordSet")
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pSku=rsQ("sku")
				pLicense=rsQ("License")
				pLocalLG=rsQ("LocalLG")
				pRemoteLG=rsQ("RemoteLG")
				set rsQ=nothing
				
				IF (pLicense<>"") and (pLicense="1") THEN
					if pLocalLG<>"" then
						SPath1=Request.ServerVariables("PATH_INFO")
						mycount1=0
						do while mycount1<1
							if mid(SPath1,len(SPath1),1)="/" then
								mycount1=mycount1+1
							end if
							if mycount1<1 then
								SPath1=mid(SPath1,1,len(SPath1)-1)
							end if
						loop
						SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1
						if Right(SPathInfo,1)="/" then
							pLocalLG=SPathInfo & "licenses/" & pLocalLG					
						else
							pLocalLG=SPathInfo & "/licenses/" & pLocalLG
						end if
						pLocalLG=replace(pLocalLG,"/pc/","/"&scAdminFolderName&"/")
						L_Action=pLocalLG
					else
						L_Action=pRemoteLG
					end if
					L_postdata=""
					L_postdata=L_postdata&"idorder=" & pIdOrder
					L_postdata=L_postdata&"&orderDate=" & pOrderDate
					L_postdata=L_postdata&"&ProcessDate=" & pProcessDate
					L_postdata=L_postdata&"&idcustomer=" & pIdCustomer
					L_postdata=L_postdata&"&idproduct=" & pIdproduct
					L_postdata=L_postdata&"&quantity=" & pQuantity
					L_postdata=L_postdata&"&sku=" & pSKU

					Set srvXmlHttp = server.createobject("Msxml2.serverXmlHttp"&scXML)
					srvXmlHttp.open "POST", L_Action, False
					srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					srvXmlHttp.send L_postdata
					result1 = srvXmlHttp.responseText
					AR=split(result1,"<br>")

					rIdOrder=AR(0)
					rIdProduct=AR(1)
					Lic1=split(AR(2),"***")
					Lic2=split(AR(3),"***")
					Lic3=split(AR(4),"***")
					Lic4=split(AR(5),"***")
					Lic5=split(AR(6),"***")
	
					For k=0 to Cint(pQuantity)-1
						if K<=ubound(Lic1) then
							PLic1=Lic1(k)
						else
							PLic1=""
						end if
						if K<=ubound(Lic2) then
							PLic2=Lic2(k)
						else
							PLic2=""
						end if
						if K<=ubound(Lic3) then
							PLic3=Lic3(k)
						else
							PLic3=""
						end if
						if K<=ubound(Lic4) then
							PLic4=Lic4(k)
						else
							PLic4=""
						end if
						if K<=ubound(Lic5) then
							PLic5=Lic5(k)
						else
							PLic5=""
						end if
						if ppStatus=0 then
							query="Insert into DPLicenses (IdOrder,IdProduct,Lic1,Lic2,Lic3,Lic4,Lic5) values (" & rIdOrder & "," & rIdProduct & ",'" & PLic1 & "','" & PLic2 & "','" & PLic3 & "','" & PLic4 & "','" & PLic5 & "')"   
							set rsQ=server.CreateObject("ADODB.RecordSet")
							set rsQ=connTemp.execute(query)
							set rsQ=nothing
						end if
					Next
				END IF
				
				DO
					Tn1=""
					Tn1=Tn1 & Year(Date()) & Month(Date()) & Day(Date()) & Hour(Time()) & Minute(Time())
					LenTn1=Len(Tn1)
						For dd=LenTn1+1 to 50
							Randomize
							myC=Fix(3*Rnd)
							Select Case myC
								Case 0: 
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+65)
								Case 1: 
									Randomize
									Tn1=Tn1 & Cstr(Fix(10*Rnd))
								Case 2: 
									Randomize
									Tn1=Tn1 & Chr(Fix(26*Rnd)+97)		
							End Select		
							Randomize		
						Next
	
						ReqExist=0
	
						query="select IDOrder from DPRequests where RequestSTR LIKE '" & Tn1 & "'" 
						set rsQ=server.CreateObject("ADODB.RecordSet")
						set rsQ=connTemp.execute(query)
	
						if not rsQ.eof then
							ReqExist=1
						end if
						set rsQ=nothing
				LOOP UNTIL ReqExist=0
	
				if ppStatus=0 then
					pTodaysDate=Date()
					if SQL_Format="1" then
						pTodaysDate=(day(pTodaysDate)&"/"&month(pTodaysDate)&"/"&year(pTodaysDate))
					else
						pTodaysDate=(month(pTodaysDate)&"/"&day(pTodaysDate)&"/"&year(pTodaysDate))
					end if
		
					'Insert Standard & BTO Products Download Requests into DPRequests Table
					if scDB="Access" then
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "',#" & pTodaysDate & "#)"   
					else
						query="Insert into DPRequests (IdOrder,IdProduct,IdCustomer,RequestSTR,StartDate) values (" & pIdOrder & "," & pIdProduct & "," & pIdCustomer & ",'" & Tn1 & "','" & pTodaysDate & "')"
					end if
					set rsQ=server.CreateObject("ADODB.RecordSet")
					set rsQ=connTemp.execute(query)
					set rsQ=nothing
				end if
			end if
			set rsQ=nothing

	End Sub
%>
