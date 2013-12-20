<%PmAdmin=9%>
<% 'void order
err.clear
    '***********************************************************************************
	' This is where you would post info to the gateway
	' START
	'*************************************************************************************
	 	' Create an empty order
        Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
        order.setPartName("order")
		' Create an empty part
        Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                

        ' Build 'orderoptions'
        ' For a test, set result to GOOD, DECLINE, or DUPLICATE
		if lp_testmode ="YES" Then
        	res=op.put("result", "GOOD")			
		else
			res=op.put("result", "LIVE")			
		End if 
		
        res=op.put("ordertype","VOID")
		
		
		
        ' add 'orderoptions to order
        res=order.addPart("orderoptions", op)
		
		 res=op.clear()
		
        res=op.put("oid",pIdOrder)
		
        ' add 'merchantinfo to order
        res=order.addPart("transactiondetails", op)


        ' Build 'merchantinfo'
        res=op.clear()
		
        res=op.put("configfile",configfile)
        ' add 'merchantinfo to order
        res=order.addPart("merchantinfo", op)
		
		 ' Build 'billing'
          res=op.clear()
		  res=op.put("name", request.form("fname" &r)&" " & request.form("lname" &r))
          res=op.put("address1", request.form("address" & r))
          res=op.put("zip",request.form("zip"& R) )   
          
     
        ' add 'billing to order
          res=order.addPart("billing", op)

        ' Build 'creditcard'
        res=op.clear()
        res=op.put("cardnumber", cardnumber)
        res=op.put("cardexpmonth", expmonth)
        res=op.put("cardexpyear", expyear)
				
        ' add 'creditcard to order
        res=order.addPart("creditcard", op)
       
        ' Build 'payment'
        res=op.clear()
        res=op.put("chargetotal", money(curamount))
        ' add 'payment to order
        res=order.addPart("payment", op)
          
        if (fLog = True) and ( logLvl > 0 ) Then

            
		  resDesc = "ORDID: " & pIdOrder
		  res1 = res
		  	if PPD="1" then
				filename2="/"&scPcFolder&"/includes"
			else
				filename2="../includes"
			end if
			logFile = Server.MapPath(filename2) &"\" & logFile
			          
          'Next call return level of accepted logging in 'res1'
          'On error 'res1' contains negative number
          'You can check 'resDesc' to get error description
          'if any
          
          res = LPTxn.setDbgOpts(logFile,logLvl,resDesc,res1)
          
        End If
        
        ' get outgoing XML from 'order' object
       
        
         outXml = order.toXML()
   
        rsp = LPTxn.send(keyfile, host, port, outXml)
		
		
      
        Set LPTxn = Nothing
        Set order = Nothing
        Set op    = Nothing

		
        R_Time = ParseTag("r_time", rsp)
        R_Ref = ParseTag("r_ref", rsp)		
        R_Approved = ParseTag("r_approved", rsp)
		R_Code = ParseTag("r_code", rsp)
	    R_OrderNum = ParseTag("r_ordernum", rsp)
        R_Message = ParseTag("r_message", rsp)		
        R_Error = ParseTag("r_error", rsp)		
        R_TDate = ParseTag("r_tdate", rsp)
       
       
        Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")

'if success add to success/void
if R_Approved = "APPROVED" then
	' Create an empty order
        Set order = Server.CreateObject("LpiCom_6_0.LPOrderPart")
        order.setPartName("order")
		' Create an empty part
        Set op = Server.CreateObject("LpiCom_6_0.LPOrderPart")                

        ' Build 'orderoptions'
        ' For a test, set result to GOOD, DECLINE, or DUPLICATE
		if lp_testmode ="YES" Then
        	res=op.put("result", "GOOD")			
		else
			res=op.put("result", "LIVE")			
		End if 
		
        res=op.put("ordertype","SALE")
		
		
		
        ' add 'orderoptions to order
        res=order.addPart("orderoptions", op)
		
		 res=op.clear()
		
        res=op.put("oid",pIdOrder)
		
        ' add 'merchantinfo to order
        res=order.addPart("transactiondetails", op)


        ' Build 'merchantinfo'
        res=op.clear()
		
        res=op.put("configfile",configfile)
        ' add 'merchantinfo to order
        res=order.addPart("merchantinfo", op)

        ' Build 'creditcard'
        res=op.clear()
        res=op.put("cardnumber", cardnumber)
        res=op.put("cardexpmonth", expmonth)
        res=op.put("cardexpyear", expyear)
				
        ' add 'creditcard to order
        res=order.addPart("creditcard", op)
       
        ' Build 'payment'
        res=op.clear()
        res=op.put("chargetotal", money(curamount))
        ' add 'payment to order
        res=order.addPart("payment", op)

        if (fLog = True) and ( logLvl > 0 ) Then
          
          'Next call return level of accepted logging in 'res1'
          'On error 'res1' contains negative number
          'You can check 'resDesc' to get error description
          'if any          
          res = LPTxn.setDbgOpts(logFile,logLvl,resDesc,res1)
          
        End If
        
        ' get outgoing XML from 'order' object
         outXml = order.toXML()
         'response.write keyfile
		 'Response.end
        ' Call LPTxn
        rsp = LPTxn.send(keyfile, host, port, outXml)
		
		 'Store transaction data on Session and redirect
      
        Set LPTxn = Nothing
        Set order = Nothing
        Set op    = Nothing
		
		
        R_Time = ParseTag("r_time", rsp)
        R_Ref = ParseTag("r_ref", rsp)		
        R_Approved = ParseTag("r_approved", rsp)
		R_Code = ParseTag("r_code", rsp)
	    R_OrderNum = ParseTag("r_ordernum", rsp)
        R_Message = ParseTag("r_message", rsp)		
        R_Error = ParseTag("r_error", rsp)		
        R_TDate = ParseTag("r_tdate", rsp)
        Set LPTxn = Server.CreateObject("LpiCom_6_0.LinkPointTxn")
end if
err.clear %>
