<%
orderid=-1
shipid=-1
sendmailid=-1
shipdateid=-1
shipmethodid=-1
trackingid=-1

	ship_order=0
	ship_ship=0
	ship_shipdate=""
	ship_sendmail=0
	ship_shipmethod=""
	ship_tracking=""
		
TempProducts=""
ErrorsReport=""
%>

<!--#include file="ship-checkfields.asp"-->

<%
For i=1 to request("validfields")
	
	BLine=0
	Select Case request("T" & i)
	Case "Order ID": orderid=request("P" & i)
	BLine=1
	Case "Ship": shipid=request("P" & i)
	BLine=2
	Case "Send Mail": sendmailid=request("P" & i)
	BLine=3
	Case "Ship Date": shipdateid=request("P" & i)
	BLine=4
	Case "Method": shipmethodid=request("P" & i)	
	BLine=5
	Case "Tracking Number": trackingid=request("P" & i)
	BLine=6
	End Select
	if BLine>0 then
	TempStr=request("F" & i) & "*****"
	if instr(ALines(BLine-1),TempStr)=0 then
	ALines(BLine-1)=ALines(BLine-1) & TempStr
	end if
	BLine=0
	end if
Next

	SavedFile = "importlogs/ship-save.txt"
	findit = Server.MapPath(Savedfile)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	Err.number=0
	Set f = fso.OpenTextFile(findit, 2)
	For dd=lbound(ALines) to ubound(ALines)
	f.WriteLine ALines(dd)
	Next
	f.close
%>