<link rel="stylesheet" type="text/css" href="charts/jquery.jqplot.css" />
<!--[if IE]><script language="javascript" type="text/javascript" src="charts/excanvas.min.js"></script><![endif]-->
<script language="javascript" type="text/javascript" src="charts/jquery-1.4.2.min.js"></script>
<script language="javascript" type="text/javascript" src="charts/jquery.jqplot.min.js"></script>
<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.logAxisRenderer.min.js"></script>
<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.pointLabels.min.js"></script>


<%Dim pcvHave30Days,gridOptions
pcvHave30Days=0
gridOptions=", grid: {borderWidth:0.5, borderColor:'#CCC', shadow:false}"

Function pcf_GetOrderStatusTXT(porderstatus)
select case porderstatus
	case "0",""
		pcf_GetOrderStatusTXT="N/A"
	case "1"
	  pcf_GetOrderStatusTXT="Incomplete"
	case "2"
	  pcf_GetOrderStatusTXT="Pending" 
	case "3"
	  pcf_GetOrderStatusTXT="Processed" 
	case "4"
	  pcf_GetOrderStatusTXT="Shipped" 
	case "5"
	  pcf_GetOrderStatusTXT="Canceled" 
	case "6"
	  pcf_GetOrderStatusTXT="Return" 
	case "7"
	  pcf_GetOrderStatusTXT="Partially Shipped"
	case "8"
	  pcf_GetOrderStatusTXT="Shipping"
	case "9"
	  pcf_GetOrderStatusTXT="Partially Return"
	case "10"
	  pcf_GetOrderStatusTXT="Delivered" 
	case "11"
	  pcf_GetOrderStatusTXT="Will Not Deliver" 
	case "12"
	  pcf_GetOrderStatusTXT="Archived"
	end select
End Function

Function pcf_CustTypeTXT(pcusttype)
select case pcusttype
	case "0",""
		pcf_CustTypeTXT="Registered"
	case "1"
		pcf_CustTypeTXT="Guests"
	case "2"
		pcf_CustTypeTXT="Duplicated" 
	end select
End Function

Private Sub pcs_Gen30daysALLOrdersCharts(DivName,ShowLegend)
Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname,xname1
Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	call opendb()

	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		line4(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(date())
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(date())
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders FROM orders WHERE ((orders.orderStatus>=2 AND orders.orderStatus<5) OR (orders.orderStatus>=6)) AND orderdate>='" & past30 & "' AND orderdate<='" & Datenow & "' GROUP BY month(orderdate),day(orderdate) ORDER BY month(orderdate) ASC,day(orderdate) ASC;"
	else
		query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders FROM orders WHERE ((orders.orderStatus>=2 AND orders.orderStatus<5) OR (orders.orderStatus>=6)) AND orderdate>=#" & past30 & "# AND orderdate<=#" & Datenow & "# GROUP BY month(orderdate),day(orderdate) ORDER BY month(orderdate) ASC,day(orderdate) ASC;"
	end if
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcvHave30Days=1
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
		xname1=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & Year(date())
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & Year(date())
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
				xname1=xname1 & ","
			end if
			tmpline1=tmpline1 & line3(i)
			xname1=xname1 & "'" & replace(line5(i),"/","\%2F") & "'"
			if ((i+1)=1) OR ((i+1)=30) OR ((i+1) mod 5 = 0) then
				xname=xname & "'" & line5(i) & "'"
			else
				xname=xname & "' '"
			end if
		Next

		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		xname = [<%=xname1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'Number of Orders',
			seriesDefaults:
				{
					renderer:$.jqplot.BarRenderer,
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'Number of Orders',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			,
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=xname%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="resultsAdvancedAll.asp?fromdate=" + xname[pointIndex] + "&todate=" + xname[pointIndex] + "&otype=0&PayType=&B1=Search+Orders"
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else
	pcvHave30Days=0%>
	<div>A quick summary for last 30 days cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_Gen30daysCharts(DivName,DivName1,ShowLegend,NumCharts)
Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname,chartTitle,xname1
Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	call opendb()

	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		line4(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(date())
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(date())
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders, Sum(Total-rmaCredit) AS TotalAmounts, Sum(Total) AS TotalLessRMA FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orderdate>='" & past30 & "' AND orderdate<='" & Datenow & "' GROUP BY month(orderdate),day(orderdate) ORDER BY month(orderdate) ASC,day(orderdate);"
	else
		query="SELECT Day(OrderDate) As TheDay,Month(OrderDate) As TheMonth,Count(*) As TotalOrders, Sum(Total-rmaCredit) AS TotalAmounts, Sum(Total) AS TotalLessRMA  FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orderdate>=#" & past30 & "# AND orderdate<=#" & Datenow & "# GROUP BY month(orderdate),day(orderdate) ORDER BY month(orderdate) ASC,day(orderdate);"
	end if
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		tmpline2=""
		xname=""
		xname1=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				if isNULL(tmpArr(3,i)) then
					tmpArr(3,i) = tmpArr(4,i)
				end if
				line4(j)=Round(tmpArr(3,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & Year(date())
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & Year(date())
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				xname=xname & ","
				xname1=xname1 & ","
			end if
			tmpline1=tmpline1 & line3(i)
			tmpline2=tmpline2 & line4(i)
			xname1=xname1 & "'" & replace(line5(i),"/","\%2F") & "'"
			if ((i+1)=1) OR ((i+1)=30) OR ((i+1) mod 5 = 0) then
			xname=xname & "'" & line5(i) & "'"
			else
			xname=xname & "' '"
			end if
		Next

		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<%if NumCharts="1" OR NumCharts="0" then%>
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'Number of Orders',
			series:[
				{
					renderer:$.jqplot.BarRenderer, 
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'Number of Orders',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			],
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=xname%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});	});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
		<%end if%>
		<%if NumCharts="2" OR NumCharts="0" then
		chartTitle="Daily Sales Amount"%>
		<script>$(document).ready(function(){
		line2 = [<%=tmpline2%>];
		datename=[<%=xname1%>];
		function myClick(ev, gridpos, datapos, neighbor, plot) {
        if ((neighbor != null) && (plot.title.text=='<%=chartTitle%>')) {
			var tmpURL="salesReport.asp?FromDate=" + datename[neighbor.pointIndex] + "&ToDate=" + datename[neighbor.pointIndex] + "&basedon=1&customerType=&CountryCode=&submit=Search"
            window.open(tmpURL,"_blank");
            }
        }
		
		function myMove(ev, gridpos, datapos, neighbor, plot) {
        if ((neighbor != null) && (plot.title.text=='<%=chartTitle%>')) {
            document.getElementById("<%=DivName1%>").style.cursor='pointer';
        }
		else
		{
			document.getElementById("<%=DivName1%>").style.cursor='default';
		}
        }
		
		$.jqplot.eventListenerHooks.push(['jqplotMouseMove', myMove]);
		
		$.jqplot.eventListenerHooks.push(['jqplotClick', myClick]);
		
		plot2 = $.jqplot('<%=DivName1%>', [line2], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'<%=chartTitle%>',
			series:[
				{
					label: "Sales Amount",
					color:'#FF9933'
				}
				
			],
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=xname%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'<%=scCurSign%>%.2f'}}
			},
			highlighter: {show:true,tooltipAxes:'y',formatString:'%s'},
    		cursor: {show:false}
			<%=gridOptions%>
		});
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName1%>").style.clear='both';
			document.getElementById("<%=DivName1%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName1%>").style.float='right';
		</script>
		<%end if%>
		<%end if%>
	<%else%>
	<%if NumCharts="1" OR NumCharts="0" then%>
	<div>A quick summary for last 30 days cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
	<%end if%>
	<script>
		<%if NumCharts="1" OR NumCharts="0" then%>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
		<%end if%>
		<%if NumCharts="2" OR NumCharts="0" then%>
		document.getElementById("<%=DivName1%>").style.height='0px';
		document.getElementById("<%=DivName1%>").style.display='none';
		<%end if%>
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_MonthlySalesChart(DivName,TheYear,FullYear,ShowLegend)
	Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname
	Dim line1(12),line2(12),tmpDate

	call opendb()
	
	For i=0 to 11
		line1(i)=MonthName(i+1, True)
		line2(i)=0
	Next
	
	yearnow=TheYear
		
	query="SELECT Month(OrderDate) As TheMonth,Sum(Total-rmaCredit) AS TotalAmounts, Sum(Total) AS TotalLessRMA, Sum(rmaCredit) AS TotalRMA FROM orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<=9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND year(orderdate)=" & yearnow & " GROUP BY month(orderdate) ORDER BY month(orderdate) ASC;"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
		For i=0 to intCount
			'Override TotalAmounts due to possible NULLS
			if NOT isNumeric(tmpArr(3,i)) then
				tmpArr(3,i)=0
			end if
			tmpArr(1,i) = tmpArr(2,i)-tmpArr(3,i)

			pcv_YearTotal=pcv_YearTotal+Clng(tmpArr(1,i))

			For j=0 to 11
			if (Cint(tmpArr(0,i))=Cint(j+1)) then
				line2(j)=Clng(tmpArr(1,i))
				exit for
			end if
			Next
		Next

		For i=11 to 0 step -1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
			end if
			tmpline1=tmpline1 & "[" & line2(i) & "," & Cint(12-i) & "]"
			if FullYear=0 then
				if (Cint(i+1)>Cint(Month(Date()))) then
					xname=xname & "' '"
				else
					xname=xname & "'" & line1(i) & "'"
				end if
			else
				xname=xname & "'" & line1(i) & "'"
			end if
		Next

		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:false, location:'se', xoffset:15, yoffset:220},
			<%end if%>
			title:'',
			series:[
				{
					renderer:$.jqplot.BarRenderer, 
					rendererOptions:{barDirection:'horizontal'},
					label:'Monthly Sales - <%=TheYear%>',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			],
			axes:{
				yaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=xname%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				xaxis:{min:0,autoscale:true, tickOptions:{formatString:'<%=scCurSign%>%.0f'}}
			}
			<%=gridOptions%>
		});});	
		</script>
	<%else%>
	<div class="pcCPmessageInfo">A sales report for the current year cannot be created as no orders have yet been processed. Please note that pending, returned, and cancelled orders that might have been placed in the current year are not included in sales reports.</div>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
	

End Sub


Private Sub pcs_Top10Prds30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
Dim Datenow,past30

	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if

	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if uCase(scDB)="SQL" then
		query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.Quantity) DESC;"
	else
		query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.Quantity) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >=#" & past30 & "# AND orders.orderDate <=#" & Datenow & "# GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.Quantity) DESC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
			end if
			query="SELECT description FROM Products WHERE idproduct=" & rs("IDProduct") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrProductCompact = rsQ("description")
			else
				pcStrProductCompact = "N/A"
			end if
			set rsQ=nothing
			if len(pcStrProductCompact)>25 then
			 pcStrProductCompact = left(pcStrProductCompact,22) & "..."
			end if
			pcStrProductCompact=replace(pcStrProductCompact,"'","\'")
			tmpline1=tmpline1 & Clng(rs("PrdSales"))
			tmpline2=tmpline2 & "'" & pcStrProductCompact & "'"
			tmpline3=tmpline3 & rs("IDProduct")
			rs.MoveNext
		loop
		set rs=nothing
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			tmpline3=tmpline3 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		prdArr1=[<%=tmpline3%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'Top 10 Selling Products (Units Sold)',
			seriesDefaults:
				{
					renderer:$.jqplot.BarRenderer,
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'Top 10 Selling Products (Units Sold)',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			,
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=tmpline2%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=" + prdArr1[pointIndex] + "&submit=Search"
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_Top10PrdsAmount30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
Dim Datenow,past30

	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if uCase(scDB)="SQL" then
		query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity-ProductsOrdered.QDiscounts-ProductsOrdered.ItemsDiscounts) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity) DESC;"
	else
		query="SELECT TOP 10 ProductsOrdered.IDProduct, SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity-ProductsOrdered.QDiscounts-ProductsOrdered.ItemsDiscounts) As PrdSales FROM ProductsOrdered, Orders where Orders.IDOrder=ProductsOrdered.IDOrder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >=#" & past30 & "# AND orders.orderDate <=#" & Datenow & "# GROUP BY ProductsOrdered.IDProduct ORDER BY SUM(ProductsOrdered.UnitPrice*ProductsOrdered.Quantity) DESC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
			end if
			query="SELECT description FROM Products WHERE idproduct=" & rs("IDProduct") & ";"
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrProductCompact = rsQ("description")
			else
				pcStrProductCompact = "N/A"
			end if
			set rsQ=nothing
			if len(pcStrProductCompact)>25 then
			 pcStrProductCompact = left(pcStrProductCompact,22) & "..."
			end if
			pcStrProductCompact=replace(pcStrProductCompact,"'","\'")
			tmpline1=tmpline1 & Round(rs("PrdSales"),2)
			tmpline2=tmpline2 & "'" & pcStrProductCompact & "'"
			tmpline3=tmpline3 & rs("IDProduct")
			rs.MoveNext
		loop
		set rs=nothing 
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			tmpline3=tmpline3 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		prdArr2 = [<%=tmpline3%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'Top 10 Selling Products (Amount Sold)',
			seriesDefaults:
				{
					renderer:$.jqplot.BarRenderer,
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'Top 10 Selling Products (Amount Sold)',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			,
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=tmpline2%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=" + prdArr2[pointIndex] + "&submit=Search"
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_Top10Custs30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount,rsQ
Dim Datenow,past30

	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if UCase(scDB)="SQL" then
		query="SELECT TOP 10 idcustomer, sum(total) As AmountTotal, count(*) As NumOrders FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY idcustomer ORDER BY sum(total) DESC,count(*) DESC;"
	else
		query="SELECT TOP 10 idcustomer, sum(total) As AmountTotal, count(*) As NumOrders FROM Orders WHERE ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >=#" & past30 & "# AND orders.orderDate <=#" & Datenow & "# GROUP BY idcustomer ORDER BY sum(total) DESC,count(*) DESC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		icount=0
		tmpline1=""
		tmpline2=""
		tmpline3=""
		do while (not rs.eof) AND (icount<10)
			icount=icount+1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
				tmpline3=tmpline3 & ","
			end if
			query="SELECT name, lastname FROM Customers WHERE idcustomer=" & rs("idcustomer")
			set rsQ=connTemp.execute(query)
			if not rsQ.eof then
				pcStrNameCompact = rsQ("name") & " " & rsQ("lastname")
			else
				pcStrNameCompact = ""
			end if
			set rsQ=nothing
			if len(pcStrNameCompact)>25 then
			 pcStrNameCompact = left(pcStrNameCompact,22) & "..."
			end if
			pcStrNameCompact=replace(pcStrNameCompact,"'","\'")
			tmpline1=tmpline1 & Clng(rs("AmountTotal"))
			tmpline2=tmpline2 & "'" & pcStrNameCompact & "'"
			tmpline3=tmpline3 & rs("idcustomer")
			rs.MoveNext
		loop
		set rs=nothing
		if icount<10 then
			For k=icount+1 to 10
			tmpline1=tmpline1 & ",0"
			tmpline3=tmpline3 & ",0"
			if k<10 then
				tmpline2=tmpline2 & ",' '"
			else
				tmpline2=tmpline2 & ",'.'"
			end if
			Next
		end if	
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		custArr1 = [<%=tmpline3%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'Top 10 Customers (Amount Sold)',
			seriesDefaults:
				{
					renderer:$.jqplot.BarRenderer,
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'Top 10 Customers (Amount Sold)',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			,
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=tmpline2%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
				
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="viewCustOrders.asp?idcustomer=" + custArr1[pointIndex]
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_OrdStatus30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount
Dim Datenow,past30

	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT OrderStatus,Count(*) As TotalOrders FROM Orders WHERE (Orders.OrderStatus>=2) AND Orderdate>='" & past30 & "' AND Orderdate<='" & Datenow & "' GROUP BY OrderStatus ORDER BY Count(*) DESC;"
	else
		query="SELECT OrderStatus,Count(*) As TotalOrders FROM Orders WHERE (Orders.OrderStatus>=2) AND Orderdate>=#" & past30 & "# AND Orderdate<=#" & Datenow & "# GROUP BY OrderStatus ORDER BY Count(*) DESC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		tmpline2=""
		For i=0 to intCount
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				tmpline2=tmpline2 & ","
			end if
			tmpline1=tmpline1 & "['" & pcf_GetOrderStatusTXT(tmpArr(0,i)) & ": " & Clng(tmpArr(1,i)) & "'," & Clng(tmpArr(1,i)) & "]"
			tmpline2=tmpline2 & tmpArr(0,i)
		Next
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		OrdStatusArr = [<%=tmpline2%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
    	title: 'Order Status',
    	seriesDefaults:{renderer:$.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,dataLabels: 'percent', dataLabelFormatString: '%.1f%%', sliceMargin:0}},
    	legend:{show:true}
		<%=gridOptions%>
		});
		
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="resultsAdvancedAll.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&otype=" + OrdStatusArr[pointIndex] + "&PayType=&B1=Search+Orders"
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub

Private Sub pcs_NewCusts30Days(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,pcArr,icount
Dim Datenow,past30

IF (pcvHave30Days=1) AND ((scGuestCheckoutOpt=0) OR (scGuestCheckoutOpt=1)) THEN
	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT pcCust_Guest,Count(*) FROM Customers WHERE pcCust_DateCreated>='" & past30 & "' AND pcCust_DateCreated<='" & Datenow & "' GROUP BY pcCust_Guest ORDER BY Count(*) DESC;"
	else
		query="SELECT pcCust_Guest,Count(*) FROM Customers WHERE pcCust_DateCreated>=#" & past30 & "# AND pcCust_DateCreated<=#" & Datenow & "# GROUP BY pcCust_Guest ORDER BY Count(*) DESC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	TotalCustomer=0
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		For i=0 to intCount
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
			end if
			TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
			tmpline1=tmpline1 & "['" & pcf_CustTypeTXT(tmpArr(0,i)) & ": " & Clng(tmpArr(1,i)) & "'," & Clng(tmpArr(1,i)) & "]"
		Next
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
    	title: 'New Customer Registrations: <%=TotalCustomer%>',
    	seriesDefaults:{renderer:$.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,dataLabels: 'percent', dataLabelFormatString: '%.1f%%',sliceMargin:0}},
    	legend:{show:true}
		<%=gridOptions%>
		});});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
ELSE
	if pcvHave30Days=1 then
		call pcs_NewCustsOnly30Days(DivName)
	else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if%>
<%END IF
End Sub

Private Sub pcs_NewCustsOnly30Days(DivName)
Dim past30,Datenow,rs,query,tmpArr,i,j,intCount,tmpline1,tmpline2,xname
Dim line1(30),line2(30),line3(30),line4(30),line5(30),tmpDate

	call opendb()
	
	Datenow=Date()
	past30=Date()-29
	
	For i=29 to 0 step -1
		tmpDate=Date()-i
		line1(29-i)=Day(tmpDate)
		line2(29-i)=Month(tmpDate)
		line3(29-i)=0
		if scDateFrmt="DD/MM/YY" then
			line5(29-i)=Day(tmpDate) & "/" & Month(tmpDate) & "/" & Year(date())
		else
			line5(29-i)=Month(tmpDate) & "/" & Day(tmpDate) & "/" & Year(date())
		end if
	Next
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT Day(pcCust_DateCreated) As TheDay,Month(pcCust_DateCreated) As TheMonth,Count(*) FROM Customers WHERE pcCust_DateCreated>='" & past30 & "' AND pcCust_DateCreated<='" & Datenow & "' GROUP BY month(pcCust_DateCreated),day(pcCust_DateCreated) ORDER BY month(pcCust_DateCreated) ASC,day(pcCust_DateCreated) ASC;"
	else
		query="SELECT Day(pcCust_DateCreated) As TheDay,Month(pcCust_DateCreated) As TheMonth,Count(*) FROM Customers WHERE pcCust_DateCreated>=#" & past30 & "# AND pcCust_DateCreated<=#" & Datenow & "# GROUP BY month(pcCust_DateCreated),day(pcCust_DateCreated) ORDER BY month(pcCust_DateCreated) ASC,day(pcCust_DateCreated) ASC;"
	end if
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	TotalCustomer=0
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		tmpline1=""
		xname=""
		tmpline1=""
		xname=""
		For i=0 to intCount
			For j=0 to 29
			if (Cint(tmpArr(0,i))=Cint(line1(j))) AND (Cint(tmpArr(1,i))=Cint(line2(j))) then
				line3(j)=Round(tmpArr(2,i),2)
				TotalCustomer=TotalCustomer+Round(tmpArr(2,i),2)
				if scDateFrmt="DD/MM/YY" then
					line5(j)=tmpArr(0,i) & "/" & tmpArr(1,i) & "/" & Year(date())
				else
					line5(j)=tmpArr(1,i) & "/" & tmpArr(0,i) & "/" & Year(date())
				end if
				exit for
			end if
			Next
		Next
		
		For i=0 to 29
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
				xname=xname & ","
			end if
			tmpline1=tmpline1 & line3(i)
			if ((i+1)=1) OR ((i+1)=30) OR ((i+1) mod 5 = 0) then
			xname=xname & "'" & line5(i) & "'"
			else
			xname=xname & "' '"
			end if
		Next
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.highlighter.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.cursor.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:55},
			<%end if%>
			title:'New Customers: <%=TotalCustomer%>',
			series:[
				{
					renderer:$.jqplot.BarRenderer, 
					rendererOptions: {
                	    barWidth:8   
                	},
					label:'New Customers',
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			],
			axes:{
				xaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: [<%=xname%>],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				yaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});	});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.height='0px';
			document.getElementById("<%=DivName%>").style.display='none';
		</script>
	<%end if
	set rs=nothing

End Sub

Function pcf_PricingCatName(IDCat)
Dim query,rs

	query="SELECT pcCC_Name FROM pcCustomerCategories WHERE idCustomerCategory=" & IDCat & ";"
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		pcf_PricingCatName=replace(rs("pcCC_Name"),"'","\'")
	else
		pcf_PricingCatName=""
	end if
	
	set rs=nothing

End Function

Private Sub pcs_PricingCatsChart(DivName)
Dim query,rs,tmpArr,intCount,i,j,tmpline1,tmpline2,xname,xvalue,pcArr,icount
Dim Datenow,past30,line1(100),line2(100)

	call opendb()
	
	query="SELECT idCustomerCategory,Count(*) FROM Customers WHERE idCustomerCategory>0 GROUP BY idCustomerCategory ORDER BY Count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	TotalCustomer=0
	tmpline1=""
	iCount=0
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		
		For i=0 to intCount
			line1(icount)=Clng(tmpArr(1,i))
			line2(icount)=pcf_PricingCatName(tmpArr(0,i))
			icount=icount+1
			TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
		Next
	end if
	set rs=nothing
	
	query="SELECT customerType,Count(*) FROM Customers WHERE idCustomerCategory=0 GROUP BY customerType ORDER BY Count(*) DESC;"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		tmpArr=rs.getRows()
		set rs=nothing
		intCount=ubound(tmpArr,2)
		For i=0 to intCount
			TotalCustomer=TotalCustomer+Clng(tmpArr(1,i))
			line1(icount)=Clng(tmpArr(1,i))
			if tmpArr(0,i)="0" then
				xname="Retail"
			else
				xname="Wholesale"
			end if
			line2(icount)=xname
			icount=icount+1
		Next
	end if
	set rs=nothing
	
	if TotalCustomer>0 then
		For i=0 to icount-1
			For j=i+1 to icount-1
				if Clng(line1(i))<Clng(line1(j)) then
					xname=line1(i)
					xvalue=line2(i)
					line1(i)=line1(j)
					line2(i)=line2(j)
					line1(j)=xname
					line2(j)=xvalue
				end if
			Next
		Next
		
		tmpline1=""
		
		For i=0 to icount-1
			if tmpline1<>"" then
				tmpline1=tmpline1 & ","
			end if
			tmpline1=tmpline1 & "['" & line2(i) & ": " & Clng(line1(i)) & "'," & Clng(line1(i)) & "]"
		Next
	
		%>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.pieRenderer.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
    	title: 'Total Customers: <%=TotalCustomer%>',
    	seriesDefaults:{renderer:$.jqplot.PieRenderer, rendererOptions:{showDataLabels: true,dataLabels: 'percent', dataLabelFormatString: '%.1f%%',sliceMargin:0}},
    	legend:{show:true}
		<%=gridOptions%>
		});});
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
	</script>
	<%end if

End Sub

Private Sub pcs_GenPrd30daysCharts(DivName,DivName1,tmpIDProduct,ShowLegend)
Dim past30,Datenow,rs,query,tmpline1,tmpline2,tmpline3
Dim TotalQty,TotalAmount,CurrentQty

	call opendb()

	Datenow=Date()
	past30=Date()-29
	
	if scDateFrmt="DD/MM/YY" then
		tmpDateNow=Day(Datenow) & "/" & Month(Datenow) & "/" & Year(Datenow)
		tmppast30=Day(past30) & "/" & Month(past30) & "/" & Year(past30)
	else
		tmpDateNow=Month(Datenow) & "/" & Day(Datenow) & "/" & Year(Datenow)
		tmppast30=Month(past30) & "/" & Day(past30) & "/" & Year(past30)
	end if
	
	
	if SQL_Format="1" then
		Datenow=(day(Datenow)&"/"&month(Datenow)&"/"&year(Datenow))
	else
		Datenow=(month(Datenow)&"/"&day(Datenow)&"/"&year(Datenow))
	end if
	
	if SQL_Format="1" then
		past30=(day(past30)&"/"&month(past30)&"/"&year(past30))
	else
		past30=(month(past30)&"/"&day(past30)&"/"&year(past30))
	end if
	
	if Ucase(scDB)="SQL" then
		query="SELECT Sum(ProductsOrdered.quantity) AS TotalQty,Sum(ProductsOrdered.quantity*ProductsOrdered.unitPrice) AS TotalAmount FROM Orders,ProductsOrdered WHERE ProductsOrdered.IDProduct=" & tmpIDProduct & " and orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >='" & past30 & "' AND orders.orderDate <='" & Datenow & "' GROUP BY ProductsOrdered.IDProduct;"
	else
		query="SELECT Sum(ProductsOrdered.quantity) AS TotalQty,Sum(ProductsOrdered.quantity*ProductsOrdered.unitPrice) AS TotalAmount FROM Orders,ProductsOrdered WHERE ProductsOrdered.IDProduct=" & tmpIDProduct & " and orders.idorder=ProductsOrdered.idorder and ((orders.orderStatus>2 AND orders.orderStatus<5) OR (orders.orderStatus>6 AND orders.orderStatus<9) OR (orders.orderStatus=10 OR orders.orderStatus=12)) AND orders.orderDate >=#" & past30 & "# AND orders.orderDate <=#" & Datenow & "# GROUP BY ProductsOrdered.IDProduct;" 
	end if
	set rs=connTemp.execute(query)
	
	if not rs.eof then
		TotalQty=rs("TotalQty")
		TotalAmount=rs("TotalAmount")
		set rs=nothing
		query="SELECT stock FROM Products WHERE idProduct=" & tmpIDProduct & ";"
		set rs=connTemp.execute(query)
		CurrentQty=0
		if not rs.eof then
			CurrentQty=rs("stock")
			if Clng(CurrentQty)<0 then
				CurrentQty=0
			end if
		end if
		set rs=nothing
		tmpline1="[" &  TotalAmount & ",1]"
		tmpline2="[" & TotalQty & ",2]"
		tmpline3="[" & CurrentQty & ",1]"
		%>
		
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.barRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.categoryAxisRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasTextRenderer.min.js"></script>
		<script language="javascript" type="text/javascript" src="charts/plugins/jqplot.canvasAxisTickRenderer.min.js"></script>
		
		<script>$(document).ready(function(){
		line1 = [<%=tmpline1%>];
		plot2 = $.jqplot('<%=DivName%>', [line1], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:15, yoffset:220},
			<%end if%>
			title:'Quick Summary: sales in last 30 days',
			series:[
				{
					renderer:$.jqplot.BarRenderer, 
					rendererOptions:{barDirection:'horizontal', barWidth:5},
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				}
			],
			axes:{
				yaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: ['Amount Ordered'],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				xaxis:{min:0,autoscale:true, tickOptions:{formatString:'<%=scCurSign%>%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$('#<%=DivName%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=<%=tmpIDProduct%>&submit=Search"
				window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				document.getElementById("<%=DivName%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName%>").style.cursor='default';
            }
        );
		
		});	
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName%>").style.clear='both';
			document.getElementById("<%=DivName%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName%>").style.float='right';
		</script>
		<%end if%>
		<script>$(document).ready(function(){
		line2 = [<%=tmpline2%>];
		line3 = [<%=tmpline3%>];
		plot2 = $.jqplot('<%=DivName1%>', [line2,line3], {
			<%if ShowLegend=1 then%>
			legend:{show:true, location:'ne', xoffset:15, yoffset:220},
			<%end if%>
			title:'',
			seriesDefaults:{
				
					renderer:$.jqplot.BarRenderer, 
					rendererOptions:{barDirection:'horizontal', barWidth:5},
					pointLabels:{show:true, stackedValue: true, hideZeros:true}
				
			},
			axes:{
				yaxis:{
					renderer:$.jqplot.CategoryAxisRenderer,
					ticks: ['Current Inventory','Qty. Ordered'],
					rendererOptions:{tickRenderer:$.jqplot.CanvasAxisTickRenderer},
            		tickOptions:{
					fontSize:'10px', 
                	fontFamily:'Arial', 
                	angle:-30
           			}
				},
				xaxis:{min:0,autoscale:true, tickOptions:{formatString:'%.0f'}}
			}
			<%=gridOptions%>
		});
		
		$('#<%=DivName1%>').bind('jqplotDataClick', 
            function (ev, seriesIndex, pointIndex, data) {
				var tmpURL="PrdsalesReport.asp?FromDate=<%=replace(tmppast30,"/","\%2F")%>&ToDate=<%=replace(tmpDateNow,"/","\%2F")%>&basedon=1&IDProduct=<%=tmpIDProduct%>&submit=Search"
				if (seriesIndex==0) window.open(tmpURL,"_blank");
            }
        );
		
		$('#<%=DivName1%>').bind('jqplotDataHighlight', 
            function (ev, seriesIndex, pointIndex, data) {
				if (seriesIndex==0)	document.getElementById("<%=DivName1%>").style.cursor='pointer';
            }
        );
		
		$('#<%=DivName1%>').bind('jqplotDataUnhighlight', 
            function (ev) {
                document.getElementById("<%=DivName1%>").style.cursor='default';
            }
        );
		
		});	
		</script>
		<%ChartCount=ChartCount+1
		if (ChartCount mod 2)=1 then%>
		<script>
			document.getElementById("<%=DivName1%>").style.clear='both';
			document.getElementById("<%=DivName1%>").style.float='left';
		</script>
		<%else%>
		<script>
			document.getElementById("<%=DivName1%>").style.float='right';
		</script>
		<%end if%>
	<%else%>
	<script>
		document.getElementById("<%=DivName%>").style.height='0px';
		document.getElementById("<%=DivName%>").style.display='none';
		document.getElementById("<%=DivName1%>").style.height='0px';
		document.getElementById("<%=DivName1%>").style.display='none';
	</script>
	<%end if
	set rs=nothing
End Sub
%>
