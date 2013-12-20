<%@ Language=VBScript %>
<!--#include file="../includes/settings.asp"-->
<% Session.LCID = 1033 %>
<%
	Dim m_dtCurrentDate		'Currently selected date/time
	Dim m_lDayofFirst		'The day of the week that the first of the current month falls on
	Dim m_lDaysInMonth		'Number of days in the selected month
	Dim m_dtBegin			'Beginning date of the selected month
	Dim m_dtEnd				'Ending date of the selected month

	Dim m_lYear				'Currently selected Year
	Dim m_lMonth			'Currently selected Month
	Dim m_lDay				'Currently selected Day of the month
	
	Dim m_sInputName		'Name of the input field from the parent page
	Dim m_dtPassedInDate	
	
	m_sInputName = Request.QueryString("N")
	
	'Build the date/time from individual parts if there has been a post back.
	'Otherwise, just get the current date/time.
	If Request.QueryString("A") <> "" Then
		m_lYear = Request.Form("fldYear")
		m_lMonth = Request.Form("fldMonth")
		m_lDay = Request.Form("fldDay")
	
		
		'Fix the day of the month if we switch from a month that has less days in the month 
		'than the previously selected month and the day selected is not on the newly selected
		'month (ie - going from March 31st and then selecting February which does not have a 31st.)
		m_dtBegin = m_lMonth & "/1/" & m_lYear
		m_dtEnd = DateAdd("m", 1, m_dtBegin)
		m_lDaysInMonth = DateDiff("d", m_dtBegin, m_dtEnd)
		If CLng(m_lDay) > CLng(m_lDaysInMonth) Then m_lDay = m_lDaysInMonth
		
		'Build the Date
		m_dtCurrentDate = m_lMonth & "/" & m_lDay & "/" & m_lYear
		m_dtCurrentDate = m_dtCurrentDate
		
		'If the date is not valid after all this then use the current date.
		If IsDate(m_dtCurrentDate) Then
			m_dtCurrentDate = CDate(m_dtCurrentDate)
		Else
			m_dtCurrentDate = Now()
		End If
		
	Else
		m_dtPassedInDate = Request.QueryString("DT")
		If CStr(m_dtPassedInDate) <> "" Then
			If IsDate(m_dtPassedInDate) Then
				If scDateFrmt="DD/MM/YY" then 
					m_dtCurrentDate = day(m_dtPassedInDate) & "/" & month(m_dtPassedInDate) & "/" & year(m_dtPassedInDate)
				Else
					m_dtCurrentDate = month(m_dtPassedInDate) & "/" & day(m_dtPassedInDate) & "/" & year(m_dtPassedInDate)
				End If
			Else
				m_dtCurrentDate = Now()
			End If
		Else
			m_dtCurrentDate = Now()	
		End If
	End If
	
	'Break out certain parts of the currently selected date/time.
	m_lYear = DatePart("yyyy", m_dtCurrentDate)
	m_lMonth = DatePart("m", m_dtCurrentDate)
	m_lDay = DatePart("d", m_dtCurrentDate)
	'Figure out if we need an AM or PM
	
	
	m_dtBegin = CDate(DatePart("m", m_dtCurrentDate) & "/1/" & DatePart("yyyy", m_dtCurrentDate))
	m_dtEnd = DateAdd("m", 1, m_dtBegin)
	m_lDayofFirst = DatePart("w", m_dtBegin)
	m_lDaysInMonth = DateDiff("d", m_dtBegin, m_dtEnd)
%>
<HTML>
<HEAD>
<TITLE>Choose a Date</TITLE>
<LINK rel="stylesheet" type="text/css" href="Calendar.css">
</HEAD>
<BODY bgcolor="#D4D0C8">
<FORM method=post action="Calendar.asp?A=1&N=<%=m_sInputName%>" id=Form1 name=Form1>
<table class=overall cellpadding=0 cellspacing=10>
<tr>
	<td>
		<%DisplayCalendar%>
	</td>
	<td valign=top>
		<table>
		<tr>
			      <td class="Title"> Month:<BR>
			</td>
			<td class="Title">
				Year:<BR>		
			</td>
		</tr>
		<tr>
			<td>
				<SELECT id=fldMonth name=fldMonth onchange="javascript:document.Form1.submit();">
				<%DisplayMonths%>
				</SELECT>
			</td>
			<td>
				<INPUT type="text" id=fldYear name=fldYear value="<%=m_lYear%>" size=4 maxlength=4  onblur="javascript:document.Form1.submit();"><BR>
			</td>
		</tr>
		</table>
		<BR>
		<INPUT type="button" value="OK" id=cmdOK name=cmdOK class=Button onclick="javascript:SetDateOnParent();">&nbsp;
		<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel class=Button onclick="javascript:window.close();">
	</td>
</tr>
</table>
</FORM>
</BODY>
</HTML>
<%
Sub DisplayMonths
	Dim arrMonths		'An array of months starting with January
	Dim i				'counter
	
	arrMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	
	For i = 0 To UBound(arrMonths)
		If CLng(i) = (CLng(m_lMonth) - 1) Then
			Response.Write "<OPTION value=" & i + 1 & " SELECTED>" & arrMonths(i) & "</OPTION>"
		Else
			Response.Write "<OPTION value=" & i + 1 & ">" & arrMonths(i) & "</OPTION>"
		End If
	Next
End Sub


Sub DisplayCalendar
	Dim arrDays				'An array of days starting with Sunday
	Dim i					'counter
	Dim lColumnCount		'Column Count
	Dim lNumber				'For building the calendar
	Dim bFinished			'For building the calendar
	Dim bStart				'For building the calendar
	
	arrDays = Array("S", "M", "T", "W", "T", "F", "S")
	
	Response.Write "<table width=100 class='Calendar' cellpadding=4 cellspacing=0><tr class=Header>"
	For i = 0 to Ubound(arrDays)
		Response.write "<td>" & arrDays(i) & "</td>"
	Next
	
	lNumber = 1
	bFinished = False
	bStart = False
	Do
		Response.Write "<tr>"
		For lColumnCount = 1 to 7
			If CLng(lColumnCount) = CLng(m_lDayofFirst) Then
				bStart = True
			End If
			
			If CLng(m_lDay) = CLng(lNumber) AND CBool(bStart) Then
				Response.Write "<td class=SelectedDay>"
				Response.Write "<input name=fldDay type=hidden value=" & lNumber & ">"
			Else
				Response.Write "<td class=NormalDay>"
			End If
			
			If NOT CBool(bFinished) AND CBool(bStart) Then
			
'			Un-Comment the lines below to disable weekend selection
'			If lColumnCount = 1 or lColumnCount = 7 Then
'				Response.Write "<font color=#990000>" & lNumber & "</font>"
'			Else
				If CLng(m_lDay) <> CLng(lNumber) Then
					Response.Write "<A href='javascript:ChangeDay(" & lNumber & ","&m_lMonth&","&m_lYear&");'>"
				End If
				Response.Write lNumber
				If CLng(m_lDay) <> CLng(lNumber) Then
					Response.Write "</A>"
				End If
'			End If
				
				
				lNumber = lNumber + 1
				If CLng(lNumber) > CLng(m_lDaysInMonth) Then
					bFinished = True
				End If
			
			Else
				Response.Write "&nbsp;"
			End If
			Response.Write "</td>"
		Next
		Response.Write "</tr>"
	Loop Until CBool(bFinished)
	
	Response.Write "</tr></table>"
End Sub
%>

<script language="javascript">
	function ChangeDay(v_lDay,v_lMonth,v_lYear) {
		document.Form1.fldDay.value = v_lDay;
		document.Form1.fldMonth.value = v_lMonth;
		document.Form1.fldYear.value = v_lYear;
		document.Form1.submit();
	}
	
	function SetDateOnParent() {
		var sRetDate;
	
		<% If scDateFrmt="DD/MM/YY" then %>
			sRetDate = '<%=Day(m_dtCurrentDate) & "/" & Month(m_dtCurrentDate) & "/" & Year(m_dtCurrentDate)%>';
		<% Else  %>
			sRetDate = '<%=Month(m_dtCurrentDate) & "/" & Day(m_dtCurrentDate) & "/" & Year(m_dtCurrentDate)%>';
		<% End If %>
		
		
		window.opener.<%=m_sInputName%>.value = sRetDate;
		window.close();
	}
</script>