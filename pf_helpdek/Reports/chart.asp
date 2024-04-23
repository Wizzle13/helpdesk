<%
RowNum=1
Sub ShowChart(ByRef aValues, ByRef aLabels, ByRef strTitle, ByRef strXAxisLabel, ByRef strYAxisLabel)
	' Some user changable graph defining constants
	' All units are in screen pixels
	Const GRAPH_WIDTH  = 550  ' The width of the body of the graph
	Const GRAPH_HEIGHT = 250  ' The heigth of the body of the graph
	Const GRAPH_BORDER = 5    ' The size of the black border
	Const GRAPH_SPACER = 2    ' The size of the space between the bars

	' Debugging constant so I can eaasily switch on borders in case
	' the tables get messed up.  Should be left at zero unless you're
	' trying to figure out which table cells doing what.
	Const TABLE_BORDER = 0
	'Const TABLE_BORDER = 10

	' Declare our variables
	Dim I
	Dim iMaxValue
	Dim iBarWidth
	Dim iBarHeight

	' Get the maximum value in the data set
	iMaxValue = 0
	For I = 0 To UBound(aValues)
		If iMaxValue < aValues(I) Then iMaxValue = aValues(I)
	Next 'I
	'Response.Write iMaxValue ' Debugging line


	' Calculate the width of the bars
	' Take the overall width and divide by number of items and round down.
	' I then reduce it by the size of the spacer so the end result
	' should be GRAPH_WIDTH or less!
	iBarWidth = (GRAPH_WIDTH \ (UBound(aValues) + 1)) - GRAPH_SPACER
	'Response.Write iBarWidth ' Debugging line


	' Start drawing the graph
	%>
	<TABLE BORDER="<%= TABLE_BORDER %>" CELLSPACING="0" CELLPADDING="0">
		<TR>
			<TD COLSPAN="3" ALIGN="center"><H2><%= strTitle %></H2></TD>
		</TR>
		<TR>
			<TD VALIGN="center"><B><%= strYAxisLabel %></B></TD>
			<TD VALIGN="top">
				<TABLE BORDER="<%= TABLE_BORDER %>" CELLSPACING="0" CELLPADDING="0">
					<TR>
						<TD ROWSPAN="2"><IMG SRC="../images/spacer.gif" BORDER="0" WIDTH="1" HEIGHT="<%= GRAPH_HEIGHT %>"></TD>
						<TD VALIGN="top" ALIGN="right"><%= iMaxValue %>&nbsp;</TD>
					</TR>
					<TR>
						<TD VALIGN="bottom" ALIGN="right">0&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
			<TD>
				<TABLE BORDER="<%= TABLE_BORDER %>" CELLSPACING="0" CELLPADDING="0">
					<TR>
						<TD VALIGN="bottom"><IMG SRC="../images/spacer_black.gif" BORDER="0" WIDTH="<%= GRAPH_BORDER %>" HEIGHT="<%= GRAPH_HEIGHT %>"></TD>
					<%
					' We're now in the body of the chart.  Loop through the data showing the bars!
					For I = 0 To UBound(aValues)
						iBarHeight = Int((aValues(I) / iMaxValue) * GRAPH_HEIGHT)

						' This is a hack since browsers ignore a 0 as an image dimension!
						If iBarHeight = 0 Then iBarHeight = 1
						 %>
						<TD VALIGN="bottom"><IMG SRC="../images/spacer.gif" BORDER="0" WIDTH="<%= GRAPH_SPACER %>" HEIGHT="1"></TD>					
						<%
					Select Case RowNum
					  case 1:
					  %>
						<TD VALIGN="bottom"><IMG SRC="../images/spacer_red.gif" BORDER="0" WIDTH="<%= iBarWidth %>" HEIGHT="<%= iBarHeight %>" ALT="<%= aValues(I) %>"></A></TD>
						<%
						RowNum =RowNum+1
						
					  case 2:
					  %>
						<TD VALIGN="bottom"><IMG SRC="../images/spacer_Blue.gif" BORDER="0" WIDTH="<%= iBarWidth %>" HEIGHT="<%= iBarHeight %>" ALT="<%= aValues(I) %>"></A></TD>
						<%
						RowNum =RowNum+1
						
					  case 3:
					  %>
						<TD VALIGN="bottom"><IMG SRC="../images/spacer_Green.gif" BORDER="0" WIDTH="<%= iBarWidth %>" HEIGHT="<%= iBarHeight %>" ALT="<%= aValues(I) %>"></A></TD>
						<%
						RowNum =1

					End Select
				
					
					Next 
					%>
					</TR>
					<!-- I was using GRAPH_BORDER + GRAPH_WIDTH but it was moving the last x axis label -->
					<TR>
						<TD COLSPAN="<%= (2 * (UBound(aValues) + 1)) + 1 %>"><IMG SRC="../images/spacer_black.gif" BORDER="0" WIDTH="<%= GRAPH_BORDER + ((UBound(aValues) + 1) * (iBarWidth + GRAPH_SPACER)) %>" HEIGHT="<%= GRAPH_BORDER %>"></TD>
					</TR>
				<% ' The label array is optional and is really only useful for small data sets with very short labels! %>
				<% If IsArray(aLabels) Then %>
					<TR>
						<TD><!-- Spacing for Left Border Column --></TD>
					<% For I = 0 To (UBound(aValues)/3)  %>
						<TD><!-- Spacing for Spacer Column --></TD>
						<TD ALIGN="center" colspan="5"><FONT SIZE="1"><%= aLabels(I) %></FONT></TD>
					<% Next %>
					</TR>
				<% End If %>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD COLSPAN="2"><!-- Place holder for X Axis label centering--></TD>
			<TD ALIGN="center"><BR><B><%= strXAxisLabel %></B></TD>
		</TR>
	</TABLE>
	<%
End Sub
%>
<%
' Static Chart (with Bar Labels)
dim NowDate
Dim objConn
NowDate=FormatDateTime (Date, vbShortDate)

Chartyear="01"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set Jan=objConn.Execute("SELECT * FROM Stats WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='All'")
Set Feb=objConn.Execute("SELECT * FROM Stats WHERE Month = 'February' and Year='"&ChartYear&"' and HelpDesk='All'")
Set Mar=objConn.Execute("SELECT * FROM Stats WHERE Month = 'March' and Year='"&ChartYear&"' and HelpDesk='All'")
Set Apr=objConn.Execute("SELECT * FROM Stats WHERE Month = 'April' and Year='"&ChartYear&"' and HelpDesk='All'")
Set May=objConn.Execute("SELECT * FROM Stats WHERE Month = 'May' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Jun=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'June' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Jul=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'July' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Aug=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'August' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Sep=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'September' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Octb=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'October' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Nov=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'November' and Year='"&ChartYear&"' and HelpDesk='All'")
'Set Dec=objConn.Execute("SELECT * FROM ChartInfo WHERE Month = 'December' and Year='"&ChartYear&"' and HelpDesk='All'")
'January
If NowDate >= "1/1/01" and Nowdate<= "1/31/01" then
ShowChart Array(Jan("Total_Tickets"),Jan("Help_Tickets"),Jan("Request_Tickets")), Array("","January",""), "Help Desk Stats 2002", "Months", "# of Tickets"
end if
'February
If NowDate >= "2/1/02" and Nowdate<= "2/28/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets")), Array("","January","","","February",""), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
'March
If NowDate >= "3/1/02" and Nowdate<= "3/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets")), Array("January","February","March"), "Help Desk Stats 2002", "Months", "# of Tickets"
end if
'April
If NowDate >= "4/1/02" and Nowdate<= "4/30/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets")), Array("","January","","","February","","","March","","","April",""), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
'May
If NowDate >= "5/1/02" and Nowdate<= "5/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets")), Array("","January","","","February","","","March","","","April","","","May",""), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
'June
If NowDate >= "6/1/02" and Nowdate<= "6/30/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets")), Array("January","February","March","April","May","June"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
If NowDate >= "7/1/02" and Nowdate<= "7/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets")), Array("January","February","March","April","May","June","July"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
If NowDate >= "8/1/02" and Nowdate<= "8/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets"),Aug("ToatalTickets"),Aug("HelpTickets"),Aug("RequestTickets")), Array("January","February","March","April","May","June","July","August"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
If NowDate >= "9/1/02" and Nowdate<= "9/30/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets"),Aug("ToatalTickets"),Aug("HelpTickets"),Aug("RequestTickets"),Sep("ToatalTickets"),Sep("HelpTickets"),Sep("RequestTickets")), Array("January","February","March","April","May","June","July","August","September"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
If NowDate >= "10/1/02" and Nowdate<= "10/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets"),Aug("ToatalTickets"),Aug("HelpTickets"),Aug("RequestTickets"),Sep("ToatalTickets"),Sep("HelpTickets"),Sep("RequestTickets"),Octb("ToatalTickets"),Octb("HelpTickets"),Octb("RequestTickets")), Array("January","February","March","April","May","June","July","August","September","October"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if
If NowDate >= "11/1/02" and Nowdate<= "11/30/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets"),Aug("ToatalTickets"),Aug("HelpTickets"),Aug("RequestTickets"),Sep("ToatalTickets"),Sep("HelpTickets"),Sep("RequestTickets"),Octb("ToatalTickets"),Octb("HelpTickets"),Octb("RequestTickets"),Nov("ToatalTickets"),Nov("HelpTickets"),Nov("RequestTickets")), Array("January","February","March","April","May","June","July","August","September","October","November"), "Help Desk Stats 2002", "Months", "# of Tickets"
end if
If NowDate >= "12/1/02" and Nowdate<= "12/31/02" then
ShowChart Array(Jan("TotalTickets"),Jan("HelpTickets"),Jan("RequestTickets"),Feb("TotalTickets"),Feb("HelpTickets"),Feb("RequestTickets"),Mar("TotalTickets"),Mar("HelpTickets"),Mar("RequestTickets"),Apr("TotalTickets"),Apr("HelpTickets"),Apr("RequestTickets"),May("TotalTickets"),May("HelpTickets"),May("RequestTickets"),Jun("ToatalTickets"),Jun("HelpTickets"),Jun("RequestTickets"),Jul("ToatalTickets"),Jul("HelpTickets"),Jul("RequestTickets"),Aug("ToatalTickets"),Aug("HelpTickets"),Aug("RequestTickets"),Sep("ToatalTickets"),Sep("HelpTickets"),Sep("RequestTickets"),Octb("ToatalTickets"),Octb("HelpTickets"),Octb("RequestTickets"),Nov("ToatalTickets"),Nov("HelpTickets"),Nov("RequestTickets"),Dec("ToatalTickets"),Dec("HelpTickets"),Dec("RequestTickets")), Array("January","February","March","April","May","June","July","August","September","October","November","December"), "Help Desk Stats 2002", "Months", "# of Tickets"
End if

' Spacing
Response.Write "<BR>" & vbCrLf
Response.Write "<table border=0 align=center>"
Response.Write "<tr><td><img src=../images/spacer_Red.gif width=50 height=50</td><td><img src=../images/spacer_blue.gif width=50 height=50</td><td><img src=../images/spacer_Green.gif width=50 height=50</td></tr>"
Response.Write "<tr><td>Total Tickets  </td><td>Help Tickets  </td><td>Request Tickets  </td></tr>"
Response.Write "</table>"
Response.Write "<BR>" & vbCrLf
Response.Write "<BR>" & vbCrLf
Response.Write "<table border=1 align=center>"
Response.Write "<tr><td><b>Month</b></td><td><b>Total<br>Tickets</b></td><td><b>Help<br>Tickets</b></td><td><b>Request<br>Tickets</b></td></tr>"

If NowDate >= "1/1/02" then
Response.Write "<tr><td>January</td><td>" & Jan("TotalTickets") & "</td><td>" & Jan("HelpTickets") & "</td><td>" & Jan("RequestTickets") & "</td></tr>"
End if
If NowDate >= "2/1/02" then
Response.Write "<tr><td>February</td><td>" & Feb("TotalTickets") & "</td><td>" & Feb("HelpTickets") & "</td><td>" & Feb("RequestTickets") & "</td></tr>"
End if
If NowDate >= "3/1/02" then
Response.Write "<tr><td>March</td><td>" & Mar("TotalTickets") & "</td><td>" & Mar("HelpTickets") & "</td><td>" & Mar("RequestTickets") & "</td></tr>"
End if
If NowDate >= "4/1/02" then
Response.Write "<tr><td>April</td><td>" & Apr("TotalTickets") & "</td><td>" & Apr("HelpTickets") & "</td><td>" & Apr("RequestTickets") & "</td></tr>"
End if
If NowDate >= "5/1/02" then
Response.Write "<tr><td>May</td><td>" & May("TotalTickets") & "</td><td>" & May("HelpTickets") & "</td><td>" & May("RequestTickets") & "</td></tr>"
End if
If NowDate >= "6/1/02" then
Response.Write "<tr><td>June</td><td>" & Jun("TotalTickets") & "</td><td>" & Jun("HelpTickets") & "</td><td>" & Jun("RequestTickets") & "</td></tr>"
End if
If NowDate >= "7/1/02" then
Response.Write "<tr><td>July</td><td>" & Jul("TotalTickets") & "</td><td>" & Jul("HelpTickets") & "</td><td>" & Jul("RequestTickets") & "</td></tr>"
End if
If NowDate >= "8/1/02" then
Response.Write "<tr><td>August</td><td>" & Aug("TotalTickets") & "</td><td>" & Aug("HelpTickets") & "</td><td>" & Aug("RequestTickets") & "</td></tr>"
End if
If NowDate >= "9/1/02" then
Response.Write "<tr><td>September</td><td>" & Sep("TotalTickets") & "</td><td>" & Sep("HelpTickets") & "</td><td>" & Sep("RequestTickets") & "</td></tr>"
End if
If NowDate >= #10/1/02# then
Response.Write "<tr><td>October</td><td>" & Octb("TotalTickets") & "</td><td>" & Octb("HelpTickets") & "</td><td>" & Octb("RequestTickets") & "</td></tr>"
End if
If NowDate >= "11/1/02" then
Response.Write "<tr><td>November</td><td>" & Nov("TotalTickets") & "</td><td>" & Nov("HelpTickets") & "</td><td>" & Nov("RequestTickets") & "</td></tr>"
End if
If NowDate >= "12/1/02" then
Response.Write "<tr><td>December</td><td>" & Dec("TotalTickets") & "</td><td>" & Dec("HelpTickets") & "</td><td>" & Dec("RequestTickets") & "</td></tr>"
end if
Response.Write "</table>"

objConn.Close
Set objConn = Nothing
%>
