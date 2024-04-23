<%
Sub ShowChart(ByRef aValues, ByRef aLabels, ByRef strTitle, ByRef strXAxisLabel, ByRef strYAxisLabel)
	' Some user changable graph defining constants
	' All units are in screen pixels
	Const GRAPH_WIDTH  = 450  ' The width of the body of the graph
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
						<TD ROWSPAN="2"><IMG SRC="./images/spacer.gif" BORDER="0" WIDTH="1" HEIGHT="<%= GRAPH_HEIGHT %>"></TD>
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
						<TD VALIGN="bottom"><IMG SRC="./images/spacer_black.gif" BORDER="0" WIDTH="<%= GRAPH_BORDER %>" HEIGHT="<%= GRAPH_HEIGHT %>"></TD>
					<%
					' We're now in the body of the chart.  Loop through the data showing the bars!
					For I = 0 To UBound(aValues)
						iBarHeight = Int((aValues(I) / iMaxValue) * GRAPH_HEIGHT)

						' This is a hack since browsers ignore a 0 as an image dimension!
						If iBarHeight = 0 Then iBarHeight = 1
					%>
						<TD VALIGN="bottom"><IMG SRC="./images/spacer.gif" BORDER="0" WIDTH="<%= GRAPH_SPACER %>" HEIGHT="1"></TD>
						<TD VALIGN="bottom"><IMG SRC="./images/spacer_red.gif" BORDER="0" WIDTH="<%= iBarWidth %>" HEIGHT="<%= iBarHeight %>" ALT="<%= aValues(I) %>"></A></TD>
					<%
					Next 'I
					%>
					</TR>
					<!-- I was using GRAPH_BORDER + GRAPH_WIDTH but it was moving the last x axis label -->
					<TR>
						<TD COLSPAN="<%= (2 * (UBound(aValues) + 1)) + 1 %>"><IMG SRC="./images/spacer_black.gif" BORDER="0" WIDTH="<%= GRAPH_BORDER + ((UBound(aValues) + 1) * (iBarWidth + GRAPH_SPACER)) %>" HEIGHT="<%= GRAPH_BORDER %>"></TD>
					</TR>
				<% ' The label array is optional and is really only useful for small data sets with very short labels! %>;
				<% If IsArray(aLabels) Then %>
					<TR>
						<TD><!-- Spacing for Left Border Column --></TD>
					<% For I = 0 To UBound(aValues)  %>
						<TD><!-- Spacing for Spacer Column --></TD>
						<TD ALIGN="center"><FONT SIZE="1"><%= aLabels(I) %></FONT></TD>
					<% Next 'I %>;
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
startdate= request("StartMonth") & "/" & request("StartDay") & "/" &request("StartYear")
Enddate= request("EndMonth") & "/" & request("EndDay") & "/" &request("EndYear")
 
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # ))  Order by Date_Opened, Date_Closed;")
Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set OpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # ))  ;")
Set CloseCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set SOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) And (((Calls.Date_Closed)=#09/22/78#)) ;")


ShowChart Array(OpenCount(0), CloseCount(0), SOpenCount(0)), Array("All", "Help", "Request"), "Chart Title", "X Label", "Y Label"


' Spacing
Response.Write "<BR>" & vbCrLf
Response.Write "<BR>" & vbCrLf
Response.Write "<BR>" & vbCrLf


' Random number chart
'Dim I
'Dim aTemp(49)

'Randomize
'For I = 0 to 49
'	aTemp(I) = Int((50 + 1) * Rnd)
'Next 'I

' Chart made from random numbers (without Bar Labels)
'ShowChart aTemp, "Note that this isn't an Array!", "Chart of 50 Random Numbers", "Index", "Value"
%>
