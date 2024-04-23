<font face="Arial" size="4">
IT Helpdesk Customer Satisfaction Survey<BR>

<%
If Request("Num") = "" or Request("Num") = "3XYA" Then
	TickNum = 0
	Response.Write "<P><font face='Arial' size='3'><B>Please enter your help ticket number.</B>"
	Response.Write "<form action=Survey.asp name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='20'><BR>"
	Response.Write "<input type=Submit value='Submit'></font>"
Else
	TickNum = Request("Num")
%>
<P>
For ticket number <%=Request("Num")%>.

<%

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "HelpDesk2"
sqlId ="SELECT * FROM Calls WHERE Ticket_Number=" & TickNum & ";"

objRec.Open sqlId, objConn

If objRec("Call_Type") = "REQUEST" Then
	Response.Write "<P>This ticket number is classified as a request and is not valid for our survey."
	Response.Write "<P><font face='Arial' size='3'><B>Please enter your help ticket number.</B>"
	Response.Write "<form action=Survey.asp name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='20'><BR>"
	Response.Write "<input type=Submit value='Submit'></font>"
Else

%>
</FONT>

<form method="post" action="Survey_Send.asp?Num=<%=TickNum%>">
<input type="Hidden" name="Survey_Time" value=<%=FormatDateTime (Time, vbShortTime)%>>
<input type="Hidden" name="Survey_Date" value=<%=FormatDateTime (Date, vbShortDate)%>>
<input type="Hidden" name="Ticket_Number" value="<%=TickNum%>">

1. On a scale of 1 (worst) to 10 (best), was your problem or issue answered acted on and resolved in a timely manner?<BR>
	&nbsp;&nbsp;1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10<BR>
	<input type=radio value=1 name=Timely_Manner>
	<input type=radio value=2 name=Timely_Manner>
	<input type=radio value=3 name=Timely_Manner>
	<input type=radio value=4 name=Timely_Manner>
	<input type=radio value=5 name=Timely_Manner>
	<input type=radio value=6 name=Timely_Manner>
	<input type=radio value=7 name=Timely_Manner>
	<input type=radio value=8 name=Timely_Manner>
	<input type=radio value=9 name=Timely_Manner>
	<input type=radio value=10 name=Timely_Manner>
<P>
2. Did the timeliness meet your expectation?<BR>
	&nbsp;&nbsp;Yes:&nbsp;<input type=radio value=Yes name=Timely_Manner_Expectation><BR>
	&nbsp;&nbsp;No:&nbsp;&nbsp;<input type=radio value=No name=Timely_Manner_Expectation>
<P>

3. Was your problem resolved correctly the first time?<BR>
	&nbsp;&nbsp;Yes:&nbsp;<input type=radio value=Yes name=Problem_Resolution><BR>
	&nbsp;&nbsp;No:&nbsp;&nbsp;<input type=radio value=No name=Problem_Resolution>
<P>

4. On a scale of 1 (worst) to 10 (best), rate your overall experience for this ticket.<BR>
	&nbsp;&nbsp;1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;10<BR>
	<input type=radio value=1 name=Overall_Experience>
	<input type=radio value=2 name=Overall_Experience>
	<input type=radio value=3 name=Overall_Experience>
	<input type=radio value=4 name=Overall_Experience>
	<input type=radio value=5 name=Overall_Experience>
	<input type=radio value=6 name=Overall_Experience>
	<input type=radio value=7 name=Overall_Experience>
	<input type=radio value=8 name=Overall_Experience>
	<input type=radio value=9 name=Overall_Experience>
	<input type=radio value=10 name=Overall_Experience>
<P>

5. Were you treated the way you would like to be treated?<BR>
	&nbsp;&nbsp;Yes:&nbsp;<input type=radio value=Yes name=Treated_Answer><BR>
	&nbsp;&nbsp;No:&nbsp;&nbsp;<input type=radio value=No name=Treated_Answer>
<BR>

	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;a. If No explain: <BR>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<textarea cols=50 rows=5 name="Treated_Answer_No"></textarea><p>

6. Other Comments About Ticket <%=Request("Num")%>:<BR>
	<textarea cols=50 rows=5 name="Other_Comments"></textarea><p>

<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
<%
End If

objRec.Close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

End If
%>
</form>
</BODY>
</HTML>

