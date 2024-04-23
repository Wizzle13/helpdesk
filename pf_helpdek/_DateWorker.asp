<%
strDate=FormatDateTime (Date, vbShortDate)
strMonth=Month(date)

strYear=Year(date)
%>
<tr>
<td valign="top">By Opened/Closed Date:
</td>
<td>
<table border="1">
<form method="post" action="SearchDateWorker.asp" name="ViewWork">
<tr>
	<td>Start Date:</td>
	<td><Select name=StartMonth>
<%
Select Case strMonth
	Case 1
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 2
	response.write "<Option value=1>January <Option Selected value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 3
	response.write "<Option value=1>January <Option value=2>February <Option Selected value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 4
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option Selected value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 5
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option Selected value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 6
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option Selected value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 7
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option Selected value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 8
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option Selected value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 9
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option Selected value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 10
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option Selected value=10>October <Option value=11>November <Option value=12>December"
	Case 11
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option Selected value=11>November <Option value=12>December"
	Case 12
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option Selected value=12>December"
	
End select
 %>	</select></td>
	
	<td><Select name=StartDay>
	<Option value=1 Selected>1
	<Option value=2>2
	<Option value=3>3
	<Option value=4>4
	<Option value=5>5
	<Option value=6>6
	<Option value=7>7
	<Option value=8>8
	<Option value=9>9
	<Option value=10>10
	<Option value=11>11
	<Option value=12>12
	<Option value=13>13
	<Option value=14>14
	<Option value=15>15
	<Option value=16>16
	<Option value=17>17
	<Option value=18>18
	<Option value=19>19
	<Option value=20>20
	<Option value=21>21
	<Option value=22>22
	<Option value=23>23
	<Option value=24>24
	<Option value=25>25
	<Option value=26>26
	<Option value=27>27
	<Option value=28>28
	<Option value=29>29
	<Option value=30>30
	<Option value=31>31
	</select></td>
	<td><Select name=StartYear>
<%
Select Case strYear
	Case 1999
	response.write "<Option selected value=99>1999 <Option value=00>2000 <Option value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2000
	response.write "<Option value=99>1999 <Option selected value=00>2000 <Option value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2001
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2002
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option selected value=02>2002 <Option value=03>2003"
	Case 2003
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003"
	Case 2004
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option selected value=04>2004"	
	Case 2005
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option value=04>2004<Option selected value=05>2005<Option value=06>2006"	
	Case 2006
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option value=04>2004<Option value=05>2005<Option selected value=06>2006"	
End select
 %>
	
	</select></td>

</tr>

<tr>
	<td>End Date:</td>
	<td><Select name=EndMonth>
<%
Select Case StrMonth
	Case 1
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 2
	response.write "<Option value=1>January <Option Selected value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 3
	response.write "<Option value=1>January <Option value=2>February <Option Selected value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 4
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option Selected value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 5
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option Selected value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 6
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option Selected value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 7
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option Selected value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 8
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option Selected value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 9
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option Selected value=9>September <Option value=10>October <Option value=11>November <Option value=12>December"
	Case 10
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option Selected value=10>October <Option value=11>November <Option value=12>December"
	Case 11
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option Selected value=11>November <Option value=12>December"
	Case 12
	response.write "<Option value=1>January <Option value=2>February <Option value=3>March <Option value=4>April <Option value=5>May <Option value=6>June <Option value=7>July <Option value=8>August <Option value=9>September <Option value=10>October <Option value=11>November <Option Selected value=12>December"
	
End select
 %>		</select></td>
	
	<td><Select name=EndDay>
<%
if strMonth = 1 OR strMonth = 3 OR strMonth = 5 OR strMonth = 7 OR strMonth = 8 OR strMonth = 10 OR strMonth = 12 Then
response.write"<Option value=1>1 <Option value=2>2 <Option value=3>3 <Option value=4>4 <Option value=5>5 <Option value=6>6 <Option value=7>7 <Option value=8>8 <Option value=9>9 <Option value=10>10 <Option value=11>11 <Option value=12>12 <Option value=13>13 <Option value=14>14 <Option value=15>15 <Option value=16>16 <Option value=17>17 <Option value=18>18 <Option value=19>19 <Option value=20>20 <Option value=21>21 <Option value=22>22 <Option value=23>23 <Option value=24>24 <Option value=25>25 <Option value=26>26 <Option value=27>27 <Option value=28>28 <Option value=29>29 <Option value=30>30 <Option value=31 selected>31" 
End if

if strMonth = 4 OR strMonth = 6 OR strMonth = 9 OR strMonth = 11 Then
response.write"<Option value=1>1 <Option value=2>2 <Option value=3>3 <Option value=4>4 <Option value=5>5 <Option value=6>6 <Option value=7>7 <Option value=8>8 <Option value=9>9 <Option value=10>10 <Option value=11>11 <Option value=12>12 <Option value=13>13 <Option value=14>14 <Option value=15>15 <Option value=16>16 <Option value=17>17 <Option value=18>18 <Option value=19>19 <Option value=20>20 <Option value=21>21 <Option value=22>22 <Option value=23>23 <Option value=24>24 <Option value=25>25 <Option value=26>26 <Option value=27>27 <Option value=28>28 <Option value=29>29 <Option selected value=30>30 <Option value=31>31" 
End if

if strMonth = 2 Then
response.write"<Option value=1>1 <Option value=2>2 <Option value=3>3 <Option value=4>4 <Option value=5>5 <Option value=6>6 <Option value=7>7 <Option value=8>8 <Option value=9>9 <Option value=10>10 <Option value=11>11 <Option value=12>12 <Option value=13>13 <Option value=14>14 <Option value=15>15 <Option value=16>16 <Option value=17>17 <Option value=18>18 <Option value=19>19 <Option value=20>20 <Option value=21>21 <Option value=22>22 <Option value=23>23 <Option value=24>24 <Option value=25>25 <Option value=26>26 <Option value=27>27 <Option selected value=28>28 <Option value=29>29 <Option value=30>30 <Option value=31>31" 
End if
 %>
	
		</select></td>
	<td><Select name=EndYear>
	<%
Select Case strYear
	Case 1999
	response.write "<Option selected value=99>1999 <Option value=00>2000 <Option value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2000
	response.write "<Option value=99>1999 <Option selected value=00>2000 <Option value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2001
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option value=03>2003"
	Case 2002
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option selected value=02>2002 <Option value=03>2003"
	Case 2003
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003"
	Case 2004
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option selected value=04>2004"
	Case 2005
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option value=04>2004<Option selected value=05>2005<Option value=06>2006"	
	Case 2006
	response.write "<Option value=99>1999 <Option value=00>2000 <Option selected value=01>2001 <Option value=02>2002 <Option selected value=03>2003 <Option value=04>2004<Option value=05>2005<Option selected value=06>2006"	
End select
 %>

	</select></td>

</tr>

<%
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from Workers where active_worker='Yes' order by First_Name, Last_Name")
	sqlId ="SELECT WORKER_ID FROM Workers;"
Response.Write "<td colspan='2'><Select name=Worker_ID><Option Selected Value='All'>All"
	While Not objRs.EOF
				Response.Write"<Option Value="& objRs("Worker_ID") &">"& objRs("First_Name") &" "& objRs("Last_Name")
		objRs.MoveNext
	Wend
	objConn.Close
	Set objConn = Nothing	
Response.Write "</td>"
Response.Write "<td colspan='2'><input type=radio value =Opened name=View checked>Date Opened<br><input type=radio value=Closed name=View>Date Closed</td>"	
	 %>


	<tr><td colspan="4" align="Right"><input type="submit" value="Search" name="B1"></td>
</tr>
</form>
</table>
</td>
</tr>
