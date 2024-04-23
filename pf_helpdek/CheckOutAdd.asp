<link rel="stylesheet" href="helpdesk.css" type="text/css">
<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This Page is the Join page for the User interface for the Helpdesk.
This page uses JavaScript to check that all the information is entered.

-->
<HTML>
<HEAD>
  <TITLE>Check Out</TITLE>


</HEAD>
<body>


<P>


<form method="post" action="CheckOutSend.asp" name="Inputform"  name='Join_Form1'>
<table border="1">
<tr>
<td>Name:</td><td><input type="text" name="UsersName" size="20"></td>
</tr>
<tr>
<td>Date Checked out:<font color="Red">*</font></td><td><input type="Date" name="Date_Out" size="20"></td>
</tr>
<tr>
<td>Time Checked out:<font color="Red">*</font></td><td><input type="Date" name="Time_Out" size="20"></td>
</tr>
<tr>
<td>Date Checked In:<font color="Red">*</font></td><td><input type="Date" name="Date_In" size="20"></td>
</tr>
<tr>
<td>Time Checked In:<font color="Red">*</font></td><td><input type="Date" name="Time_In" size="20"></td>
</tr>
<tr>
<td>Ticket Number:</td><td><input type="text" name="Ticket_Number" size="20" value="N/A"></td>
</tr>
<tr>
<td>Location:</td><td><input type="text" name="Location" size="20" value="N/A"></td>
</tr>
<tr>
<td>Check Out Item:</td><td><Select name="CheckoutItemID">
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select * from CheckOutItems order by ID")
		While Not objRs.EOF
			Response.Write "<Option value="& objRs("ID") &">"& objRs("CheckOutItem") &""
			objRs.MoveNext
	Wend
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
	
 %>		
 </select></td>
 </tr>
</table>
<input type="Hidden" name="UserAdded" Value="1">
<br>
<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
</form>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>

