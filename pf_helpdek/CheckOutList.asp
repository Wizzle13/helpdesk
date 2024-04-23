
<%

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select * from CheckOutItems order by ID")
Response.Write "<table border=1 cellpadding=0 cellborder=0>"
While Not objRs.EOF
	
	Response.Write "<tr><td Colspan=6 bgcolor=#A8EAA8>"& objRs("CheckOutItem") &"</td><td bgcolor=#A8EAA8>"%><form method="post" action="CheckOutAdd.asp" name="Inputform" onsubmit="return Join_Form1_Validator(this)"><input type="submit" value="Add" name="B1"></form><form method="post" action="CheckOutListAll.asp" name="Inputform" onsubmit="return Join_Form1_Validator(this)"><input type="submit" value="View All" name="B2"></form></td>
	<% 
	Response.Write "<td bgcolor=#A8EAA8>"& objRs("Info") &"</td></tr>"
	Response.Write "<tr><td>Modify</td>"
	Response.Write "<td>Name</td>"
	Response.Write "<td>Date Out</td>"
	Response.Write "<td>Time Out</td>"
	Response.Write "<td>Date In</td>"
	Response.Write "<td>Time In</td>"
	Response.Write "<td>Ticket Number</td>"
	Response.Write "<td>Location</td></tr>"
	
	
		Set objRs2=objConn.Execute("Select * from CheckOutList  Where Returned='No' and CheckoutItemID="& objRs("ID") &" order by Date_Out,Time_Out,Date_In,Time_In")
	
		While Not objRs2.EOF
		Response.Write "<tr><td><a href=""CheckoutModify.asp?Num=" & objRs2("ID") & """>Modify</td>"
		Response.Write "<td>"& objRs2("UsersName") &"</td>"
		Response.Write "<td>"& objRs2("Date_Out") &"</td>"
		Response.Write "<td>"& objRs2("Time_Out") &"</td>"
		Response.Write "<td>"& objRs2("Date_In") &"</td>"
		Response.Write "<td>"& objRs2("Time_In") &"</td>"
		Response.Write "<td><a href=""modifyticket.asp?Num=" & objRs2("Ticket_Number") & """>"& objRs2("Ticket_Number") &"</td>"
		Response.Write "<td>"& objRs2("Location") &"</td></tr>"
		
		objRs2.MoveNext
		Wend
		Response.Write "</tr>"
objRs.MoveNext
Wend
	Response.Write "</table>"
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
%>
