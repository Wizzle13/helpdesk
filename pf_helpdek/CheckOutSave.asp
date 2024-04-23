
<%

Dim rsCount
Dim objConn
Dim objRec
Dim sql

Set objConn=Server.CreateObject("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")

objConn.Open "DSN=Helpdesk2"
objRec.Open "CheckOutList", objConn

sql = "UPDATE CheckOutList SET  UsersName= '" & Request.Form("UsersName") & "', Date_Out= '" & Request.Form("Date_Out") & "', Time_Out= '" & Request.Form("Time_Out") & "', Date_In= '" & Request.Form("Date_In") & "', Time_In= '" & Request.Form("Time_In") & "', Location= '" & Request.Form("Location") & "', Ticket_Number= '" & Request.Form("Ticket_Number") & "', CheckOutItemID= '" & Request.Form("CheckOutItemID") & "', Returned= '" & Request.Form("Returned") & "' WHERE ID=" & Request.form("Num") & ";"
		
objConn.Execute(sql)

objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("CheckOutList.asp")


%>