
<%

Dim rsCount
Dim objConn
Dim objRec
Dim sql

strUsersName = Request.Form("UsersName")
strDate_Out = Request.Form("Date_Out")
strTime_Out = Request.Form("Time_Out")
strDate_In = Request.Form("Date_In")
strTime_In = Request.Form("Time_In")
strLocation = Request.Form("Location")
strTicket_Number = Request.Form("Ticket_Number")
strCheckoutItemID = Request.Form("CheckoutItemID")


Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set objREC=objConn.Execute("Select * from CheckOutList")


sql = "INSERT INTO CheckOutList(UsersName, Date_Out, Time_Out, Date_In, Time_In, Location, CheckOutItemID, Ticket_Number, Returned) VALUES('"& strUsersName & "','"& strDate_Out & "','"& strTime_Out & "','"& strDate_In & "','"& strTime_In & "','"& strLocation & "','"& strCheckOutItemID & "','"& strTicket_Number & "','No');"
		
objConn.Execute(sql)

objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("CheckOutList.asp")


%>