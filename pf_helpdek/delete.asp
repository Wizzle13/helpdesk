<%
	Dim sqlId
	Dim objConn
	Dim objRec
	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=HelpDesk2"
	objRec.Open "Calls", objConn	
	
	objConn.Execute "DELETE * FROM Calls WHERE Ticket_Number= " & Request("Num") & ";"

	Response.Redirect "main.asp"

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
%>