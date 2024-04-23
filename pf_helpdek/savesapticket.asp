<%
Dim objConn
Dim objRec
Dim objRs
Dim sql
Dim sql2
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

sql = "UPDATE Calls SET  SAPPriority = '" & Request.Form("tickpriority") & "' WHERE Ticket_Number=" & Request("Num") & ";"

objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect Request.ServerVariables("HTTP_Referer")

'Response.Redirect "javascript:history.go(-1)"
	
%>

