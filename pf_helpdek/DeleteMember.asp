<%
	Dim sqlId
	Dim objConn
	Dim objRec
	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "UserInfo", objConn	
	
	objConn.Execute "DELETE * FROM UserInfo WHERE ID= " & Request("Num") & ";"

	Response.Redirect "ManageUsers.asp"

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
%>