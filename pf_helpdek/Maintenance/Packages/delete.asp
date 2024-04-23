<%
	Dim sqlId
	Dim objConn
	Dim objRec
	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "GROUP_MEMBERS", objConn	
	
	objConn.Execute "DELETE * FROM GROUP_MEMBERS WHERE PACK_ID= " & Request("Num") & ";"

	Response.Redirect "index.asp"

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
%>