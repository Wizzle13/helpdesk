<%
	Dim objConn
	Dim objRec
	Dim sql
	Dim sql2

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "UserInfo", objConn
	
	sql = "UPDATE UserInfo SET  FirstName= '" & Request.Form("FirstName") & "', LastName= '" & Request.Form("LastName") & "', Email= '" & Request.Form("Email") & "', Country= '" & Request.Form("Country") & "', Extention= '" & Request.Form("Extention") & "', Department= '" & Request.Form("Department") & "', Status= '" & Request.Form("Status") & "', WorkerID= '" & Request.Form("WorkerID") & "', Password= '" & Request.Form("Password") & "' WHERE ID=" & Request("Num") & ";"
	
	objConn.Execute(sql)
	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

	Response.Redirect "ManageUsers.asp"
%>

