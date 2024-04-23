<%
	Dim objConn
	Dim objRec
	Dim sql
	Dim sql2

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "UserInfo", objConn
	
	sql = "UPDATE UserInfo SET  FirstName= '" & Request.Form("FirstName") & "', LastName= '" & Request.Form("LastName") & "',  Email= '" & Request.Form("Email") & "', Extention= '" & Request.Form("Extention") & "', Department= '" & Request.Form("Department") & "', Status= '" & Request.Form("Status") & "', Country= '" & Request.Form("Country") & "', WorkerID= '" & Request.Form("WorkerID") & "', Password= '" & Request.Form("Password") & "' WHERE ID=" & Request("Num") & ";"
	
	objConn.Execute(sql)
	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
	
	Session("IsValid") = "True"
	Session("FirstName") = Request.Form("FirstName")
	Session("LastName") = Request.Form("LastName")
	Session("Extention") = Request.Form("Extention")
	Session("Email") = Request.Form("Email")
	Session("Status") = Request.Form("Status")
	Session("WorkID") = Request.Form("WorkerID")
	Session("Password") = Request.Form("Password")
	Session("Department") = Request.Form("Department")
	Session("Country") = Request.Form("Country")
	
	Response.Redirect "main.asp"
%>

