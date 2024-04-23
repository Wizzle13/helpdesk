<%
	Dim objConn
	Dim objRec
	Dim sql
	Dim sql2

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=HelpDesk2"
	objRec.Open "Domains", objConn
	
	sql = "UPDATE Domains SET  Domain= '" & Request.Form("Domain") & "', Country= '" & Request.Form("Country") & "', CountryFull='" & Request.Form("CountryFull") & "', Helpdesk='" & Request.Form("Helpdesk") & "' WHERE ID=" & Request("Num") & ";"
	
	objConn.Execute(sql)
	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

	Response.Redirect "Index.asp"
%>
