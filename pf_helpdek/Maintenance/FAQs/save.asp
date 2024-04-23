<%
	Dim objConn
	Dim objRec
	Dim sql
	Dim sql2

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=HelpDesk2"
	objRec.Open "FAQs", objConn
	
	sql = "UPDATE FAQs SET  Question= '" & Request.Form("Question") & "', Anwser= '" & Request.Form("Anwser") & "' WHERE FAQID=" & Request("Num") & ";"
	
	objConn.Execute(sql)
	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

	Response.Redirect "Index.asp"
%>
