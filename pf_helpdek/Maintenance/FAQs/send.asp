<%

Dim strQuestion
Dim strAnwser
Dim objConn
Dim objRec
Dim sql

strQuestion = Replace(Request.Form("Question"), "'", "''")
strAnwser = Replace(Request.Form("Anwser"), "'", "''")
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "DSN=Helpdesk2"
objRec.Open "FAQs", objConn

sql = "INSERT INTO FAQs(Question, Anwser) VALUES('"& strQuestion & "', '"& strAnwser & "');"
		
objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("index.asp")

%>