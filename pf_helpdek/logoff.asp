<%

Dim sqlId
Dim objConn
Dim objRec
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
strEmail= session("email")

objConn.Open "DSN=Helpdesk2"
	objRec.Open "WhosOn", objConn	
	Set objRec=objConn.Execute("Select * from WhosOn where Logon='" & strEmail & "' ")
		sql = "UPDATE Whoson SET  Status= 'Off' WHERE LogOn= ('" & strEmail & "')"
		objConn.Execute(sql)
	
'objConn.Open "DSN=Helpdesk2"
'objRec.Open "WhosOn", objConn	
	
'objConn.Execute "DELETE * FROM WhosOn WHERE LogOn= ('" & strEmail & "')"

objRec.Close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Session("UID") = ""
Session("WorkID") = ""
Session("IsValid") = "False"
Response.Redirect "login.asp"
%>