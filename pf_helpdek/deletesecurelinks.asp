<%

Dim sqlId
Dim objConn
Dim objRec

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")

objConn.Open "DSN=Helpdesk2"
objRec.Open "SecureLinks", objConn	

objConn.Execute "DELETE * FROM SecureLinks WHERE LinkID= " & Request("Num") & ";"

objRec.Close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect Request.ServerVariables("HTTP_Referer")
%>