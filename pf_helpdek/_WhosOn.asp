<%
'This page writes information to the WhosOn Table in the database
Dim objConn
Dim objRec
Dim sql
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
	strDate=FormatDateTime (Date, vbShortDate)
	strTime=FormatDateTime (Time, vbShortTime)
	strEmail= session("Email")
	strFirstName= session("FirstName")
	strLastName= session("LastName")
	strCountry= session("Country")
	
Set objREC=objConn.Execute("Select * from WhosOn")
Set rsCount=objConn.Execute("SELECT Count(*) from WhosOn WHERE Logon = ('" & strEmail & "')")

If rsCount(0) > 0 then
	Set objRec=objConn.Execute("Select * from WhosOn where Logon='" & strEmail & "' ")
		sql = "UPDATE Whoson SET  Status= 'On',LogonDate='"& strDate & "',LogonTime='"& strTime & "' WHERE ID=" & objRec("ID") & ";"
		objConn.Execute(sql)
	'objConn.Execute "DELETE * FROM Whoson WHERE Logon = ('" & strEmail & "')"
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

Else
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "WhosOn", objConn

	sql = "INSERT INTO WhosOn(Logon, FirstName, LastName,LogonDate,LogonTime,Country,Status) VALUES('"& strEmail & "','"& strFirstName & "','"& strLastName & "','"& strDate & "','"& strTime & "','"& strCountry & "','On');"
		
	objConn.Execute(sql)

	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

End IF

%>