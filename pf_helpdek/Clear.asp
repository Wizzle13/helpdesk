<%


Dim sqlId
	Dim objConn
	Dim objRec
	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=Helpdesk2"
	objRec.Open "WhosOn", objConn	
	Set objRec=objConn.Execute("Select * from WhosOn")


		sql = "UPDATE Whoson SET Status= 'Off' WHERE LogonDate <= #" & FormatDateTime (Date -2, vbShortDate) & "#;"
		objConn.Execute(sql)
	
	Response.Redirect "WhosOn.asp"

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing


%>