<%

Dim strCategory
Dim strGroup
Dim strPackage
Dim objConn
Dim objRec
Dim sql

strCategory = Replace(Request.Form("Category"), "'", "''")
strGroup = Replace(Request.Form("Group"), "'", "''")
strPackage = Replace(Request.Form("Package"), "'", "''")
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "DSN=Helpdesk2"
objRec.Open "GROUP_MEMBERS", objConn



sql = "INSERT INTO GROUP_MEMBERS(HW_SW_ID, GROUP_NAME, PACKAGE_NAME) VALUES('"& strCategory & "', '"& strGroup & "', '"& strPackage & "');"
		
objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("index.asp")

%>