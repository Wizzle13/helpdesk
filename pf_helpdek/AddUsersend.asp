
<%

Dim rsCount
Dim objConn
Dim objRec
Dim sql

strFirstName = Request.Form("First_Name")
strLastName = Request.Form("Last_Name")
strExtention = Request.Form("Extention")
strPassword = Request.Form("Password")
strDepartment = Request.Form("Department")
strCountry = Request.Form("Country")
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
strEmail=Request.Form("Logon_Name")
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set objREC=objConn.Execute("Select * from UserInfo")
Set rsCount=objConn.Execute("SELECT Count(*) from UserInfo WHERE Email = ('" & strEmail & "')")

If rsCount(0) > 0 then
Response.Write "This E-mail already exists in our database.  If you think this is incorrect, please contact the IT Help Desk. If you have forgotten your password, click <a href=login.asp>here</a>."

objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Else
objRec.close
objRec.Open "UserInfo", objConn

sql = "INSERT INTO UserInfo(FirstName, LastName, Extention, Email, Password, Department, Country) VALUES('"& strFirstName & "','"& strLastName & "','"&strExtention & "', '"&strEmail & "', '"& strPassword & "', '"& strDepartment & "', '"& strCountry & "');"
		
objConn.Execute(sql)

objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("Login.asp")
End if 

%>