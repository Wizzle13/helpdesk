<%

strDomain = Replace(Request.Form("Domain"), "'", "''")
strCountry = Replace(Request.Form("Country"), "'", "''")
strCountryFull = Replace(Request.Form("CountryFull"), "'", "''")
strHelpDesk = Replace(Request.Form("HelpDesk"), "'", "''")
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "DSN=Helpdesk2"
objRec.Open "Domains", objConn

sql = "INSERT INTO Domains(Domain, Country, CountryFull, HelpDesk) VALUES('"& strDomain &"','"& strCountry &"','"& strCountryFull &"','"& strHelpDesk & "');"
		
objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Response.Redirect ("index.asp")

%>