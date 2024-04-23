<%

strMonth = Request.Form("strMonth")
strYear = Request.Form("strYear")
strCount = Request.Form("strCount")
strRCount = Request.Form("strRCount")
strHCount = Request.Form("strHCount")
strHSOpenCount = Request.Form("strHSOpenCount")
strRSOpenCount = Request.Form("strRSOpenCount")
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "DSN=HelpDesk2"
objRec.Open "Stats", objConn
sql = "INSERT INTO Stats(Month,Year,Total_Tickets, Total_Help_Tickets, Total_Request_Tickets,Request_Tickets_Open,Help_Tickets_Open) VALUES('"& strMonth & "','"& strYear & "','"& strCount & "','"& strHCount & "','"& strRCount & "','"& strRSOpenCount & "','"& strHSOpenCount & "');"
objConn.Execute(sql)

	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing
%>