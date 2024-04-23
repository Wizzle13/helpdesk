<%
Chartyear="2002"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
'All Help Desk Tickets
Set JanCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02# )) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02# )) ;")
Set JanHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Help')) ;")
Set JanRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'January', Year= '"&ChartYear&"', TotalTickets= '" & JanCount(0) & "', HelpTickets= '" & JanHCount(0) & "', RequestTickets= '" & JanRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('January','"& ChartYear & "','"& JanCount(0) & "','"& JanHCount(0) & "','"& JanRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
'US Help Desk Tickets
Set JanCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02# )) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02# ))) AND (Country='US' Or Country='CA') ;")
Set JanHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Help'))) AND (Country='US' Or Country='CA') ;") 
Set JanRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Request'))) AND (Country='US' Or Country='CA') ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='US'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='US'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'January', Year= '"&ChartYear&"', TotalTickets= '" & JanCount(0) & "', HelpTickets= '" & JanHCount(0) & "', RequestTickets= '" & JanRCount(0) & "', HelpDesk= 'US' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('January','"& ChartYear & "','"& JanCount(0) & "','"& JanHCount(0) & "','"& JanRCount(0) & "','US');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing


Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
'Sweden Help Desk Tickets
Set JanCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02# )) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02# ))) AND (Country='SE' Or Country='DK' Or Country='FI' Or Country='FR' Or Country='UK' Or Country='NO') ;")
Set JanHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Help'))) AND (Country='SE' Or Country='DK' Or Country='FI' Or Country='FR' Or Country='UK' Or Country='NO') ;")
Set JanRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Request'))) AND (Country='SE' Or Country='DK' Or Country='FI' Or Country='FR' Or Country='UK' Or Country='NO') ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='SE'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='SE'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'January', Year= '"&ChartYear&"', TotalTickets= '" & JanCount(0) & "', HelpTickets= '" & JanHCount(0) & "', RequestTickets= '" & JanRCount(0) & "', HelpDesk= 'SE' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('January','"& ChartYear & "','"& JanCount(0) & "','"& JanHCount(0) & "','"& JanRCount(0) & "','SE');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
'Taiwan Help Desk Tickets
Set JanCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02# )) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02# ))) AND (Country='TW' Or Country='NZ' Or Country='JP' Or Country='AU') ;")
Set JanHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Help'))) AND (Country='TW' Or Country='NZ' Or Country='JP' Or Country='AU') ;")
Set JanRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>#01/01/02# And (Calls.Date_Opened)<#01/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#01/01/02# And (Calls.Date_Closed)<#01/31/02#  And (Calls.Call_Type)='Request'))) AND (Country='TW' Or Country='NZ' Or Country='JP' Or Country='AU') ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='TW'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'January' and Year='"&ChartYear&"' and HelpDesk='TW'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'January', Year= '"&ChartYear&"', TotalTickets= '" & JanCount(0) & "', HelpTickets= '" & JanHCount(0) & "', RequestTickets= '" & JanRCount(0) & "', HelpDesk= 'TW' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('January','"& ChartYear & "','"& JanCount(0) & "','"& JanHCount(0) & "','"& JanRCount(0) & "','TW');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing


%>