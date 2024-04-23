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
Set FebCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#02/01/02# And (Calls.Date_Opened)<#02/28/02# )) OR (((Calls.Date_Closed)>#02/01/02# And (Calls.Date_Closed)<#02/28/02# )) ;")
Set FebHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#02/01/02# And (Calls.Date_Opened)<#02/28/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#02/01/02# And (Calls.Date_Closed)<#02/28/02#  And (Calls.Call_Type)='Help')) ;")
Set FebRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#02/01/02# And (Calls.Date_Opened)<#02/28/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#02/01/02# And (Calls.Date_Closed)<#02/28/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'February' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'February' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'February', Year= '"&ChartYear&"', TotalTickets= '" & FebCount(0) & "', HelpTickets= '" & FebHCount(0) & "', RequestTickets= '" & FebRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('February','"& ChartYear & "','"& FebCount(0) & "','"& FebHCount(0) & "','"& FebRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set MarCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#03/01/02# And (Calls.Date_Opened)<#03/31/02# )) OR (((Calls.Date_Closed)>#03/01/02# And (Calls.Date_Closed)<#03/31/02# )) ;")
Set MarHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#03/01/02# And (Calls.Date_Opened)<#03/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#03/01/02# And (Calls.Date_Closed)<#03/31/02#  And (Calls.Call_Type)='Help')) ;")
Set MarRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#03/01/02# And (Calls.Date_Opened)<#03/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#03/01/02# And (Calls.Date_Closed)<#03/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'March' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'March' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'March', Year= '"&ChartYear&"', TotalTickets= '" & MarCount(0) & "', HelpTickets= '" & MarHCount(0) & "', RequestTickets= '" & MarRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('March','"& ChartYear & "','"& MarCount(0) & "','"& MarHCount(0) & "','"& MarRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set AprCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#04/01/02# And (Calls.Date_Opened)<#04/30/02# )) OR (((Calls.Date_Closed)>#04/01/02# And (Calls.Date_Closed)<#04/30/02# )) ;")
Set AprHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#04/01/02# And (Calls.Date_Opened)<#04/30/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#04/01/02# And (Calls.Date_Closed)<#04/30/02#  And (Calls.Call_Type)='Help')) ;")
Set AprRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#04/01/02# And (Calls.Date_Opened)<#04/30/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#04/01/02# And (Calls.Date_Closed)<#04/30/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'April' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'April' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'April', Year= '"&ChartYear&"', TotalTickets= '" & AprCount(0) & "', HelpTickets= '" & AprHCount(0) & "', RequestTickets= '" & AprRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('April','"& ChartYear & "','"& AprCount(0) & "','"& AprHCount(0) & "','"& AprRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set MayCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#05/01/02# And (Calls.Date_Opened)<#05/31/02# )) OR (((Calls.Date_Closed)>#05/01/02# And (Calls.Date_Closed)<#05/31/02# )) ;")
Set MayHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#05/01/02# And (Calls.Date_Opened)<#05/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#05/01/02# And (Calls.Date_Closed)<#05/31/02#  And (Calls.Call_Type)='Help')) ;")
Set MayRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#05/01/02# And (Calls.Date_Opened)<#05/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#05/01/02# And (Calls.Date_Closed)<#05/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'May' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'May' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'May', Year= '"&ChartYear&"', TotalTickets= '" & MayCount(0) & "', HelpTickets= '" & MayHCount(0) & "', RequestTickets= '" & MayRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('May','"& ChartYear & "','"& MayCount(0) & "','"& MayHCount(0) & "','"& MayRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set JunCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#06/01/02# And (Calls.Date_Opened)<#06/30/02# )) OR (((Calls.Date_Closed)>#06/01/02# And (Calls.Date_Closed)<#06/30/02# )) ;")
Set JunHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#06/01/02# And (Calls.Date_Opened)<#06/30/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#06/01/02# And (Calls.Date_Closed)<#06/30/02#  And (Calls.Call_Type)='Help')) ;")
Set JunRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#06/01/02# And (Calls.Date_Opened)<#06/30/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#06/01/02# And (Calls.Date_Closed)<#06/30/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'June' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'June' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'June', Year= '"&ChartYear&"', TotalTickets= '" & JunCount(0) & "', HelpTickets= '" & JunHCount(0) & "', RequestTickets= '" & JunRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('June','"& ChartYear & "','"& JunCount(0) & "','"& JunHCount(0) & "','"& JunRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set JulCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#07/01/02# And (Calls.Date_Opened)<#07/31/02# )) OR (((Calls.Date_Closed)>#07/01/02# And (Calls.Date_Closed)<#07/31/02# )) ;")
Set JulHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#07/01/02# And (Calls.Date_Opened)<#07/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#07/01/02# And (Calls.Date_Closed)<#07/31/02#  And (Calls.Call_Type)='Help')) ;")
Set JulRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#07/01/02# And (Calls.Date_Opened)<#07/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#07/01/02# And (Calls.Date_Closed)<#07/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'July' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'July' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'July', Year= '"&ChartYear&"', TotalTickets= '" & JulCount(0) & "', HelpTickets= '" & JulHCount(0) & "', RequestTickets= '" & JulRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('July','"& ChartYear & "','"& JulCount(0) & "','"& JulHCount(0) & "','"& JulRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set AugCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#08/01/02# And (Calls.Date_Opened)<#08/31/02# )) OR (((Calls.Date_Closed)>#08/01/02# And (Calls.Date_Closed)<#08/31/02# )) ;")
Set AugHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#08/01/02# And (Calls.Date_Opened)<#08/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#08/01/02# And (Calls.Date_Closed)<#08/31/02#  And (Calls.Call_Type)='Help')) ;")
Set AugRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#08/01/02# And (Calls.Date_Opened)<#08/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#08/01/02# And (Calls.Date_Closed)<#08/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'August' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'August' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'August', Year= '"&ChartYear&"', TotalTickets= '" & AugCount(0) & "', HelpTickets= '" & AugHCount(0) & "', RequestTickets= '" & AugRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('August','"& ChartYear & "','"& AugCount(0) & "','"& AugHCount(0) & "','"& AugRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set SepCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#09/01/02# And (Calls.Date_Opened)<#09/30/02# )) OR (((Calls.Date_Closed)>#09/01/02# And (Calls.Date_Closed)<#09/30/02# )) ;")
Set SepHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#09/01/02# And (Calls.Date_Opened)<#09/30/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#09/01/02# And (Calls.Date_Closed)<#09/30/02#  And (Calls.Call_Type)='Help')) ;")
Set SepRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#09/01/02# And (Calls.Date_Opened)<#09/30/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#09/01/02# And (Calls.Date_Closed)<#09/30/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'September' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'September' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'September', Year= '"&ChartYear&"', TotalTickets= '" & SepCount(0) & "', HelpTickets= '" & SepHCount(0) & "', RequestTickets= '" & SepRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('September','"& ChartYear & "','"& SepCount(0) & "','"& SepHCount(0) & "','"& SepRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"


Set OctCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#10/01/02# And (Calls.Date_Opened)<#10/31/02# )) OR (((Calls.Date_Closed)>#10/01/02# And (Calls.Date_Closed)<#10/31/02# )) ;")
Set OctHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#10/01/02# And (Calls.Date_Opened)<#10/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#10/01/02# And (Calls.Date_Closed)<#10/31/02#  And (Calls.Call_Type)='Help')) ;")
Set OctRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#10/01/02# And (Calls.Date_Opened)<#10/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#10/01/02# And (Calls.Date_Closed)<#10/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'October' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'October' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'October', Year= '"&ChartYear&"', TotalTickets= '" & OctCount(0) & "', HelpTickets= '" & OctHCount(0) & "', RequestTickets= '" & OctRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('October','"& ChartYear & "','"& OctCount(0) & "','"& OctHCount(0) & "','"& OctRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set NovCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#11/01/02# And (Calls.Date_Opened)<#11/30/02# )) OR (((Calls.Date_Closed)>#11/01/02# And (Calls.Date_Closed)<#11/30/02# )) ;")
Set NovHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#11/01/02# And (Calls.Date_Opened)<#11/30/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#11/01/02# And (Calls.Date_Closed)<#11/30/02#  And (Calls.Call_Type)='Help')) ;")
Set NovRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#11/01/02# And (Calls.Date_Opened)<#11/30/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#11/01/02# And (Calls.Date_Closed)<#11/30/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'November' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'November' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'November', Year= '"&ChartYear&"', TotalTickets= '" & NovCount(0) & "', HelpTickets= '" & NovHCount(0) & "', RequestTickets= '" & NovRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('November','"& ChartYear & "','"& NovCount(0) & "','"& NovHCount(0) & "','"& NovRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set DecCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#12/01/02# And (Calls.Date_Opened)<#12/31/02# )) OR (((Calls.Date_Closed)>#12/01/02# And (Calls.Date_Closed)<#12/31/02# )) ;")
Set DecHCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#12/01/02# And (Calls.Date_Opened)<#12/31/02#  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#12/01/02# And (Calls.Date_Closed)<#12/31/02#  And (Calls.Call_Type)='Help')) ;")
Set DecRCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#12/01/02# And (Calls.Date_Opened)<#12/31/02#  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#12/01/02# And (Calls.Date_Closed)<#12/31/02#  And (Calls.Call_Type)='Request')) ;")
Set Count=objConn.Execute("SELECT Count(*) FROM ChartInfo WHERE Month = 'December' and Year='"&ChartYear&"' and HelpDesk='All'")
Set ObjRec=objConn.Execute("SELECT ID FROM ChartInfo WHERE Month = 'December' and Year='"&ChartYear&"' and HelpDesk='All'")
IF Count(0) >0 then
Count(0)
sql = "UPDATE ChartInfo SET  Month= 'December', Year= '"&ChartYear&"', TotalTickets= '" & DecCount(0) & "', HelpTickets= '" & DecHCount(0) & "', RequestTickets= '" & DecRCount(0) & "', HelpDesk= 'All' WHERE ID=" & ObjRec("ID") & ";"
	
objConn.Execute(sql)
Else
sql = "INSERT INTO ChartInfo(Month, Year, TotalTickets, HelpTickets, RequestTickets, HelpDesk) VALUES('December','"& ChartYear & "','"& DecCount(0) & "','"& DecHCount(0) & "','"& DecRCount(0) & "','All');"
		
objConn.Execute(sql)
End if
Count=0
objConn.Close
Set objConn = Nothing


%>