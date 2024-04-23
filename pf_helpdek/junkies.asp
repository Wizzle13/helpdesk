<!--#include file="_head.asp"-->
<%
Dim NameNew 
Dim NameOld 
Dim Name1 
Dim Name2 
Dim Name3 
Dim Name4 
Dim Name5 
Dim Name6 
Dim Name7 
Dim Name8 
Dim Name9 
Dim Name10
Dim Num1 
Dim Num2 
Dim Num3 
Dim Num4 
Dim Num5
Dim Num6 
Dim Num7 
Dim Num8 
Dim Num9 
Dim Num10

'Name1 = "" 
'Name2 = "" 
'Name3 = "" 
'Name4 = ""
'Name5 = ""
'Name6 = "" 
'Name7 = "" 
'Name8 = "" 
'Name9 = "" 
'Name10 = ""

Num1 = 0 
'Num2 = 0 
'Num3 = 0 
'Num4 = 0 
'Num5 = 0 
'Num6 = 0 
'Num7 = 0 
'Num8 = 0 
'Num9 = 0 
'Num10 = 0
strNameCount = 0
Response.Write "<td valign=top>"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from Calls ORDER by Email")
NameOld = objRs("Email")

	While Not objRs.EOF
	NameNew = objRs("Email")
If 	objRs("Email")<>"No_Longer_with_Company" and objRs("Email")<>"Retired" and objRs("Email")<>"SAP_Consultant" and objRs("Email")<>"N/A" and objRs("Email")<>"" then
	If NameNew = NameOld then
		strNameCount = strNameCount + 1
	DisName = objRs("User_First_Name") &" "& objRs("User_Last_Name")
	else
	strNameCount = strNameCount + 1
	IF (Num1 <= strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = Num5
			Num5 = Num4
			Num4 = Num3
			Num3 = Num2
			Num2 = Num1
			Num1 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = Name5
			Name5 = Name4
			Name4 = Name3
			Name3 = Name2
			Name2 = Name1
			Name1 = DisName
			
		End IF
		IF (Num2 <= strNameCount and Num1 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = Num5
			Num5 = Num4
			Num4 = Num3
			Num3 = Num2
			Num2 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = Name5
			Name5 = Name4
			Name4 = Name3
			Name3 = Name2
			Name2 = DisName
			
		End IF
		IF (Num3 <= strNameCount and Num2 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = Num5
			Num5 = Num4
			Num4 = Num3
			Num3 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = Name5
			Name5 = Name4
			Name4 = Name3
			Name3 = DisName
			
		End IF
		IF (Num4 <= strNameCount and Num3 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = Num5
			Num5 = Num4
			Num4 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = Name5
			Name5 = Name4
			Name4 = DisName

		End IF
		IF (Num5 <= strNameCount and Num4 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = Num5
			Num5 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = Name5
			Name5 = DisName

		End IF
		IF (Num6 <= strNameCount and Num5 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Num7 = Num6
			Num6 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = Name6
			Name6 = DisName
 
		End IF
		IF (Num7 <= strNameCount and Num6 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = Num7
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Num7 = strNameCount
			Name10 = Name9
			Name9 = Name8 
			Name8 = Name7
			Name7 = DisName
	
		End IF
		IF (Num8 <= strNameCount and Num7 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = Num8
			Num8 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = Name8 
			Name8 = DisName
			
		End IF
		IF (Num9 <= strNameCount and Num8 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = Num9
			Num9 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = Name9
			Name9 = DisName
			 
		End IF
		IF (Num10 <= strNameCount and Num9 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = Num10
			Num10 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = Name10
			Name10 = DisName
			
		End IF
		IF (Num11 <= strNameCount and Num10 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = Num11
			Num11 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = Name11
			Name11 = DisName
			
		End IF
		IF (Num12 <= strNameCount and Num11 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = Num12
			Num12 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = Name12
			Name12 = DisName
			
		End IF
		IF (Num13 <= strNameCount and Num12 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = Num13
			Num13 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = Name13
			Name13 = DisName
			
		End IF

		IF (Num14 <= strNameCount and Num13 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = Num14
			Num14 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = Name14
			Name14 = DisName
			
		End IF
		IF (Num15 <= strNameCount and Num14 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = Num15
			Num15 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = Name15
			Name15 = DisName
			
		End IF
		IF (Num16 <= strNameCount and Num15 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = Num16
			Num16 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = Name16
			Name16 = DisName
			
		End IF
		IF (Num17 <= strNameCount and Num16 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = Num17
			Num17 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = Name17
			Name17 = DisName
			
		End IF
		IF (Num18 <= strNameCount and Num17 > strNameCount) then
			Num20 = Num19
			Num19 = Num18
			Num18 = strNameCount
			Name20 = Name19
			Name19 = Name18
			Name18 = DisName
			
		End IF
		IF (Num19 <= strNameCount and Num18 > strNameCount) then
			Num20 = Num19
			Num19 = strNameCount
			Name20 = Name19
			Name19 = DisName
			
		End IF
		IF (Num20 <= strNameCount and Num19 > strNameCount) then
			Num20 =strNameCount
			Name20 = DisName
			
		End IF

		strNameCount = 0
		NameOld = NameNew
		
	End if	
End if		
		objRs.MoveNext
	Wend
Response.Write "<table border=1><tr>"
Response.Write "<td>1)</td><td>"& Name1 &"</td><td>"& Num1 & "</td></tr>"
Response.Write "<tr><td>2)</td><td>"& Name2 &"</td><td>"& Num2 & "</td></tr>"
Response.Write "<tr><td>3)</td><td>"& Name3 &"</td><td>"& Num3 & "</td></tr>"
Response.Write "<tr><td>4)</td><td>"& Name4 &"</td><td>"& Num4 & "</td></tr>"
Response.Write "<tr><td>5)</td><td>"& Name5 &"</td><td>"& Num5 & "</td></tr>"
Response.Write "<tr><td>6)</td><td>"& Name6 &"</td><td>"& Num6 & "</td></tr>"
Response.Write "<tr><td>7)</td><td>"& Name7 &"</td><td>"& Num7 & "</td></tr>"
Response.Write "<tr><td>8)</td><td>"& Name8 &"</td><td>"& Num8 & "</td></tr>"
Response.Write "<tr><td>9)</td><td>"& Name9 &"</td><td>"& Num9 & "</td></tr>"
Response.Write "<tr><td>10)</td><td>"& Name10 &"</td><td>"& Num10 & "</td>"
Response.Write "<tr><td>11)</td><td>"& Name11 &"</td><td>"& Num11 & "</td>"
Response.Write "<tr><td>12)</td><td>"& Name12 &"</td><td>"& Num12 & "</td>"
Response.Write "<tr><td>13)</td><td>"& Name13 &"</td><td>"& Num13 & "</td>"
Response.Write "<tr><td>14)</td><td>"& Name14 &"</td><td>"& Num14 & "</td>"
Response.Write "<tr><td>15)</td><td>"& Name15 &"</td><td>"& Num15 & "</td>"
Response.Write "<tr><td>16)</td><td>"& Name16 &"</td><td>"& Num16 & "</td>"
Response.Write "<tr><td>17)</td><td>"& Name17 &"</td><td>"& Num17 & "</td>"
Response.Write "<tr><td>18)</td><td>"& Name18 &"</td><td>"& Num18 & "</td>"
Response.Write "<tr><td>19)</td><td>"& Name19 &"</td><td>"& Num19 & "</td>"
Response.Write "<tr><td>20)</td><td>"& Name20 &"</td><td>"& Num20 & "</td>"
Response.Write "</tr></table>"
Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing%>

</BODY>
</HTML>
