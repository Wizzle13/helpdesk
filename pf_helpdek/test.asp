 <%

If FormatDateTime (Time, vbShortTime) < "12:00" then
	Response.Write "Good Morning"
End if

If FormatDateTime (Time, vbShortTime) > "12:00" and FormatDateTime (Time, vbShortTime) < "17:00" then
	Response.Write "Good Afternoon"
End if

If FormatDateTime (Time, vbShortTime) > "17:00" then
	Response.Write "Good Evening"
End if

%>
