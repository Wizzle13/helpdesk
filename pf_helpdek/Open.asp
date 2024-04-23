<!--#include file="_head.asp"-->
<%

Dim RefPage
RefPage = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
'Response.Write RefPage
%>
<script>
<!-- Script to open User selection window.
function remote(){
win2=window.open("userwindow.asp?sortby=LName","win2","width=450,height=281,titlebar=0,scrollbars=1")
win2.creator=self
}
function uploadfile(){
win3=window.open("uploads/index.asp?Num=<%=Request("Num")%>&RefPage=<%Response.Write RefPage%>","win3","width=375,height=95,titlebar=0,scrollbars=0")
win3.creator=self
}
//-->
</script>
<script language="JavaScript">
function Join_Form1_Validator(theForm)
{	
  if (theForm.FirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.FirstName.focus();
    return (false);
  }
  
  if (theForm.LastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.LastName.focus();
    return (false);
  }
  if (theForm.Extention.value == "")
  {
    alert("Please enter a value for the \"Extention\" field.");
    theForm.Extention.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.FirstName.value;
  var checkStr = theForm.LastName.value;
  var checkStr = theForm.Extention.value;

  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
   
  }
  
  return (true);
}
</script>

<td>
<form method="post" action="send.asp" name="Inputform" onsubmit='return Join_Form1_Validator(this)' name='Join_Form1'>


<table border="0">
<tr>
<td>Date:</td><td><input type="Hidden" name="Date" value=<%=FormatDateTime (Date, vbShortDate)%>><%=FormatDateTime (Date, vbShortDate)%></td>
</tr>
<tr>
<td>Time:</td><td><input type="Hidden" name="Time" value=<%=FormatDateTime (Time, vbShortTime)%>><%=FormatDateTime (Time, vbShortTime)%></td>
<input type="Hidden" name="Priority" value="2">
<input type="Hidden" name="SAP_Transport_Date" value="12/31/2049">
<input type="Hidden" name="tickpriority" value="99">
<input type="Hidden" name="Upgrade" value="off">
<input type="Hidden" name="Call_Serviced_By" value="">
<input type="Hidden" name="SAP_Module" value="N/A">
<input type=Hidden name=Projected_Complete_Date2 value='12/31/2049'>"		

<input type="Hidden" name="Country" value=<%=session("Country") %>>
</tr>
<tr>
<td>First Name:</td><td><input type="Hidden" name="FirstName" size="20" Value=<%=session("FirstName") %>><%=session("FirstName") %></td>
</tr>
<tr>
<td>Last Name:</td><td><input type="Hidden" name="LastName" size="20" Value=<%=session("LastName") %>><%=session("LastName") %></td>
</tr>
<tr>
<td>Contact Number:</td><td><input type="Hidden" name="Extention" size="20" Value= <%=session("Extention") %>><%=session("Extention") %></td>
</tr>
<%
Country=session("country")
Set objREC1=objConn.Execute("Select * from Domains WHERE Country = '" & Country & "'")
 %>
<tr>
<td>Email:</td><td><input type="Hidden" name="Logon_Name" size="20" Value= <%=session("Email") %>><%=session("Email")&objRec1("Domain")%></td>
</tr>
<tr>
<td>
<input type="Hidden" name="Package_Name" size="20" Value= "">
</td>
</tr>
<%
Response.Write "<tr><td><img src=images/1_circle.jpg></td><td colspan=2><font size=2><b>Step 1: If you want to attach a file to this ticket, do it next. If you have multiple files to upload, please zip them together and upload the zipped file.</b></font></td></tr>"

If Request("FileName") <> "" Then
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>File<br>Attachment:</td><td><a href='/uploads/data/" & Request("FileName") & "' target='new'>" & Request("FileName") & "</a>&nbsp;"
	Response.Write "<input type=hidden name='Upload_File_Location' value='/uploads/data/" & Request("FileName") & "'>"
Else
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>File<br>Attachment:</td><td>&nbsp;"
End If
Response.Write "<input type=button onClick=uploadfile() value='Attach File'></td></tr>"

Response.Write "<tr><td><img src=images/2_circle.jpg></td><td colspan=2><font size=2><b>Step 2: Complete the rest of the information and submit.</b></font></td></tr>"

%>
<tr>
<td>Call Type<font color=red>*</font>: </td><td>
	<select name="Call_Type">
		<option selected>HELP
		<option>REQUEST
		
	</select> <font color=red>(See Below)</font></td>
	
</tr>
</table>


Description: <font color=red>*Please be as specific as possible.</font><BR>
<textarea cols=50 rows=5 name="problem"></textarea><p>



<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
</form>

<font color=red>*</font>
<B>Help</B> - "It’s broken.  It used to work but now it doesn’t."<br>
&nbsp;&nbsp;&nbsp;<B>Request</B> - "I need something new, changed or updated."
</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>

</BODY>
</HTML>

