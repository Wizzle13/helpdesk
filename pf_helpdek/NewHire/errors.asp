<HTML>
<HEAD>
  <TITLE>Pure Fishing Emplyment Hire/Change/Termination Error</TITLE>

<style>
<!--
body, td, center, p { font-family: arial; font-size: 10pt }
.small { font-size: 8pt; }
.welcome { font-family: arial; font-size: 12pt; color: ff0000; }
.large { font-size: 12pt; }
a:hover{font-weight:bold}
//-->
</style>

</HEAD>
<BODY  link=blue vlink=blue alink=blue>

<a href="/"><img src="/images/purefishing2.gif" width=100 border=0 align=left></A><P>

<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<%
If Request("Message") = "1" Then
	Response.Write "<p><center><font size=4 color=red>Some information was missing from your form.<P>Please click the 'BACK' button in your browser window and try again.<P><HR><P></font></center>"
End If
If Request("Message") = "2" Then
	Response.Write "<p><center><font size=4 color=red>Your information has been successfully submitted.</font><p><font color=black>Someone from IT will contact you shortly to verify the information you have provided.<P>Please click the Pure Fishing logo to the left to return to the Intranet, or wait and you will be forwarded there automatically.</font><P><HR><P><font color=black><a href='addindex.asp'>New/Update Employee Form</a> | <a href='termindex.asp'>Termination of Employment Form</a></font></center>"
End If
If Request("Message") = "3" Then
	Response.Write "<p><center><font size=4 color=red>'Expected Start Date' must be later than " + FormatDateTime (Date, vbShortDate) + ".<P>Please click the 'BACK' button in your browser window, enter the correct start date and click 'SEND'.<P><HR><P></font></center>"
End If
%>


</BODY>
</HTML>

