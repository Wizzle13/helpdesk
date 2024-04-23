<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 2-2503

Description:
This page is the Add Page for the Email Domains.
-->
<td valign="top">
<form method="post" action="send.asp" name="Inputform">
Domain:<br><input type="text" name="Domain" size="20" style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal"><p>
Country Abvr:<br><input type="text" name="Country" size="20" style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal"><p>
Country Full Name:<br><input type="text" name="CountryFull" size="20" style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal"><p>
Help Desk:<br><select name="Helpdesk">
<option style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal" value="America">North America
<option style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal" value="Atlantic">Atlantic Rim
<option style="FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal" value="Pacific">Pacific Rim
</select>
<br>
<br>
<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
</form>
