<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 7-1-02

Description:
This page is the Add Page for the FAQs.
-->
<td valign="top">
<form method="post" action="send.asp" name="Inputform">
<table>
<tr><td>
Category:</td><td><select name="Category">
	<option value=HW selected>Hardware
	<option value=SW>Software
	</select>
</td></tr>
<tr><td>
Group Name:</td><td>
<select name="Group">
<option value="Accounting Package">Accounting Package
<option value="Communications">Communications
<option value="Copier">Copier
<option value="Desktop Productivity">Desktop Productivity
<option value="Desktop Publishing">Desktop Publishing
<option value="Engineering Package">Engineering Package
<option value="Fax">Fax
<option value="Forecasting Package">Forecasting Package
<option value="Human Resources Package">Human Resources Package
<option value="Manufacturing Package">Manufacturing Package
<option value="Network">Network
<option value="Operating System">Operating System
<option value="PCs">PCs
<option value="Phones">Phones
<option value="Printers">Printers
<option value="Server">Server
<option value="Shipping System">Shipping System
<option value="System Administration">System Administration
<option value="User Error">User Error
</select>
</td></tr>
<tr><td>
Package Name:</td><td>
<input name="Package">
</td></tr>
<tr><td colspan=2>
<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
</form>
</td></tr></table>
