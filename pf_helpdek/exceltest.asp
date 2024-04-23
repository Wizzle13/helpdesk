<%@ LANGUAGE="VBSCRIPT" %>
<%
' Create Object 
Set MyExcelChart = CreateObject("Excel.Sheet")

' show or dont show excel to user, TRUE or FALSE
MyExcelChart.Application.Visible = True

' populate the cells 
MyExcelChart.ActiveSheet.Range("B2:k2").Value = Array("Week1", "Week2", "Week3", "Week4", "Week5", "Week6", "Week7", "Week8", "Week9", "Week10")
MyExcelChart.ActiveSheet.Range("B3:k3").Value = Array("67", "87", "5", "9", "7", "45", "45", "54", "54", "10")
MyExcelChart.ActiveSheet.Range("B4:k4").Value = Array("10", "10", "8", "27", "33", "37", "50", "54", "10", "10")
MyExcelChart.ActiveSheet.Range("B5:k5").Value = Array("23", "3", "86", "64", "60", "18", "5", "1", "36", "80")
MyExcelChart.ActiveSheet.Cells(3,1).Value="Internet Explorer"
MyExcelChart.ActiveSheet.Cells(4,1).Value="Netscape"
MyExcelChart.ActiveSheet.Cells(5,1).Value="Other"

' Select the contents that need to be in the chart
MyExcelChart.ActiveSheet.Range("b2:k5").Select
    
' Add the chart
MyExcelChart.Charts.Add
' Format the chart, set type of chart, shape of the bars, show title, get the data for the chart, show datatable, show legend 
MyExcelChart.activechart.ChartType = 97
MyExcelChart.activechart.BarShape =3
MyExcelChart.activechart.HasTitle = True
MyExcelChart.activechart.ChartTitle.Text = "Visitors log for each week shown in browsers percentage"
MyExcelChart.activechart.SetSourceData MyExcelChart.Sheets("Sheet1").Range("A1:k5"),1
MyExcelChart.activechart.Location 1
MyExcelChart.activechart.HasDataTable = True
MyExcelChart.activechart.DataTable.ShowLegendKey = True


' Save the the excelsheet to chart.xls
MyExcelChart.SaveAs "c:\chart.xls"


%>
<HTML>
<HEAD>
<TITLE>MyExcelChart</TITLE>
</HEAD>
<BODY>
</BODY>
</HTML>