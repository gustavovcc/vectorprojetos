<% @Language = "VBScript" %>
<%
Option Explicit
On Error Resume Next
%>
<!--#include file="charts_common.asp"-->
<!--#include file="cSingleVerticalChart.asp"-->
<!--#include file="cMultipleVerticalChart.asp"-->
<!--#include file="cStackedVerticalChart.asp"-->
<!--#include file="cSingleHorizontalChart.asp"-->
<!--#include file="cMultipleHorizontalChart.asp"-->
<!--#include file="cStackedHorizontalChart.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>HTMLGraph - Web Charts without Components</title>
<meta name="copyright" content="www.u229.no &copy; 2002" />
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
<link rel="stylesheet" type="text/css" href="HTMLGraph.css" />
</head>
<body>

<div class="u229">u229</div>
<a href="http://www.u229.no/stuff/HTMLGraph/" class="navline">Application Home</a>|<a href="#" onmouseover="this.href='&#x6D;&#97;&#x69;&#108;&#x74;o&#x3A;' + 'postmaster' + '&#x40;' + 'u229.no'" class="navline">Email</a>

<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
	<td valign="top" width="100%">
	  <div class="main_content">
	  <h1 class="pagetitle">This is a live demo showing the different chart types you can create with HTMLGraph:</h1>

<%

Const IMAGES_PATH = "images/"
Const WIDTH_YGRAF_IMAGE = 36
Const HEIGHT_YGRAF_IMAGE = 111
Const adStateOpen = &H00000001
Const adCmdTableDirect = &H0200

Dim oChart1, oChart2, oChart3, oChart4, oChart5, oChart6
Dim i, bSuccess

Dim m_oRs
Dim m_sConnectionString
Dim SQLApples, SQLOranges, SQLPeaches
Dim m_arrApples, m_arrOranges, m_arrPeaches
Dim m_arrData()
Dim m_arrLabels(), m_arrLabels2, m_arrLabels3


'==========================================================
' Demonstrating 2 ways to add data to the charts: Dynamically from database or hard code an array of values.
'==========================================================

'// 1) FROM A DATABASE:

 '// Create some labels for the charts. (These could also have been taken from the database)
m_arrLabels2 = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
m_arrLabels3 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)

'// Define the connection string to database
m_sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("demo.mdb")

'// Build some simpel SQL
SQLApples = "SELECT Apples FROM Table1"
SQLOranges = "SELECT Oranges FROM Table1"
SQLPeaches = "SELECT Peaches FROM Table1"

'// Create our recordset object
Set m_oRs = Server.CreateObject("ADODB.Recordset")

'// Fill 3 Arrays with data from database using the ADO GetRows method. We will use these data arrays to create our charts
m_arrApples = GetRows(SQLApples)
m_arrOranges = GetRows(SQLOranges)
m_arrPeaches = GetRows(SQLPeaches)

'// Clean up
m_oRs.Close
Set m_oRs = Nothing


'// 2) PRODUCE SOME FAKE DATA:
Call LoadFakeData

'// Produce labels for the fake data
For i = 0 To UBound(m_arrData)
    Redim Preserve m_arrLabels(i)
	m_arrLabels(i) = i + 1
Next


'// THE FOLLOWING DEMONSTRATES HOW TO CREATE THE DIFFERENT CHARTS, AND WHAT
'// PROPERTIES EACH OF THEM EXPOSE, WITH THE DATA WE CREATED ABOVE:


'==========================================================
' Chart Type 1: Single Vertical Bar
'==========================================================
%>
<h3>Chart Type 1: Single Vertical Bar</h3>
<%

Set oChart1 = New cSingleVerticalChart

With oChart1
    .ChartTitle = "Page Hits May 2005"
	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarWidth = 12
	.BarImage = IMAGES_PATH & "1x1_83C4FE.gif"
	'// Print values for the bars: 0 = none (Default). 1 = Top of bar
	.BarValueMode = 1
    '// Where to print the totals: 0 = None (Default). 1 = Right of chart. 2 = Bottom
	.TotalsMode = 1
	'// If empty no yGraf image will be displayed
	.DisplayYGrafImage = IMAGES_PATH & "y100.gif"
	'// If empty no average bar will be displayed
	.DisplayAverageBar = IMAGES_PATH & "1x1navy.gif"
	'// Add text for the legend
	.AddLegendText = "Total"
	'// Shrink the data values with this factor
'	.ShrinkFactor = 10
    
	'// Pass only 1 data array
	If Not .AddData(m_arrData) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels for the bars
	If Not .AddLabels(m_arrLabels) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart1 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 1 = " & Err.Number & "<br>" & Err.Description

'==========================================================
' Chart Type 2: Multiple Vertical Bars
'==========================================================
%>
<h3>Chart Type 2: Multiple Vertical Bars</h3>
<%

Set oChart2 = New cMultipleVerticalChart

With oChart2
    .ChartTitle = "Page Hits 2002 - 2004"
	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarWidth = 12
	'// Where to print the totals: 0 = None (Default). 1 = Right of chart. 2 = Bottom
	.TotalsMode = 2
    '// If empty no yGraf image will be displayed
'	.DisplayYGrafImage = IMAGES_PATH & "y100.gif"
    '// Add images for the bars matching the number of bars/data arrays
	.AddBarImages = Array(IMAGES_PATH & "1x1_B0C4DD.gif", IMAGES_PATH & "1x1_FFC19F.gif", IMAGES_PATH & "1x1_CC9A7F.gif")
	'// Add text for the legend matching the number of bars/data arrays
	.AddLegendText = Array("2002", "2003", "2004")
	'// Shrink the data values with this factor
'	.ShrinkFactor = 10

    '// Add data to the chart. This chart supports 2 or 3 data arrays
	If Not .AddData(m_arrApples) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrOranges) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrPeaches) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels matching the number of data arrays passed above
	If Not .AddLabels(m_arrLabels2) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart2 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 2 = " & Err.Number & "<br>" & Err.Description

'==========================================================
' Chart Type 3: Stacked Vertical Bars
'==========================================================
%>
<h3>Chart Type 3: Stacked Vertical Bars</h3>
<%

Set oChart3 = New cStackedVerticalChart

With oChart3
    .ChartTitle = "Page Hits 2002 - 2004"
	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarWidth = 20
	'// Where to print the totals: 0 = None (Default). 1 = Right of chart. 2 = Bottom
	.TotalsMode = 2
	'// Print values for the bars: 0 = none (Default) 1 = Top of bar
	.BarValueMode = 1

    '// If empty no yGraf image will be displayed
'	.DisplayYGrafImage = IMAGES_PATH & "y100.gif"
    '// If empty no average bar will be displayed
'	.DisplayAverageBar = IMAGES_PATH & "1x1_837D10.gif"
    '// Add images for the bars matching the number of bars/data arrays
	.AddBarImages = Array(IMAGES_PATH & "1x1_F0ECBC.gif", IMAGES_PATH & "1x1_C2BE77.gif", IMAGES_PATH & "1x1_E4E3C5.gif")
	'// Add text for the legend matching the number of bars/data arrays
	.AddLegendText = Array("2002", "2003", "2004")
	'// Shrink the data values with this factor
	.ShrinkFactor = 2


    '// Add data to the chart. This chart supports 2 or 3 data arrays
	If Not .AddData(m_arrApples) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrOranges) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrPeaches) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels matching the number of data arrays passed above
	If Not .AddLabels(m_arrLabels3) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart3 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 3 = " & Err.Number & "<br>" & Err.Description


'==========================================================
' Chart Type 4: Single Horizontal Bar
'==========================================================
%>
<h3>Chart Type 4: Single Horizontal Bar</h3>
<%

Set oChart4 = New cSingleHorizontalChart

With oChart4
    .ChartTitle = "Page Hits May 2005"
	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarHeight = 12
	.BarImage = IMAGES_PATH & "1x1_B0C4DD.gif"
	'// Print values for the bars: 0 = none (Default) 1 = right of bar
	.BarValueMode = 1
    '// Where to print the totals: 0 = None (Default). 1 = Bottom
	.TotalsMode = 1
	'// If empty no average bar will be displayed
'	.DisplayAverageBar = IMAGES_PATH & "1x1navy.gif"
	'// Add text for the legend
	.AddLegendText = "Total"
	'// Shrink the data values with this factor
'	.ShrinkFactor = 30
    
	'// Pass 1 data array
	If Not .AddData(m_arrData) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels for the bars
	If Not .AddLabels(m_arrLabels) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart4 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 4 = " & Err.Number & "<br>" & Err.Description


'==========================================================
' Chart Type 5: Multiple Horizontal Bar
'==========================================================
%>
<h3>Chart Type 5: Multiple Horizontal Bar</h3>
<%

Set oChart5 = New cMultipleHorizontalChart

With oChart5
    .ChartTitle = "Page Hits 2003 - 2004"
	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarHeight = 12
	'// Print values for the bars: 0 = none (Default) 1 = right of bar
	.BarValueMode = 0
    '// Where to print the totals: 0 = None (Default). 1 = Bottom
	.TotalsMode = 1
	'// If empty no average bar will be displayed
'	.DisplayAverageBar = IMAGES_PATH & "1x1navy.gif"
    '// Add images for the bars matching the number of bars/data arrays
'	.AddBarImages = Array(IMAGES_PATH & "1x1_B0C4DD.gif", IMAGES_PATH & "1x1_FFC19F.gif", IMAGES_PATH & "1x1_CC9A7F.gif")
    .AddBarImages = Array(IMAGES_PATH & "1x1_B0C4DD.gif", IMAGES_PATH & "1x1_FFC19F.gif")
	'// Add text for the legend matching the number of bars/data arrays
'	.AddLegendText = Array("2002", "2003", "2004")
    .AddLegendText = Array("2003", "2004")
	'// Shrink the data values with this factor
'	.ShrinkFactor = 10

    '// Add data to the chart. This chart supports 2 or 3 data arrays
	If Not .AddData(m_arrApples) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrOranges) Then Response.Write "Error Message: " &  .ErrorMessage
'	If Not .AddData(m_arrPeaches) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels for the bars
	If Not .AddLabels(m_arrLabels2) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart5 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 5 = " & Err.Number & "<br>" & Err.Description


'==========================================================
' Chart Type 6: Stacked Horizontal Bar
'==========================================================
%>
<h3>Chart Type 6: Stacked Horizontal Bar</h3>
<%

Set oChart6 = New cStackedHorizontalChart

With oChart6
    .ChartTitle = "Page Hits 2002 - 2004"
'	.FooterText = "FooterText: Edit style in HTMLGraph.css"
	.TableCaption = "Listing unique visitors"
	.BarHeight = 16
	'// Print values for the bars: 0 = none (Default) 1 = right of bar
	.BarValueMode = 1
    '// Where to print the totals: 0 = None (Default). 1 = Bottom
	.TotalsMode = 1
	'// If empty no average bar will be displayed
'	.DisplayAverageBar = IMAGES_PATH & "1x1navy.gif"
    '// Add images for the bars matching the number of bars/data arrays
	.AddBarImages = Array(IMAGES_PATH & "1x1_B0C4DD.gif", IMAGES_PATH & "1x1_FFC19F.gif", IMAGES_PATH & "1x1_CC9A7F.gif")
	'// Add text for the legend matching the number of bars/data arrays
	.AddLegendText = Array("2002", "2003", "2004")
	'// Shrink the data values with this factor
'	.ShrinkFactor = 10

    '// Add data to the chart. This chart supports 2 or 3 data arrays
	If Not .AddData(m_arrApples) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrOranges) Then Response.Write "Error Message: " &  .ErrorMessage
	If Not .AddData(m_arrPeaches) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Add labels for the bars
	If Not .AddLabels(m_arrLabels2) Then Response.Write "Error Message: " &  .ErrorMessage
	'// Create the chart
    If Not .CreateChart Then Response.Write "Error Message: " &  .ErrorMessage
End With

Set oChart6 = Nothing

If Err Then Response.Write "<br>Err.Number in Chart 6 = " & Err.Number & "<br>" & Err.Description


%>
      </div>
    </td>
  </tr>
</table>

<div class="copy">Therefor: For that or this for it.<br />©&nbsp;www.u229.no</div>
<br />

</body>
</html>
<%

'------------------------------------------------------------------------------------------------------------
' Comment: Using GetRows - returns the data in an array.
'------------------------------------------------------------------------------------------------------------
Function GetRows(sSql)
    On Error Resume Next

	With m_oRs
	    If .State = adStateOpen Then .Close
		.Open sSql, m_sConnectionString, , , adCmdTableDirect
		GetRows = .GetRows
	End With

End Function

'------------------------------------------------------------------------------------------------------------
' Comment:  Let's produce some fake data.
'------------------------------------------------------------------------------------------------------------
Sub LoadFakeData()
    On Error Resume Next

	Redim m_arrData(30)

	m_arrData(0) = "18"
	m_arrData(1) = "4"
	m_arrData(2) = "0"
	m_arrData(3) = "16"
	m_arrData(4) = "70"
	m_arrData(5) = "75"
	m_arrData(6) = "72"
	m_arrData(7) = "90"
	m_arrData(8) = "65"
	m_arrData(9) = "59"
	m_arrData(10) = "59"
	m_arrData(11) = "1"
	m_arrData(12) = "6"
	m_arrData(13) = "44"
	m_arrData(14) = "56"
	m_arrData(15) = "24"
	m_arrData(16) = "21"
	m_arrData(17) = "22"
	m_arrData(18) = "88"
	m_arrData(19) = "12"
	m_arrData(20) = "74"
	m_arrData(21) = "41"
	m_arrData(22) = "52"
	m_arrData(23) = "33"
	m_arrData(24) = "19"
	m_arrData(25) = "81"
	m_arrData(26) = "65"
	m_arrData(27) = "25"
	m_arrData(28) = "12"
	m_arrData(29) = "21"
	m_arrData(30) = "3"


'// When working with BIG numbers remember to set a value for ShrinkFactor!
'	m_arrData(0) = "1822"
'	m_arrData(1) = "44"
'	m_arrData(2) = "0"
'	m_arrData(3) = "16332"
'	m_arrData(4) = "7022"
'	m_arrData(5) = "7510"
'	m_arrData(6) = "7200"
'	m_arrData(7) = "90666"
'	m_arrData(8) = "6514"
'	m_arrData(9) = "5900"
'	m_arrData(10) = "5903"
'	m_arrData(11) = "1"
'	m_arrData(12) = "65577"
'	m_arrData(13) = "44441"
'	m_arrData(14) = "56548"
'	m_arrData(15) = "24339"
'	m_arrData(16) = "21774"
'	m_arrData(17) = "22144"
'	m_arrData(18) = "88995"
'	m_arrData(19) = "12331"
'	m_arrData(20) = "742"
'	m_arrData(21) = "4188"
'	m_arrData(22) = "525"
'	m_arrData(23) = "33447"
'	m_arrData(24) = "1911"
'	m_arrData(25) = "81884"
'	m_arrData(26) = "65321"
'	m_arrData(27) = "25225"
'	m_arrData(28) = "12995"
'	m_arrData(29) = "23322"
'	m_arrData(30) = "37788"

End Sub
%>