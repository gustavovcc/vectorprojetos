<%
'============================================================
' MODULE:        cStackedHorizontalChart.asp
' AUTHOR:        www.u229.no
' CREATED:      June 2005
' HOME PAGE:   http://www.u229.no/stuff/HTMLGraph/
'============================================================
' COMMENT: This class will produce horizontal stacked bar charts using Classic ASP.
'                  There are 5 more classes producing different types of charts.
'============================================================
' ROUTINES:
' - Private Sub Class_Initialize()
' - Private Sub Class_Terminate()
' - Public Function CreateChart()
' - Public Sub AddData(arrData)
' - Public Sub AddLabels(arrLabels)
'============================================================


Class cStackedHorizontalChart

'// MODULE VARIABLES:
Private m_sChartTitle                          '// Title of chart
Private m_sFooterText                         '// Footer text
Private m_sTableCaption                     '// Text for the table caption
Private m_lngBarImageHeight              '// Set width for the bar image in pixels
Private m_lngAverage                         '// Average value of dataset
Private m_lngBarValueMode                '// Where to print the value for the bar. 0 = None (Default) 1 = right of bar
Private m_sErrorMsg                           '// Return human readable error message to user if something goes wrong
Private m_arrData1                             '// Array holding the data passed to us
Private m_arrData2                             '// Array holding the data passed to us
Private m_arrData3                             '// Array holding the data passed to us
Private m_lngArrayCounter                  '// Number of data arrays passed
Private m_arrLabels                            '// Print labels for the bars
Private m_arrBarImages                     '// What images to use for the bars?
Private m_sDisplayAverageBar            '// Display an image for the average bar. If empty, no average bar is shown
Private m_lngTotalsMode                     '// 0 = None (Default). 1 = Right of chart. 2 = Bottom
Private m_lngShrinkFactor                   '// Divide all data values with this factor
Private m_arrLegendText                    '// Array holding text for the legend totals

'// MODULE PROPERTIES:
Public Property Let ChartTitle(s)
    m_sChartTitle = s
End Property
Public Property Let FooterText(s)
    m_sFooterText = s
End Property
Public Property Let TableCaption(s)
    m_sTableCaption = s
End Property
Public Property Let BarHeight(i)
    On Error Resume Next

	If Len(i) = 0 Or Not IsNumeric(i) Then m_lngBarImageHeight = 14: Exit Property
    m_lngBarImageHeight = Clng(i)
End Property
Public Property Let LegendMode(i)
    On Error Resume Next

	If Len(i) = 0 Or Not IsNumeric(i) Then m_lngLegendMode = 0: Exit Property
    m_lngLegendMode = CLng(i)
End Property
Public Property Let AddBarImages(arr)
    m_arrBarImages = arr
End Property
Public Property Let BarValueMode(i)            '// 0=None (Default) 1=Top 2= Bottom
    On Error Resume Next

    If Len(i) = 0 Or Not IsNumeric(i) Then m_lngBarValueMode = 0: Exit Property
    m_lngBarValueMode = CLng(i)
End Property
Public Property Let DisplayAverageBar(s)
    m_sDisplayAverageBar = s
End Property
Public Property Let TotalsMode(i)
    On Error Resume Next

    If Len(i) = 0 Or Not IsNumeric(i) Then m_lngTotalsMode = 0: Exit Property

    m_lngTotalsMode = CLng(i)
End Property
Public Property Let ShrinkFactor(i)
    On Error Resume Next

    If Len(i) = 0 Or Not IsNumeric(i) Then m_lngShrinkFactor = 1: Exit Property

    m_lngShrinkFactor = CLng(i)
End Property
Public Property Let AddLegendText(arr)
    m_arrLegendText = arr
End Property
Public Property Get ErrorMessage()
    ErrorMessage = m_sErrorMsg
End Property


'------------------------------------------------------------------------------------------------------------
' Comment: Init our module variables
'------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error Resume Next
    
	m_sChartTitle = ""
	m_sFooterText = ""
	m_sTableCaption = ""
	m_lngBarImageHeight = 14
	m_lngLegendMode = 0
	m_lngBarValueMode = 0
	m_sDisplayAverageBar = ""
	m_lngAverage = 0
	m_lngTotalsMode = 0
	m_lngShrinkFactor = 1
	m_sErrorMsg = ""

End Sub


'--------------------------------------------------------------------------------------------------------
' Comment: Clean up.
'--------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub


'------------------------------------------------------------------------------------------------------------
' Comment: Print the chart.
'------------------------------------------------------------------------------------------------------------
Public Function CreateChart()
    On Error Resume Next

	Dim i, iMax, iTmp1, iTmp2, iTmp3, iShrinked1, iShrinked2, iShrinked3
	Dim lngTotal1, lngTotal2, lngTotal3, lngTotal1_2, lngTotal1_3 

'---------------------------- Validate user input

	If (m_lngArrayCounter < 2 Or m_lngArrayCounter > 3) Then m_sErrorMsg = "Only 2 or 3 data arrays allowed for this chart": Exit Function
	If Not IsSafeArray(m_arrData1) Then m_sErrorMsg = "Invalid data": Exit Function
	iMax = UBound(m_arrData1)
	If iMax <> UBound(m_arrData2) Then m_sErrorMsg = "Data arrays don't have same size": Exit Function
	If m_lngArrayCounter = 3 Then
	    If iMax <> UBound(m_arrData3) Then m_sErrorMsg = "Data arrays don't have same size": Exit Function
	End If
	If Not iMax = UBound(m_arrLabels) Then m_sErrorMsg = "Labels array don't have correct size": Exit Function
	If m_lngTotalsMode = 1 And Not IsSafeArray(m_arrLegendText) Then m_sErrorMsg = "Invalid data for the Legend Text": Exit Function
	If (m_lngArrayCounter - 1) <> UBound(m_arrBarImages) Then m_sErrorMsg = "# of images for the bars don't mach number of data arrays": Exit Function

	'// Start printing the chart
    Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width="""" id=""chart6_table"">"
	If Len(m_sTableCaption) > 0 Then Response.Write "<caption>" & m_sTableCaption  & "</caption>"
	Response.Write "<tr><td colspan= ""2"" class=""chart6_title"">" & m_sChartTitle & "</td></tr>"
	Response.Write "<tr><td colspan= ""2"" class=""chart6_spacer_top"">&nbsp;</td></tr>"

'---------------------------- Print data & labels	

    For i = 0 To iMax
        
		'// We should have 2 or 3 data arrays
		Select Case m_lngArrayCounter

			Case 2
				iTmp1 = CLng(m_arrData1(i))
				iTmp2 = CLng(m_arrData2(i))

				iShrinked1 = iTmp1
				iShrinked2 = iTmp2

				If m_lngShrinkFactor > 1 Then
				    iShrinked1 = DataShrink(iTmp1, m_lngShrinkFactor)
					iShrinked2 = DataShrink(iTmp2, m_lngShrinkFactor)
				End If

				lngTotal1 = (lngTotal1 + iTmp1)
				lngTotal2 = (lngTotal2 + iTmp2)
				lngTotal1_2 = (iTmp1 + iTmp2)
				
				Response.Write "<tr><td class=""chart6_labels"">" & m_arrLabels(i) & "</td>"
				Response.Write "<td class=""chart6_barcell"">"
				
				Response.Write "<div class=""w""><img src=""" & m_arrBarImages(0) & """ width=""" & _
					iShrinked1 & """ height=""" & m_lngBarImageHeight & """ alt="""" title=""" & iTmp1 & """ />"
				Response.Write "<img src=""" & m_arrBarImages(1) & """ width=""" & iShrinked2 & """ height=""" & _
				    m_lngBarImageHeight & """ alt="""" title=""" & iTmp2 & """ />"
				If m_lngBarValueMode = 1 Then Response.Write "&nbsp;" & lngTotal1_2
				Response.Write "</div></td></tr>"

			Case 3
				iTmp1 = CLng(m_arrData1(i))
				iTmp2 = CLng(m_arrData2(i))
				iTmp3 = CLng(m_arrData3(i))

				iShrinked1 = iTmp1
				iShrinked2 = iTmp2
				iShrinked3 = iTmp3

				If m_lngShrinkFactor > 1 Then
				    iShrinked1 = DataShrink(iTmp1, m_lngShrinkFactor)
					iShrinked2 = DataShrink(iTmp2, m_lngShrinkFactor)
					iShrinked3 = DataShrink(iTmp3, m_lngShrinkFactor)
				End If

				lngTotal1 = (lngTotal1 + iTmp1)
				lngTotal2 = (lngTotal2 + iTmp2)
				lngTotal3 = (lngTotal3 + iTmp3)
				lngTotal1_3 = (iTmp1 + iTmp2 + iTmp3)

				Response.Write "<tr><td class=""chart6_labels"">" & m_arrLabels(i) & "</td>"
				Response.Write "<td class=""chart6_barcell"">"
				
				Response.Write "<div class=""container_bars""><img src=""" & m_arrBarImages(0) & """ width=""" & _
					iShrinked1 & """ height=""" & m_lngBarImageHeight & """ alt="""" title=""" & iTmp1 & """ />"
				Response.Write "<img src=""" & m_arrBarImages(1) & """ width=""" & iShrinked2 & """ height=""" & _
				    m_lngBarImageHeight & """ alt="""" title=""" & iTmp2 & """ />"
				Response.Write "<img src=""" & m_arrBarImages(2) & """ width=""" & iShrinked3 & """ height=""" & _
				    m_lngBarImageHeight & """ alt="""" title=""" & iTmp3 & """ />"
				If m_lngBarValueMode = 1 Then Response.Write "&nbsp;" & lngTotal1_3

				Response.Write "</div></td></tr>"

			Case Else
			    m_sErrorMsg = "Only 2 or 3 data arrays are allowed for this chart type": Exit Function

		End Select

	Next

'---------------------------- Print Totals at the bottom of the chart?

    If m_lngTotalsMode = 1 Then

	    Select Case m_lngArrayCounter
	        Case 2
			    Response.Write "<tr><td colspan=""3"" class=""chart6_totals_bottom"">" & _
				    "<img src=""" & m_arrBarImages(0) & """ width=""" & _
				    m_lngBarImageHeight & """ height=""" & m_lngBarImageHeight & """ alt=""""  />&nbsp;" & m_arrLegendText(0) & " = " & lngTotal1 & _
				    "&nbsp;&nbsp;&nbsp;<img src=""" & m_arrBarImages(1) & """ width=""" & _
				    m_lngBarImageHeight & """ height=""" & m_lngBarImageHeight & """ alt=""""  />&nbsp;" & m_arrLegendText(1) & " = " & lngTotal2 & "</td></tr>"

			Case 3
			    Response.Write "<tr><td colspan=""3"" class=""chart6_totals_bottom"">" & _
				    "<img src=""" & m_arrBarImages(0) & """ width=""" & _
				    m_lngBarImageHeight & """ height=""" & m_lngBarImageHeight & """ alt=""""  />&nbsp;" & m_arrLegendText(0) & " = " & lngTotal1 & _
				    "&nbsp;&nbsp;&nbsp;<img src=""" & m_arrBarImages(1) & """ width=""" & _
				    m_lngBarImageHeight & """ height=""" & m_lngBarImageHeight & """ alt=""""  />&nbsp;" & m_arrLegendText(1) & " = " & lngTotal2 & _
				    "&nbsp;&nbsp;&nbsp;<img src=""" & m_arrBarImages(2) & """ width=""" & _
				    m_lngBarImageHeight & """ height=""" & m_lngBarImageHeight & """ alt=""""  />&nbsp;" & m_arrLegendText(2) & " = " & lngTotal3 & "</td></tr>"

	        Case Else
			    m_sErrorMsg = "Only 2 or 3 data arrays are allowed for this chart type": Exit Function

	    End Select

	End If

'---------------------------- Should we print a footer text?
    If Len(m_sFooterText) > 0 Then Response.Write "<tr><td colspan=""3"" class=""chart6_footer"">" & m_sFooterText & "</td></tr>"

'---------------------------- Finish the table and return a boolean
	Response.Write "</table>"

    CreateChart = (Err.Number = 0)

End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Add data array to the chart & check that arrray have valid dimensions.
'------------------------------------------------------------------------------------------------------------
Public Function AddData(arrData)
    On Error Resume Next

	Dim lDimension

	If Not IsSafeArray(arrData) Then m_sErrorMsg = "No valid data": Exit Function

    '// Check the dimension of the passed data array. We accept only 1 or 2 dimensional arrays for the data.
	lDimension = CLng(ReturnArrayDimensions(arrData))                              

	If lDimension < 1 Or lDimension > 2 Then m_sErrorMsg = "Only 1 or 2 dimensional data arrays are supported": Exit Function

	'// Reduce the dimension from 2 to 1?
	If lDimension = 2 Then arrData = ShrinkDimension(arrData)

	'// Count number of passed data arrays. We allow max 3
	m_lngArrayCounter = (CLng(m_lngArrayCounter) + 1)

	Select Case m_lngArrayCounter
	    Case 1
		    m_arrData1 = arrData
		Case 2
		    m_arrData2 = arrData
		Case 3
		    m_arrData3 = arrData
	    Case Else
		    m_sErrorMsg = "Maximum 3 data arrays are allowed": Exit Function
	End Select

	AddData = (Err.Number = 0)

End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Add labels to print with the bars.
'------------------------------------------------------------------------------------------------------------
Public Function AddLabels(arrLabels)
    On Error Resume Next

	Dim lDimension

	If Not IsSafeArray(arrLabels) Then m_sErrorMsg = "No valid data for the labels": Exit Function
    
	'// Check the dimension of the passed array. We accept only 1 dimensional arrays for the labels.
	lDimension = CLng(ReturnArrayDimensions(arrLabels))

	If lDimension <> 1 Then m_sErrorMsg = "Only 1 dimensional arrays are supported for the labels": Exit Function

	m_arrLabels = arrLabels

	AddLabels = (Err.Number = 0)

End Function


End Class
%>