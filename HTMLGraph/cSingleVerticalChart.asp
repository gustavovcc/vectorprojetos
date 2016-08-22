<%
'============================================================
' MODULE:        cSingleVerticalChart.asp
' AUTHOR:        www.u229.no
' CREATED:      June 2005
' HOME PAGE:   http://www.u229.no/stuff/HTMLGraph/
'============================================================
' COMMENT: This class will produce vertical single bar charts using Classic ASP.
'                  There are 5 more classes producing different types of charts.
'============================================================
' ROUTINES:
' - Private Sub Class_Initialize()
' - Private Sub Class_Terminate()
' - Public Function CreateChart()
' - Public Sub AddData(arrData)
' - Public Sub AddLabels(arrLabels)
'============================================================


Class cSingleVerticalChart

'// MODULE VARIABLES:
Private m_sChartTitle                          '// Title of chart
Private m_sFooterText                         '// Footer text
Private m_sTableCaption                     '// Text for the table caption
Private m_lngBarImageWidth               '// Set width for the bar image in pixels
Private m_lngAverage                         '// Average value of dataset
Private m_lngBarValueMode                '// Where to print the value for the bar. 0 = None (Default) 1 = top
Private m_sErrorMsg                           '// Return human readable error message to user if something goes wrong
Private m_arrData                              '// Array holding the data passed to us
Private m_arrLabels                            '// Print labels for the bars
Private m_sBarImage                         '// What image to use for the bar?
Private m_sDisplayYGrafImage            '// Display an image for the Y-graf. If this is empty, no image will be shown
Private m_sDisplayAverageBar            '// Display an image for the average bar. If empty, no average bar is shown
Private m_lngTotalsMode                     '// 0 = None (Default). 1 = Right of chart. 2 = Bottom
Private m_lngTotal                              '// Total sum of all bars
Private m_lngShrinkFactor                   '// Divide all data values with this factor
Private m_sLegendText                        '// Text for the legend


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
Public Property Let BarWidth(i)
    On Error Resume Next

	If Len(i) = 0 Or Not IsNumeric(i) Then m_lngBarImageWidth = 14: Exit Property
    m_lngBarImageWidth = Clng(i)
End Property
Public Property Let LegendMode(i)
    On Error Resume Next

	If Len(i) = 0 Or Not IsNumeric(i) Then m_lngLegendMode = 0: Exit Property
    m_lngLegendMode = CLng(i)
End Property
Public Property Let BarValueMode(i)            '// 0=None (Default) 1=Top 2= Bottom
    On Error Resume Next

    If Len(i) = 0 Or Not IsNumeric(i) Then m_lngBarValueMode = 0: Exit Property
    m_lngBarValueMode = CLng(i)
End Property
Public Property Let BarImage(s)
    m_sBarImage = s
End Property
Public Property Let DisplayYGrafImage(s)
    m_sDisplayYGrafImage = s
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
Public Property Let AddLegendText(s)
    m_sLegendText = s
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
	m_lngBarImageWidth = 14
	m_lngLegendMode = 0
	m_lngBarValueMode = 0
	m_sBarImage = ""
	m_sDisplayYGrafImage = ""
	m_sDisplayAverageBar = ""
	m_lngAverage = 0
	m_lngTotalsMode = 0
	m_lngTotal = 0
	m_lngShrinkFactor = 1
	m_sLegendText = "Total"
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

	Dim i, iMax, iTmp, iShrinked

'---------------------------- Validate user input
	If Not IsSafeArray(m_arrData) Then m_sErrorMsg = "Invalid data": Exit Function
	iMax = UBound(m_arrData)
	If iMax <> UBound(m_arrLabels) Then m_sErrorMsg = "# Labels don't match the size of the data": Exit Function

	'// Calculate the average value for this data set
	m_lngAverage = ReturnAverage

	'// Start printing the chart
    Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width="""" id=""chart1_table"">"
	If Len(m_sTableCaption) > 0 Then Response.Write "<caption>" & m_sTableCaption  & "</caption>"
	Response.Write "<tr><td colspan= """ & (iMax + 4) & """ class=""chart1_title"">" & m_sChartTitle & "</td></tr>"

'---------------------------- Should we print an image for the Y graf?

    If Len(m_sDisplayYGrafImage) > 0 Then
	    Response.Write "<tr><td valign=""bottom""><img src=""" & m_sDisplayYGrafImage & """ width=""" & _
		        WIDTH_YGRAF_IMAGE & """ height=""" & HEIGHT_YGRAF_IMAGE & """ alt="""" /></td>"
	Else
	    Response.Write "<tr><td>&nbsp;</td>"
	End If

'---------------------------- Should we print an image for the average bar?

    If Len(m_sDisplayAverageBar) > 0 Then
	    iShrinked = m_lngAverage
		If m_lngShrinkFactor > 1 Then iShrinked = DataShrink(m_lngAverage, m_lngShrinkFactor)

	    Response.Write "<td valign=""bottom"" class=""chart1_barcell"">" & m_lngAverage & "<br /><img src=""" & _
		        m_sDisplayAverageBar & """ width=""" & m_lngBarImageWidth & """ height=""" & _
		        iShrinked & """ alt=""Average"" title=""Average: " & m_lngAverage & """ /></td>"
	Else
	    Response.Write "<td valign=""bottom"" class=""chart1_barcell"">&nbsp;</td>"
	End If
	
'---------------------------- Where should we print the values for each bar?
    Select Case m_lngBarValueMode

        Case 0    '// Don't print bar values
			For i = 0 To iMax
				iTmp = m_arrData(i)
				iShrinked = iTmp
				If m_lngShrinkFactor > 1 Then iShrinked = DataShrink(iTmp, m_lngShrinkFactor)
				m_lngTotal = (m_lngTotal + iTmp)
				Response.Write "<td valign=""bottom"" class=""chart1_barcell""><img src=""" & m_sBarImage & """ width=""" & _
				        m_lngBarImageWidth & """ height=""" & iShrinked & """ alt="""" title=""" & iTmp & """ /></td>"
			Next

		Case 1    '// Print bar values above the bars
			For i = 0 To iMax
			    iTmp = m_arrData(i)
				iShrinked = iTmp
				If m_lngShrinkFactor > 1 Then iShrinked = DataShrink(iTmp, m_lngShrinkFactor)
				m_lngTotal = (m_lngTotal + iTmp)
			    Response.Write "<td valign=""bottom"" class=""chart1_barcell"">" & iTmp & "<br /><img src=""" & m_sBarImage & """ width=""" & _
				        m_lngBarImageWidth & """ height=""" & iShrinked & """ alt="""" title=""" & iTmp & """ /></td>"
			Next

        Case Else
    End Select

'---------------------------- Print Totals to the right of the chart?

    If m_lngTotalsMode = 1 Then
	    Response.Write "<td rowspan=""2"" class=""chart1_totals_right""><img src=""" & m_sBarImage & """ width=""" & _
				        m_lngBarImageWidth & """ height=""" & m_lngBarImageWidth & """ alt=""""  />&nbsp;" & m_sLegendText & " = " & m_lngTotal & "<br />"
						
		'// Print the average info?
		If Len(m_sDisplayAverageBar) > 0 Then
		    Response.Write "<img src=""" & m_sDisplayAverageBar & """ width=""" & m_lngBarImageWidth & """ height=""" & _
		        m_lngBarImageWidth & """ alt="""" />&nbsp;Average = " & m_lngAverage
		End If

		Response.Write "</td></tr>"

	Else
	    Response.Write "<td rowspan=""2"" class=""chart1_totals_right"">&nbsp;</td></tr>"
	End If

'---------------------------- Should we print labels for the bars?

    If IsSafeArray(m_arrLabels) Then
	    iMax = UBound(m_arrLabels)
	    Response.Write "<tr><td>&nbsp;</td><td>&nbsp;</td>"

		For i = 0 To iMax
			iTmp = m_arrLabels(i)
			Response.Write "<td class=""chart1_labels"">" & iTmp & "</td>"
		Next

	End If

'---------------------------- Print Totals at the bottom of the chart?

    If m_lngTotalsMode = 2 Then
	    Response.Write "</tr><tr><td colspan=""" & (iMax + 4) & """ class=""chart1_totals_bottom"">" & _
		    "<img src=""" & m_sBarImage & """ width=""" & _
		    m_lngBarImageWidth & """ height=""" & m_lngBarImageWidth & """ alt=""""  />&nbsp;" & m_sLegendText & " = " & m_lngTotal
						
		'// Print the average info?
		If Len(m_sDisplayAverageBar) > 0 Then
		    Response.Write "&nbsp;&nbsp;&nbsp;<img src=""" & m_sDisplayAverageBar & """ width=""" & m_lngBarImageWidth & """ height=""" & _
		        m_lngBarImageWidth & """ alt="""" />&nbsp;Average = " & m_lngAverage
		End If

		Response.Write "</td></tr>"

	Else
	    Response.Write "</tr><tr><td colspan=""" & (iMax + 4) & """ class=""chart1_totals_bottom"">&nbsp;</td></tr>"
	End If


'---------------------------- Should we print a footer text?
    If Len(m_sFooterText) > 0 Then Response.Write "<tr><td colspan=""" & (iMax + 4) & """ class=""chart1_footer"">" & m_sFooterText & "</td></tr>"

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

	m_arrData = arrData

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