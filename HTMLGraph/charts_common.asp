<%
'============================================================
' MODULE:    charts_common.asp
' APP:           HTMLGraph
' AUTHOR:    www.u229.no
' CREATED:  June 2005
'============================================================
' COMMENT: Support routines for the 6 class files
'============================================================
' ROUTINES:
' - Function ReturnArrayDimensions(arr)
' - Function IsSafeArray(arr)
' - Function ShrinkDimension(arrData)
' - Function DataShrink(iData, iFactor)
' - Function ReturnAverage()
'============================================================


'------------------------------------------------------------------------------------------------------------
' Comment: Find the dimension of the data arrays to be used with the charts. We accept only 1 or 2 dimensions.
'------------------------------------------------------------------------------------------------------------
Function ReturnArrayDimensions(arr)
    On Error Resume Next

	Dim l, iDummy

	For l = 1 To 60
		iDummy = UBound(arr, l)
		If Err Then Err.Clear: Exit For
	Next

	ReturnArrayDimensions = CLng(l - 1)

End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Returns true if array has values, false if empty.
'------------------------------------------------------------------------------------------------------------
Function IsSafeArray(arr)
    On Error Resume Next

    Dim lUbound

    If Not IsArray(arr) Then Exit Function

    lUbound = UBound(arr, 1)

    IsSafeArray = (Err.Number = 0)
        
End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Reduces 2 dimensional arrays to 1 dimension.
'------------------------------------------------------------------------------------------------------------
Function ShrinkDimension(arrData)
    On Error Resume Next

	Dim arrTemp()
	Dim i, j

    i = 0	

    For j = 0 To UBound(arrData, 2)
		Redim Preserve arrTemp(i + j)
        arrTemp(i + j) = arrData(i, j)
    Next

	ShrinkDimension = arrTemp

End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Shrink your data to fit into the charts.
'------------------------------------------------------------------------------------------------------------
Function DataShrink(iData, iFactor)
    On Error Resume Next

    Dim iTmp
    
	'// FormatNumber will cut off any decimals: 216.55 becomes 217
    iTmp = CLng(iData / CLng(iFactor))
	If iTmp < 1 And iTmp > 0 Then iTmp = 1
	DataShrink = iTmp

End Function


'------------------------------------------------------------------------------------------------------------
' Comment: Calculate the average value of a single data set
'------------------------------------------------------------------------------------------------------------
Function ReturnAverage()
    On Error Resume Next

    Dim i, iMax, iTotal
    
	iMax = UBound(m_arrData)

	For i = 0 To iMax
	    iTotal = (iTotal + CLng(m_arrData(i)))
	Next

	ReturnAverage = CLng(iTotal / (iMax + 1))

End Function
%>