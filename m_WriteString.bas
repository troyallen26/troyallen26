Attribute VB_Name = "m_WriteString"
Option Explicit

Function writeDimensionString(dataDM)
'This writes out the dimension string for the simpleWriteback call

Dim x As Long, y As Long
Dim maxX As Long, minX As Long
Dim maxY As Long, minY As Long
Dim tempArr As Variant
Dim myDataString As String

minX = LBound(dataDM)
maxX = UBound(dataDM)
y = gDimIDCount     'This is the y value where the dimensionID is stored in the dm Array

myDataString = ""
For x = minX To maxX
    If x = minX Then
        myDataString = myDataString & "[" & Chr(34) & dataDM(x, y) & Chr(34) & ","  'Get Opening Brackets
    ElseIf x = maxX Then
        myDataString = myDataString & Chr(34) & dataDM(x, y) & Chr(34) & "]"        'Get the closing Brackets
        Exit For
    Else
        myDataString = myDataString & Chr(34) & dataDM(x, y) & Chr(34) & ","
    End If
Next x
writeDimensionString = myDataString
'Debug.Print myDataString
End Function

Function GetLastRecordArray(dataDM)
Dim x As Long, y As Long
Dim maxX As Long, minX As Long
Dim maxY As Long, minY As Long
Dim tempArr As Variant
Dim myDataString As String
Dim bStopRun As Boolean
    
    'maxX = UBound(dataDM, 1)
    'minX = LBound(dataDM, 1)
    x = 0
    maxY = UBound(dataDM, 2)
    minY = LBound(dataDM, 2)
    
For y = maxY To minY Step -1
    If Not IsEmpty(dataDM(x, y)) Then
        GetLastRecordArray = y
        Exit For
    End If
Next y
    
    
End Function

Public Function Join2D(ByVal dataDM As Variant, Optional ByVal runType As String = " ", Optional ByVal sLineDelim As String = vbNewLine) As String
'*****************************************************************************************************
'This writes the the input data for our writeback
'This code uses join for better performance
'5/4/2022
'*****************************************************************************************************
    
Dim i As Long, j As Long
Dim aReturn() As String
Dim bReturn() As String
Dim aLine() As String
Dim stringIt As String
Dim maxX, minX, minY, maxY

'Get Upper and Lower Bounds
If runType = "New" Then        'If we are getting the passed in "New" value then we only care about the second to last X in the array
    maxX = UBound(dataDM, 1) - 1
    minX = maxX
ElseIf runType = "Old" Then    'If we are getting the passed in "Old" value then we only care about the last X in the array
    maxX = UBound(dataDM, 1)
    minX = maxX
ElseIf runType = "Data" Then
    minX = LBound(dataDM, 1)
    maxX = UBound(dataDM, 1) - 2
ElseIf runType = "Currency" Then
    maxX = gCurrencyDimCount
    minX = maxX
End If

'Set the number of records for each dimension
minY = LBound(dataDM, 2)
maxY = GetLastRecordArray(dataDM)   'If each row has a different number of values then the array will not be full.  This finds out where the last record stops

'Resize the Arrays
ReDim aReturn(minX To maxX)
ReDim bReturn(minX To maxX)
ReDim aLine(minY To maxY)

For i = minX To maxX
    For j = LBound(dataDM, 2) To maxY
        aLine(j) = dataDM(i, j)                         'Creates a single dimensional array for each dimension value
    Next j
    
    If runType = "Data" Then
        aReturn(i) = Chr(34) & Join(aLine, Chr(34) & "," & Chr(34)) & Chr(34)   'This takes that entire single dimensional array and puts it all into 1 string line with the needed quotes, commas, brackets, etc...
    Else
        aReturn(i) = Join(aLine, ",")
    End If
    
    If i = maxX Then
        bReturn(i) = "[" & aReturn(i) & "]"
    Else
        bReturn(i) = "[" & aReturn(i) & "],"
    End If
Next i


Join2D = Join(bReturn, sLineDelim)
'Debug.Print Join2D
'Call FSOCreateAndWriteToTextFile(Join2D, "Join2D.txt")
End Function

Function getArray(dm As Variant) As Variant
'****************************************************************************************************
'This creates a dimension array for required dimensions.  I remove the Trader, Source and Frequency
'This is used for the writebackToken
'dm is our Original Dimension with Value array.
'   dm(x,0) = Dimension Code
'   dm(x,1) = Dimension Member Code
'   dm(x,2) = Dimension Id
'   dm(x,3) = Dimension Member ID
'   dm(x,4) = Dimension Type Id
'5/4/2022
'****************************************************************************************************
Dim x As Long, y As Long
Dim maxX As Long, minX As Long
Dim maxY As Long, minY As Long
Dim myValue As String
Dim tempArr As Variant
Dim removeArray As Variant
Dim iCount As Long

removeArray = Array("TRADER", "SOURCE", "FREQUENCY")
'Get Upper and Lower Bounds
maxX = UBound(dm, 1)
minX = LBound(dm, 1)

y = gDimTypeIdCount             'This is the dimensionTypeId
ReDim tempArr(minX To maxX - 3) 'I subtract 1 for the 0 base, and the other two because the last two sets are for NewValue and OldValue of the crossings

iCount = 0
For x = minX To maxX
    myValue = dm(x, 0)
    If Not IsInArray(myValue, removeArray) Then
        tempArr(iCount) = dm(x, y)
        iCount = iCount + 1
    End If
Next x

getArray = tempArr
    
End Function


