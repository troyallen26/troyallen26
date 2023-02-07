Attribute VB_Name = "m_GlobalValues"
Option Explicit

'Public cube As FMCube
'Public tcube As FMCube
'Public xcube As FMCube
'Public uCube As FMCube
'
'Public Crossings As FMCrossingsCollection
'Public Crossing As FMCrossing
'Public tCrossings As FMCrossingsCollection
'Public tCrossing As FMCrossing
'Public xCrossings As FMCrossingsCollection
'Public xCrossing As FMCrossing
'Public uCrossings As FMCrossingsCollection
'Public uCrossing As FMCrossing
'Public crs As FMCrossing
'Public member As FMMember
'
'Public dm() As String
'Public tdm() As String
'Public xdm() As String
'Public gTable As FMTable
'Public bln_pwdOk As Boolean
'Public RunRefresh As Boolean
'
Public gStrSourceSheetNm As String
Public gExceptionSheetNm As String
'Public DPMaterialPathRng As Range
'Public SlsOrgPathRng As Range
'Public DPCustomerPathRng As Range
'Public DPLocationPathRng As Range
'Public ChannelPathRng As Range
'
'Public gMatGroupPathRng As Range        'xxx New for Canada
'Public gMonthCount As Integer
Public gOrigCurrency As String
Public gBlnRunFirst As Boolean
Public iExceptionRow As Long
'Public gGSVTotal As Variant
'
Public gPeriodDimCount As Long
Public gProdDimCount As Long
Public gSlsOrgDimCount As Long
Public gCustDimCount As Long
Public gLocDimCount As Long
Public gCurrencyDimCount As Long        'xxx New for Canada
Public gYearDimCount As Long
'Public gWkDimCount As Long
Public gChannelDimCount As Long
Public gMaterialGroupDimCount As Long   'xxx New for Canada
'
Public sMainTableName As String
Public gURL As String
'Public sMainModel As String
Public gVirtualCubeId As String
Public gFormsetId As String
Public gDimIDCount As Long
Public gDimTypeIdCount As Long
Public gServer As String

Public gFolderHeader As String
Public gGetModDetailGSV As Object

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function


Sub TurnOffEvents()
'*****************************************************
'   This turns off all events to make code run faster
'*****************************************************

With Application
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .EnableEvents = False
    .DisplayAlerts = False
End With

End Sub


Sub TurnOnEvents()
'*****************************************************
'   This turns on all events
'*****************************************************

With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
    .DisplayAlerts = True
End With

End Sub

Function GetRange(strName, strSheet) As Range
'***************************************************************************************************
'   gets the range of a string name on a passed in worksheet.
'   strName = What you are searching for
'   strSheet = Worksheet you want to check
'***************************************************************************************************
Dim rNa As Range

Sheets(strSheet).Select         'Selects the sheet that was passed in
With ActiveSheet
    .UsedRange.Select
    .AutoFilterMode = False     'Turns of filtering to ensure all data is captured
End With

On Error Resume Next
With Selection
    Set rNa = .Find(What:=strName, _
        LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False)
    Set GetRange = rNa
    Set rNa = Nothing
End With

End Function


Function Last(choice As String, rng As Range, Optional ByVal SearchString)
' getRow = last row
' getColumn = last column
' getCell = last cell
Dim lrw As Long
Dim lcol As Long

Select Case choice

Case "getRow":
    On Error Resume Next
    Last = rng.Find(What:="*", _
                    After:=rng.Cells(1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    On Error GoTo 0

Case "getColumn":
    On Error Resume Next
    Last = rng.Find(What:="*", _
                    After:=rng.Cells(1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
    On Error GoTo 0

Case "getCell":
    On Error Resume Next
    lrw = rng.Find(What:="*", _
                   After:=rng.Cells(1), _
                   LookAt:=xlPart, _
                   LookIn:=xlFormulas, _
                   SearchOrder:=xlByRows, _
                   SearchDirection:=xlPrevious, _
                   MatchCase:=False).Row
    On Error GoTo 0

    On Error Resume Next
    lcol = rng.Find(What:="*", _
                    After:=rng.Cells(1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
    On Error GoTo 0

    On Error Resume Next
    Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
    If Err.Number > 0 Then
        Last = rng.Cells(1).Address(False, False)
        Err.Clear
    End If
    On Error GoTo 0

End Select
End Function

Sub CheckValueLen(iStartRow, iLastRow, LenArray)
'*****************************************************************************************************************************************
' This checks the length of the required members (Customer, location, material, etc..) and if it isn't correct it adds the leading zeros
'*****************************************************************************************************************************************
Dim myZeroString As String
Dim iCheckCount As Long, iCheckIt As Long, iGetLength As Long, iZeroToAdd As Long, iZero As Long

myZeroString = ""                       'setting the string to be blank
Sheets(gStrSourceSheetNm).Select        'selecting the zupload sheet

'Loops through each row that is passed in to check the values
For iCheckCount = iStartRow To iLastRow
    For iCheckIt = 0 To UBound(LenArray)
        iGetLength = Len(Cells(iCheckCount, LenArray(iCheckIt, 2)))         'Getting the length of the cell it is looking at
        If IsNumeric(Cells(iCheckCount, LenArray(iCheckIt, 2))) Then        'Checking to see if it is numeric, if not it skips it
            If iGetLength < LenArray(iCheckIt, 1) Then                      'Checking to see if the cell length matches the required lenght and if not it adds required number of leading zeros
                iZeroToAdd = LenArray(iCheckIt, 1) - iGetLength
                    For iZero = 1 To iZeroToAdd
                        myZeroString = myZeroString & "0"
                    Next iZero
                Cells(iCheckCount, LenArray(iCheckIt, 2)).Select
                Selection.NumberFormat = "@"                                'sets the format of the cell to be generic
                Cells(iCheckCount, LenArray(iCheckIt, 2)) = myZeroString & Cells(iCheckCount, LenArray(iCheckIt, 2))
                myZeroString = ""
            ElseIf iGetLength > LenArray(iCheckIt, 1) Then
                Cells(iCheckCount, LenArray(iCheckIt, 2)).Select
                Selection.NumberFormat = "@"
                Cells(iCheckCount, LenArray(iCheckIt, 2)) = Trim(Cells(iCheckCount, LenArray(iCheckIt, 2)))  'If the cell value is longer than the required lenght we try and trim it, otherwise i will be picked up in the Exceptions
            End If
        End If
    Next iCheckIt
Next iCheckCount

End Sub

Function IsArrayBlank(arr As Variant) As Boolean
'*********************************************************************************************
' Checks an array to see if it has values and returns a TRUE or FALSE based on the results
'*********************************************************************************************

On Error Resume Next
    IsArrayBlank = IsArray(arr) And _
        Not IsError(LBound(arr, 1)) And _
         LBound(arr, 1) <= UBound(arr, 1)
End Function

Function GetTableNameFromSheet(MySheetName) As FMTable
Dim addin As FMAddIn
Dim conn As Connect
Dim table As FMTable
Dim cSheet As String
Dim cRow As Long
Dim cCol As Long
Dim tmpCell As Range
Dim rng As Range
Dim sCurrentSheet As String

    If addin Is Nothing Then
        Set conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
        Set addin = conn.FMAddIn
    End If

    If addin.IsLoggedIn = False Then
        MsgBox "You must be logged in to SAS Financial Management."
        Exit Function
    End If
    
    'Get the Table name from the active sheet
    cSheet = MySheetName
    sCurrentSheet = ActiveSheet.Name
    Sheets(cSheet).Select
    
        For Each tmpCell In Range("A1:Z25")
            If Not IsEmpty(tmpCell) Then
                Set table = addin.findTable(cSheet, tmpCell.Row, tmpCell.Column)
                If Not (table Is Nothing) Then
                    Set GetTableNameFromSheet = table
                    Exit For
                End If
            End If
            'Debug.Print tmpCell.Address
        Next tmpCell
        
    Sheets(sCurrentSheet).Select
    Set conn = Nothing
    Set addin = Nothing
End Function




