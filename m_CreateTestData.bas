Attribute VB_Name = "m_CreateTestData"
Option Explicit

Sub createdata()
Dim myArray As Variant
Dim iLastLongestRow As Long, iLastSecondLongestRow As Long, iLastThirdLongestRow As Long, iLastFourthLongestRow As Long, iLastRow As Long, iCount As Long, iFirst As Long, iSecond As Long, iThird As Long
Dim iFourth As Long
Dim iArrayCount As Double

iLastLongestRow = Range("A5000").End(xlUp).Row          'Cust
iLastSecondLongestRow = Range("B5000").End(xlUp).Row    'Loc
iLastThirdLongestRow = Range("C5000").End(xlUp).Row     'Channel
iLastFourthLongestRow = Range("D5000").End(xlUp).Row    'Product

iLastRow = Cells(1).CurrentRegion.Rows.Count
iCount = 0
iArrayCount = (iLastLongestRow) * (iLastSecondLongestRow) * (iLastThirdLongestRow) * (iLastFourthLongestRow)
ReDim myArray(iArrayCount - 1, 5)
For iFirst = 1 To iLastRow
    'myArray(iProd - 1, iLoc - 1) = Cells(iProd, 2)
    For iSecond = 1 To iLastSecondLongestRow
        For iThird = 1 To iLastThirdLongestRow
            For iFourth = 1 To iLastFourthLongestRow
                myArray(iCount, 0) = Cells(iFourth, 4)      'Prod
                myArray(iCount, 1) = Cells(iSecond, 2)      'Loc
                myArray(iCount, 2) = Cells(iFirst, 1)       'Cust
                myArray(iCount, 3) = Cells(iThird, 3)       'Channel
                myArray(iCount, 4) = "CS"                   'UOM
                myArray(iCount, 5) = "1001"                 'SlsOrg
            
                iCount = iCount + 1
            Next iFourth
        Next iThird
    Next iSecond
    'icount = icount - 1
Next iFirst

Sheets("Zupload (2)").Select
Range(Cells(2, 1), Cells(iArrayCount, 6)) = myArray

End Sub


Sub WriteWeeks()
Dim i As Integer
Dim iPeriodCnt As Integer
Dim iCount As Integer
Dim wb As Workbook
Dim rng As Range
Dim SourceSheetNm As String
Dim wbName As String
Dim iColumn As Long
Dim sStartYear As String
Dim sStartWeekResult As String
Dim iStartWeek As Integer, sWeekCount As Integer
Dim sWeekCountResult As String
Dim iColumnStart As Long, iLastColumn As Long
Dim bNewYear As Boolean
Dim t


Call TurnOffEvents
iPeriodCnt = 52
SourceSheetNm = ActiveSheet.Name    'Picks up the upload sheet
Set wb = ThisWorkbook
wbName = wb.Name

If InStr(1, SourceSheetNm, "Z") = 1 Then
    If InStr(1, wbName, "MOD") Then
        iColumn = 8
    Else
        iColumn = 7
    End If
Else
    MsgBox "You must be in a Zupload tab to run WriteWeeks."
    Exit Sub
End If

 sStartYear = InputBox(Prompt:="Enter the Start Year.", _
          Title:="ENTER START YEAR")
If StrPtr(sStartYear) = 0 Then
    'User Cancelled
    Exit Sub
ElseIf sStartYear = vbNullString Then
    'No data entered
    MsgBox ("Please enter year and try again.")
    Exit Sub
Else
    'Checking to see if value is numeric
    If Not IsNumeric(sStartYear) Then
        MsgBox ("Numeric start year is required")
        Exit Sub
    End If
End If

sStartWeekResult = InputBox(Prompt:="Enter the Start Week.", _
      Title:="ENTER START WEEK")
If StrPtr(sStartWeekResult) = 0 Then
    'User Cancelled
    Exit Sub
ElseIf sStartWeekResult = vbNullString Then
    'No data entered
    MsgBox ("Please enter start week and try again.")
    Exit Sub
Else
    'Checking to see if value is numeric
    If IsNumeric(sStartWeekResult) Then
        iStartWeek = CInt(sStartWeekResult)
        If iStartWeek > 53 Or iStartWeek < 1 Then
            MsgBox ("Please select a start week between 1 and 53")
            Exit Sub
        End If
    Else
        'non numeric value entered
        MsgBox ("Numeric start week is required")
        Exit Sub
    End If
End If

sWeekCountResult = InputBox(Prompt:="Enter the number of Weeks.", _
      Title:="ENTER NUMBER OF WEEKS")
If StrPtr(sWeekCountResult) = 0 Then
    'User Cancelled
    Exit Sub
ElseIf sWeekCountResult = vbNullString Then
    'No data entered
    MsgBox ("Please enter number of weeks and try again.")
    Exit Sub
Else
    'Checking to see if value is numeric
    If IsNumeric(sWeekCountResult) Then
        sWeekCount = CInt(sWeekCountResult)
    Else
        'non numeric value entered
        MsgBox ("Numeric week count is required")
        Exit Sub
    End If
End If

t = Timer

iColumnStart = iColumn
bNewYear = False

'This clears out the existing time columns
Set rng = ActiveSheet.Rows(1).Cells      'I only want the first row since that is what the time is based off.  If there is extra data in there without a time column i don't want it
iLastColumn = Last("getColumn", rng)

iLastColumn = LastCol(ActiveSheet)
With ActiveSheet

    If iLastColumn >= iColumn Then
        .Columns(ReturnName(iColumn) & ":" & ReturnName(iLastColumn)).Delete Shift:=xlToLeft
    End If
End With

iCount = iStartWeek
'iYearCnt = Application.WorksheetFunction.RoundUp((sWeekCount - iTYWeeks) / iPeriodCount, 0)
'Dim wkArray(0 To sWeekCount - 1) As Variant
'iWeek = 0
iColumnStart = iColumn
For i = 1 To sWeekCount
    bNewYear = isDivisible(iCount, iPeriodCnt + 1)
    If iCount <> iStartWeek And bNewYear Then
        'This is a new Year
        sStartYear = sStartYear + 1
        iCount = 1
    End If

    If iCount < 10 Then
        Cells(1, iColumnStart) = "W0" & iCount & "/" & sStartYear
        'wkArray(iWeek) = "W0" & iCount & "/" & sStartYear
    Else
        Cells(1, iColumnStart) = "W" & iCount & "/" & sStartYear
        'wkArray(iWeek) = "W" & iCount & "/" & sStartYear
    End If
    iColumnStart = iColumnStart + 1
    iCount = iCount + 1
    'iWeek = iWeek + 1

Next i

'Dim rngData As Range
'Set rngData = Sheets(SourceSheetNm).Range(Cells(1, iColumn), Cells(1, (iColumn + sWeekCount) - 1))
'
'Dim arrayTransposed As Variant
'arrayTransposed = Application.Transpose(wkArray)
'rngData = arrayTransposed
Call TurnOnEvents
Debug.Print "Write Time Members Time:  " & Timer - t * 1000

End Sub



Function isDivisible(x As Integer, d As Integer) As Boolean
    Dim rmndr As Double
    rmndr = x Mod d
    If rmndr = 0 Then
        isDivisible = True
    Else
        isDivisible = False
    End If
End Function


Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
    If IsEmpty(LastCol) Then
        LastCol = 16384
    End If
    
End Function

Sub Sample()
    With Sheet1
        'A:CV
        .Columns(ReturnName(1) & ":" & ReturnName(100)).Delete Shift:=xlToLeft
    End With
End Sub

'~~> Returns Column Name from Col No
Function ReturnName(ByVal num As Integer) As String
    ReturnName = Split(Cells(, num).Address, "$")(1)
End Function
