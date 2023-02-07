Attribute VB_Name = "ModLoop_Alternative"
Option Explicit

Function getModMembers() As Object

    Dim rng, rngStart, rngEnd As Range
    Dim MOD_addin As FMAddIn
    Dim MOD_conn As Connect
    Dim TableAtCell As FMTable
    Dim table As FMTable
    Dim cSheet, cRow, cCol, sCurrentSheet, MainTableName, MainTableIndex, t, iTotalValue, modMember
    Dim iDataRowStart, iDataRowEnd, i, mgColumn, iMerged As Integer
    Dim myString As String
    
    If MOD_addin Is Nothing Then
        Set MOD_conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
        Set MOD_addin = MOD_conn.FMAddIn
    End If
    
    cSheet = "Product"      'This is the table that we loop through the Mods
    cRow = 2
    cCol = 1

    sCurrentSheet = ActiveSheet.Name
    Sheets(cSheet).Select
    Set TableAtCell = MOD_addin.findTable(cSheet, cRow, cCol)

    If TableAtCell Is Nothing Then
        MsgBox "You must be in a table."
        Exit Function
    Else
        MainTableName = TableAtCell.Code
        MainTableIndex = TableAtCell.index
    End If
    
    t = Timer
    Set table = MOD_addin.Tables(MainTableName)
    table.Refresh True
    'Debug.Print "MOD Table Refresh took : " & Timer - t
    
    'I have these for error checking
    iDataRowStart = table.Position(fmArea_Data, fmType_startRow)
    iDataRowEnd = table.Position(fmArea_Data, fmType_endRow)
    mgColumn = table.Position(fmArea_Row, fmType_startColumn)   'It should always be the first column in the rows section
    
    Set getModMembers = New Collection
    Set gGetModDetailGSV = New Collection
    iTotalValue = 0
    
    For i = iDataRowStart To iDataRowEnd
        Set rng = Cells(i, mgColumn)
        If rng.MergeCells Then

            Set rng = rng.MergeArea
            rng.Select
            Set rngStart = rng.Cells(1, 1)
            Set rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

            For iMerged = rngStart.Row To rngEnd.Row
                myString = myString & Cells(iMerged, mgColumn + 1).Value & ";" & Cells(iMerged, mgColumn + 2).Value
                iTotalValue = iTotalValue + Cells(iMerged, mgColumn + 2).Value

                If iMerged <> rngEnd.Row Then  'I don't add the underscore on the last record
                    myString = myString & "_"
                End If
            Next iMerged
            
            modMember = Cells(rngStart.Row, mgColumn).Value
            getModMembers.Add myString, modMember   'add the string of products and their values to the collection with the Mod as the key (value, key)
            myString = vbNullString
            i = iMerged - 1
            
        Else
            modMember = Cells(rng.Row, mgColumn).Value
            myString = Cells(rng.Row, mgColumn + 1).Value & ";" & Cells(rng.Row, mgColumn + 2).Value
            getModMembers.Add myString, modMember   'add the string of products and their values to the collection with the Mod as the key (value, key)
            myString = vbNullString
            iTotalValue = Cells(rng.Row, mgColumn + 2).Value
        End If
        
        'Create a GSV collection in case any of the members use the GSV value.
        gGetModDetailGSV.Add iTotalValue, modMember
        iTotalValue = 0
    
    Next
    
    Sheets(sCurrentSheet).Select
    Set MOD_conn = Nothing
    Set MOD_addin = Nothing
End Function

Function getMods(modRange) As Object
    'Get a list of all the mod/shippers in the the zupload table and add to a collection
    'Requires you to send in the range of the mod/shippers

    Dim rng As Range
    Dim addin As FMAddIn
    Dim conn As Connect
    Dim modMember As String
    Dim iDataRowStart As Integer, iDataRowEnd As Integer, i As Integer, mgColumn As Integer
    Dim myString As String
    Dim ws_zupload As Worksheet
    Dim ZuploadRng As Range
    
    If addin Is Nothing Then
        Set conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
        Set addin = conn.FMAddIn
    End If

    Set ws_zupload = ThisWorkbook.Worksheets(gStrSourceSheetNm)            'This is the Zupload Worksheet (Zws)
    ws_zupload.Select
    ws_zupload.AutoFilterMode = False
    Set ZuploadRng = ws_zupload.Rows(1).Cells      'I only want the first row since that is what the time is based off.  If there is extra data in there without a time column i don't want it

    Set ZuploadRng = ws_zupload.Cells
    iDataRowEnd = Last("getRow", ZuploadRng)
    iDataRowStart = modRange.Row + 1
    mgColumn = modRange.Column
    
    Set getMods = New Collection
    
    For i = iDataRowStart To iDataRowEnd
        Set rng = Cells(i, mgColumn)
        modMember = Cells(rng.Row, mgColumn).Value
        myString = modMember
            On Error Resume Next                    'Catches any errors with duplicate Mods being added
                getMods.Add myString, modMember     '(Value, Key) - Need to the key so i can stop duplicates from getting in
            On Error GoTo 0

            modMember = vbNullString
    Next

    Set conn = Nothing
    Set addin = Nothing
End Function


