Attribute VB_Name = "m_WriteException"
Option Explicit

Sub WriteException(dataArray As Variant, sFailureReason As String)
'********************************************************************************************
' writes the exceptions to the exception sheet along with the reason for the write failure
'********************************************************************************************
Dim oExceptionsheet As Worksheet

    'if the exception worksheet doesn't exist then we don't want to error when setting the global exception sheet value
    On Error Resume Next
    Set oExceptionsheet = ThisWorkbook.Worksheets(gExceptionSheetNm)
    On Error GoTo 0
    
    'Creates the exception sheet if it doesn't exist or selects it if it does
    If oExceptionsheet Is Nothing Then
        Sheets.Add().Name = gExceptionSheetNm
    Else
        Sheets(gExceptionSheetNm).Select
    End If
    
    'sets the exception sheet again in the case that it didnt' exist in first try
    Set oExceptionsheet = ThisWorkbook.Worksheets(gExceptionSheetNm)
    
    
    With oExceptionsheet
        'if it's the first time the exception sheet has been accessed during the current zupload run the headers are added
        If gBlnRunFirst = True Then
            .Range(Cells(iExceptionRow, 1), Cells(iExceptionRow, 1)) = "DP Material"
            .Range(Cells(iExceptionRow, 2), Cells(iExceptionRow, 2)) = "DP Location"
            .Range(Cells(iExceptionRow, 3), Cells(iExceptionRow, 3)) = "DP Customer"
            .Range(Cells(iExceptionRow, 4), Cells(iExceptionRow, 4)) = "Channel"
            .Range(Cells(iExceptionRow, 5), Cells(iExceptionRow, 5)) = "DispMods/Shipper"
            .Range(Cells(iExceptionRow, 6), Cells(iExceptionRow, 6)) = "Currency"
            .Range(Cells(iExceptionRow, 7), Cells(iExceptionRow, 7)) = "Sales Org"
            .Range(Cells(iExceptionRow, 8), Cells(iExceptionRow, 8)) = "Year"
            .Range(Cells(iExceptionRow, 9), Cells(iExceptionRow, 9)) = "Period"
            .Range(Cells(iExceptionRow, 10), Cells(iExceptionRow, 10)) = "Failure Reason"
            .Range(Cells(iExceptionRow, 11), Cells(iExceptionRow, 11)) = "Upload Process"
            gBlnRunFirst = False
            iExceptionRow = iExceptionRow + 1
        End If
        'Adding the exception data
        .Range(Cells(iExceptionRow, 1), Cells(iExceptionRow, 1)) = dataArray(gProdDimCount, 1) 'Product
        .Range(Cells(iExceptionRow, 2), Cells(iExceptionRow, 2)) = dataArray(gLocDimCount, 1) 'Location
        .Range(Cells(iExceptionRow, 3), Cells(iExceptionRow, 3)) = dataArray(gCustDimCount, 1) 'Customer
        .Range(Cells(iExceptionRow, 4), Cells(iExceptionRow, 4)) = dataArray(gChannelDimCount, 1) 'Channel
        .Range(Cells(iExceptionRow, 5), Cells(iExceptionRow, 5)) = dataArray(gMaterialGroupDimCount, 1) 'DispMods/Shipper
        .Range(Cells(iExceptionRow, 6), Cells(iExceptionRow, 6)) = gOrigCurrency 'Currency
        .Range(Cells(iExceptionRow, 7), Cells(iExceptionRow, 7)) = dataArray(gSlsOrgDimCount, 1) 'Sales Org
        .Range(Cells(iExceptionRow, 8), Cells(iExceptionRow, 8)) = dataArray(gYearDimCount, 1) 'Year
        .Range(Cells(iExceptionRow, 9), Cells(iExceptionRow, 9)) = dataArray(gPeriodDimCount, 1) 'Period
        .Range(Cells(iExceptionRow, 10), Cells(iExceptionRow, 10)) = sFailureReason 'Reason for Failure
        .Range(Cells(iExceptionRow, 11), Cells(iExceptionRow, 11)) = gStrSourceSheetNm 'Zupload sheet where failure occured
    End With
    iExceptionRow = iExceptionRow + 1
    ActiveSheet.Cells.EntireColumn.AutoFit
    ActiveSheet.Cells.EntireRow.AutoFit
    Sheets(gStrSourceSheetNm).Select
    

End Sub

