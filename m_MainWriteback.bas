Attribute VB_Name = "m_MainWriteback"
Option Explicit

'This is the newest code for VIYA
Sub Caller1()
'**************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPCAUF1 which is for (Innovation/PWR)
'**************************************************************************************************************************

    Call Main_Writeback("ZDPCAUF1")

End Sub

Sub Caller3()
'****************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPEXKF1 which is for (Sales Adjustment)
'****************************************************************************************************************************

    Call Main_Writeback("ZDPEXKF1")
    
End Sub

Sub Caller4()
'*******************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPMKTFC which is for (Marketing Adjustment)
'*******************************************************************************************************************************

    Call Main_Writeback("ZDPMKTFC")
    
End Sub

Sub Caller5()
'*************************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPCAUF2 which is for (Cannibalization Adjustment)
'*************************************************************************************************************************************

    Call Main_Writeback("ZDPCAUF2")
    
End Sub

Sub Caller6()
'*******************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPCAUF4 which is for (Other Adjustment)
'*******************************************************************************************************************************

    'Call cubeJustWrite4("ZDPCAUF4")
    Call Main_Writeback("ZDPCAUF4")
    
End Sub
   
Sub Caller7()
'*******************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPSTFC3 which is for (Planner Adjustment)
'*******************************************************************************************************************************

    Call Main_Writeback("ZDPSTFC3")
    
End Sub

Sub Caller8()
''*******************************************************************************************************************************
''calls the main code (Writeback_Main and passes the key figure value ZDPCUSTF which is for (Customer Forecast)
''*******************************************************************************************************************************

    Call Main_Writeback("ZDPCUSTF")

End Sub

Sub Caller9()
'*******************************************************************************************************************************
'calls the main code (Writeback_Main and passes in the adjustment key figure value ZDPFSACC which is for (All In Stat Accepted)
'*******************************************************************************************************************************

    Call Main_Writeback("ZDPFSACC")
    
End Sub

Public Sub Main_Writeback(sKeyFigure)
'*************************************************************************************************************************************************************
' Main code section.  This code creates the crossings from the zupload sheet that the user pastes in and writes all valid data from that data back to the CWB
'*************************************************************************************************************************************************************
Dim newValue As Variant
Dim sSheet As String
Dim cSheet As String
Dim returncode
Dim addin As FMAddIn
Dim conn As Connect
Dim table As FMTable
'Public cube As FMCube
'Public Crossings As FMCrossingsCollection
'Public Crossing As FMCrossing
Dim cube As FMCube
Dim Crossings As FMCrossingsCollection
Dim Crossing As FMCrossing
Dim iStartRow As Long, iEndRow As Long
Dim iStartColumn As Long, iEndColumn As Long, iRequiredColumns As Long
Dim myRowCount As Long
Dim LenArray As Variant
Dim sMainModel As String
Dim blnMod As Boolean
Dim blnModFirstRun As Boolean
Dim bStopTime As Boolean
Dim sMaterialGroup As String
Dim sOrg As String, myFormsetName As String, y As String, wk As String
'Dim sCurrency As String
Dim ModMemberArray() As Variant
Dim dm() As String
Dim ProdMemberArray() As Variant
Dim UniqueValueArray() As Variant
Dim dimArray() As Variant
Dim myAPIcheckArray(4, 1) As Variant
Dim apiLoop As Long
Dim token_body As String
Dim iModMemberCount As Long, iUniqueRecordCount As Long, iRowCount As Long, iDimCount As Long, modMax As Long, iColumnCount As Long, iDm As Long
Dim ColRng As Range
Dim Zws As Worksheet
'Dim writeBackSheet As Worksheet
Dim ElementRng As Range, DPMaterialPathRng As Range, DPCustomerPathRng As Range, DPLocationPathRng As Range, ChannelPathRng As Range, CurrencyPathRng As Range, SlsOrgPathRng As Range
Dim ZuploadRng As Range
Dim MatGroupPathRng As Range
Dim index As Integer
Dim iDataCol As Long
Dim sTime As String
Dim errorReason As String
Dim sHierarchies As FMHierarchiesCollection
Dim ModMemberColl As Collection
Dim ModColl As Collection
Dim tTime, t, rowTime
Dim dimension, writebackSuccess
Dim checkString As String
Dim yr As Integer, lastYr As Integer
Dim sYear As String
Dim token As String
Dim GSVTotal As Double, PctValue As Double
Dim planningAreaId As String
Dim FolderId As String
Dim writebackToken As String
Dim p_body As String
Dim wb As Workbook

tTime = Timer   'This is the total time
t = Timer       'This is the time we use to the next major code break

'makes sure there is a connection to the SAS Addin for microsoft excel, if not the connection is made
If addin Is Nothing Then
    Set conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
    Set addin = conn.FMAddIn
End If

'Makes sure user is connected to the SAS Financial Managment add in and if not gives messages and exits
If addin.IsLoggedIn = False Then
    MsgBox "You must be logged in to SAS VIYA DP."
    Exit Sub
End If

'Turn off all events to make code run faster
Call TurnOffEvents

'Hardcoded stuff
gStrSourceSheetNm = ActiveSheet.Name    'Picks up the upload sheet
Set wb = ThisWorkbook
iExceptionRow = 1               'Just setting the first row of the exception sheet to be 1
gExceptionSheetNm = "Exception" 'Setting the exception sheet name (excpetion sheet is where we capture crossings that are not valid or not writeable
gBlnRunFirst = True

blnMod = False

'This prevents users form running in a non zupload sheet
If InStr(1, gStrSourceSheetNm, "Zupload") = 1 Then
    'If InStr(1, addin.User.FormSetName, "MOD") Then
    If InStr(1, wb.Name, "MOD") Then
        blnMod = True
    End If
Else
    MsgBox "You must be in a Zupload tab to run Zupload."
    Exit Sub
End If

'This deletes any left over exception sheets if they still exist
On Error Resume Next
Sheets(gExceptionSheetNm).Delete

myRowCount = 1

If blnMod = True Then
    ReDim LenArray(4, 2)    'Zupload MOD
Else
    ReDim LenArray(3, 2)    'Zupload
End If

' adding the required crossing member types to an array with the expected length of the value and the column the value lives in (Users don't have to add leading zeros but system needs them)
Set DPMaterialPathRng = GetRange("DP Material", gStrSourceSheetNm)
LenArray(0, 0) = "DP Material"
LenArray(0, 1) = 18
LenArray(0, 2) = DPMaterialPathRng.Column
Set DPCustomerPathRng = GetRange("DP Customer", gStrSourceSheetNm)
LenArray(1, 0) = "DP Customer"
LenArray(1, 1) = 10
LenArray(1, 2) = DPCustomerPathRng.Column
Set DPLocationPathRng = GetRange("DP Location", gStrSourceSheetNm)
LenArray(2, 0) = "DP Location"
LenArray(2, 1) = 4
LenArray(2, 2) = DPLocationPathRng.Column
Set ChannelPathRng = GetRange("Channel", gStrSourceSheetNm)
LenArray(3, 0) = "Channel"
LenArray(3, 1) = 2
LenArray(3, 2) = ChannelPathRng.Column

'Get the range for DispMods/Shipper (Material Group) if it's a Zupload MOD
If blnMod = True Then
    Set MatGroupPathRng = GetRange("DispMods/Shipper", gStrSourceSheetNm)
    
    'If DP Material doesn't exist on zupload sheet then turn events back on and exit
    If DPMaterialPathRng Is Nothing Then
        MsgBox "There is no data to upload"
        Call TurnOnEvents
        Exit Sub
    Else
        LenArray(4, 0) = "DispMods/Shipper"
        LenArray(4, 1) = 18
        LenArray(4, 2) = MatGroupPathRng.Column
    End If
End If

Set CurrencyPathRng = GetRange("Currency", gStrSourceSheetNm)
Set SlsOrgPathRng = GetRange("Sales Org", gStrSourceSheetNm)

cSheet = "Review"
Set table = GetTableNameFromSheet(cSheet)

If table Is Nothing Then
   MsgBox "There must be a table in the formset on the " & cSheet & " tab"
   Call TurnOnEvents
   Exit Sub
Else
   sMainTableName = table.Code
End If

sMainModel = table.Model
gURL = addin.Url
gServer = addin.Server

Set cube = addin.Cubes(sMainModel)      'set the cube to equal the model you are using for your hierarchy (used for weekly zupload files)

'Get values from the Addin and Cube - New for Viya
gVirtualCubeId = cube.ID
y = InStr(1, wb.Name, " ", False)
gFolderHeader = Mid(wb.Name, y + 1, Len(wb.Name))
'************************************************************
token = getTokenAuth()
planningAreaId = getPlanningAreaId(token)
FolderId = getFolder(token, planningAreaId)
myAPIcheckArray(0, 0) = "gVirtualCubeId"
myAPIcheckArray(0, 1) = gVirtualCubeId
myAPIcheckArray(1, 0) = "gFolderHeader"
myAPIcheckArray(1, 1) = gFolderHeader
myAPIcheckArray(2, 0) = "planningAreaId"
myAPIcheckArray(2, 1) = planningAreaId
myAPIcheckArray(3, 0) = "FolderId"
myAPIcheckArray(3, 1) = FolderId
myAPIcheckArray(4, 0) = "gFormsetId"
myAPIcheckArray(4, 1) = gFormsetId

For apiLoop = LBound(myAPIcheckArray) To UBound(myAPIcheckArray)
    If myAPIcheckArray(apiLoop, 1) = "" Then
        MsgBox myAPIcheckArray(apiLoop, 0) & "is blank, Please screenshot and send to your System Administrator"
        GoTo errorFail
    End If
Next


Debug.Print ThisWorkbook.Name
Debug.Print gVirtualCubeId
Debug.Print gFolderHeader
Debug.Print planningAreaId
Debug.Print FolderId
Debug.Print gFormsetId          'We get this value in the getFolder code
'************************************************************

'Clear the cube query
cube.ClearQuery

'Selects the zupload sheet and sets the start row (minus the header) and end row of the data
Set Zws = ThisWorkbook.Worksheets(gStrSourceSheetNm)            'This is the Zupload Worksheet (Zws)
Zws.Select
Zws.AutoFilterMode = False
Set ZuploadRng = Zws.Rows(1).Cells      'I only want the first row since that is what the time is based off.  If there is extra data in there without a time column i don't want it
iEndColumn = Last("getColumn", ZuploadRng)
Set ZuploadRng = Zws.Cells
iEndRow = Last("getRow", ZuploadRng)
iStartRow = DPMaterialPathRng.Row + 1

'Checks to make sure there is data to upload, if not exits the sub
If iEndRow <= 1 Then
    MsgBox "There is no data to upload"
    Call TurnOnEvents
    Exit Sub    'This will exit the sub if there is no data in the zupload sheet
End If

'This makes sure the values that are being used to write data arent' missing the leading zeros
Call CheckValueLen(iStartRow, iEndRow, LenArray)

Set sHierarchies = cube.ServerHierarchies

If blnMod = True Then
    iRequiredColumns = 7
    blnModFirstRun = True
    iUniqueRecordCount = ((iEndRow - 1) * 10) * (iEndColumn - iRequiredColumns)
    ReDim dataDM(sHierarchies.Count + 1, iUniqueRecordCount)    'dataDm(x,y) - for MOD we use a 10x the number of material groups x the week count.  If for some reason it's not enough i've added a check to redim preserve dataDM
Else
    iRequiredColumns = 6
    iUniqueRecordCount = (iEndRow - 1) * (iEndColumn - iRequiredColumns)    'This will be used as the number of members in a multidimensional array
    ReDim dataDM(sHierarchies.Count + 1, iUniqueRecordCount - 1)            'dataDm(x,y) - we add two to the x for newValue and OldValue, the Y for Zupload is based on the number of time columns
End If

iDataCol = 0            'This is record row for the dataDM array
gDimIDCount = 2         'This is the dm(x,y) y value.  This is the data point for the DimensionID
gDimTypeIdCount = 4     'This is the dm(x,y) y value.  This is the data point for the DimensionTypeID
Zws.Select              'Select the zupload sheet

rowTime = Timer         'This is the start time for the DM and dataDM Array Creation

For iRowCount = iStartRow To iEndRow
    'Set the dm(dimension member) array to be equal to the number of dimensions in the data cube
    ReDim dm(sHierarchies.Count - 1, 4)
    ReDim cubeDm(sHierarchies.Count - 1, 1)
    
    iDimCount = 0
    For Each dimension In sHierarchies
        
        'For the defined case's below we get the information from the Zupload sheet, everything else is the default member for the dimension in which it belongs
        dm(iDimCount, 0) = dimension.DimensionCode
        dm(iDimCount, gDimIDCount) = dimension.DimensionId
        dm(iDimCount, gDimTypeIdCount) = dimension.DimensionTypeId
        Select Case dimension.DimensionCode
            Case "ACCOUNT"
                dm(iDimCount, 1) = sKeyFigure
            Case "ANALYSIS"
                dm(iDimCount, 1) = "BASE"
            Case "ORG"
                dm(iDimCount, 1) = WorksheetFunction.Trim(Cells(iRowCount, SlsOrgPathRng.Column))       'Sales Org is captured from the zupload data
                gSlsOrgDimCount = iDimCount
            Case "PRODUCT"
                dm(iDimCount, 1) = Cells(iRowCount, DPMaterialPathRng.Column)                           'DP Material is captured from the zupload data
                gProdDimCount = iDimCount
            Case "LOCATION"
                dm(iDimCount, 1) = WorksheetFunction.Trim(Cells(iRowCount, DPLocationPathRng.Column))   'DP Location is captured from the zupload data
                gLocDimCount = iDimCount
            Case "CUSTOMER"
                 dm(iDimCount, 1) = Cells(iRowCount, DPCustomerPathRng.Column)                          'DP Customer is captured from the zupload data
                 gCustDimCount = iDimCount
            Case "CHANNEL"
                dm(iDimCount, 1) = WorksheetFunction.Trim(Cells(iRowCount, ChannelPathRng.Column))      'Channel is captured from the zupload data
                gChannelDimCount = iDimCount
            Case "CALENDAR"
                gPeriodDimCount = iDimCount
            Case "MATERIAL_GROUP"
                If blnMod = False Then
                    dm(iDimCount, 1) = "N"
                Else
                    dm(iDimCount, 1) = Cells(iRowCount, MatGroupPathRng.Column)     'xxx New - DispMods/Shipper (Material Group) is captured from the zupload data
                End If
                gMaterialGroupDimCount = iDimCount
            Case "CURRENCY"
                gOrigCurrency = UCase(Cells(iRowCount, CurrencyPathRng.Column))
                dm(iDimCount, 1) = gOrigCurrency
                If blnMod = True Then
                    dm(iDimCount, 1) = UCase(Cells(iRowCount, CurrencyPathRng.Column))     'xxx New - DispMods/Shipper (Material Group) is captured from the zupload data
                    If dm(iDimCount, 1) = "CS" Then
                        MsgBox "CS is an invalid currency for Mod uploads.  Please use PAL or GSV."     'We force PAL so they know the value the enter for the crossing should be at pallets.
                        Call TurnOnEvents
                        Exit Sub    'This will exit the sub if there is no data in the zupload sheet
                    ElseIf dm(iDimCount, 1) = "PAL" Then
                        dm(iDimCount, 1) = "CS"             'Since we load to cases at the product level we change the MOD PAL to CS for the products we will
                    End If
                End If
                gCurrencyDimCount = iDimCount
            Case Else
                'all other members of the hierarchy are added to the array based on the default member of each dimension.  All dimenions must be accounted for in this array.
                dm(iDimCount, 1) = dimension.WriteDefaultMember.Code
        End Select
        dm(iDimCount, 3) = cube.ServerHierarchies(dimension.DimensionCode).LeafMembers(dm(iDimCount, 1)).ID         'Adding the memberId to the dimension array
        iDimCount = iDimCount + 1
    Next dimension
    
    'Gets the first active weekly column of data for each record.  Since a record.
    With Zws.Range(Cells(iRowCount, iRequiredColumns + 1), Cells(iRowCount, iEndColumn))
        Set ColRng = .Find(What:="*", _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not ColRng Is Nothing Then
            iStartColumn = ColRng.Column
        Else                                                'If there are no columns of data for the crossing we write exception and check the next row to see if there is anything there
            Call WriteException(dm, "No weekly values for Crossing")
            GoTo StartNextWeek
        End If
    End With
    
    'Getting the products associated with the mod and the case conversion values for it.
    If blnMod = True And blnModFirstRun = True Then
        
        'Set the product table with the right mod/shipper members
        Set ModColl = getMods(MatGroupPathRng)  'get the distinct mod/shipppers
        Call modSelector(ModColl)       'Set the product table to the zupload mod/shippers
        
        Set ModMemberColl = getModMembers() 'Add all the mods and their associated products into a collection
        
        If ModMemberColl.Count = 0 Then 'If no members are in the collection then exit the zupload and let the user know
            modMax = 0
            dm(gProdDimCount, 1) = ""
            MsgBox "No Valid MOD Members:  ModMemberArray was empty.  Please check the Product Tab.  Writeback Will now Exit"
            Call TurnOnEvents
            Exit Sub
        Else
            blnModFirstRun = False
        End If
    End If

    If blnMod = True Then
        checkString = dm(gMaterialGroupDimCount, 1)
        If Exists(ModMemberColl, checkString) = True Then  'Check to see if the users mod is in the mod collection.  If not write an exception and move to next row
            ModMemberArray() = ModCollSplit(ModMemberColl, dm(gMaterialGroupDimCount, 1))   'Gets the products and values for the material group passed in
            modMax = UBound(ModMemberArray)
        Else
            Call WriteException(dm, "Material Group " & dm(gMaterialGroupDimCount, 1) & " is not a valid member.")
            GoTo StartNextWeek
        End If
    Else
        modMax = 0
    End If
    '########################  Here is the end of the moved code  #################################
    
    
    For iColumnCount = iStartColumn To iEndColumn   'This loops through the zupload week columns

        If InStr(1, Cells(iStartRow - 1, iColumnCount), "WK") Then
            dm(gPeriodDimCount, 1) = "WK" & Mid(Cells(iStartRow - 1, iColumnCount), 3, 2)   'This gets the week when WK is the prefix
        ElseIf InStr(1, Cells(iStartRow - 1, iColumnCount), "W") Then
            If InStr(1, Mid(Cells(iStartRow - 1, iColumnCount), 2, 2), "/") Then
                wk = "WK" & "0" & Mid(Cells(iStartRow - 1, iColumnCount), 2, 1)
            Else
                wk = "WK" & Mid(Cells(iStartRow - 1, iColumnCount), 2, 2)
            End If
            yr = Mid(Cells(iStartRow - 1, iColumnCount), InStr(1, Cells(iStartRow - 1, iColumnCount), "/", vbTextCompare) + 1)
            If iColumnCount <> iStartColumn And yr <> lastYr Then
                sYear = Str(lastYr)
            End If
            lastYr = yr
            dm(gPeriodDimCount, 1) = yr & wk  'This gets the week when W is the prefix
            
        Else
            MsgBox "Weeks are not in the correct format"
            Call WriteException(dm, "Time Column " & Cells(iStartRow - 1, iColumnCount) & " not in correct format.  Should be W##/YYYY or WK##/YYYY")
            GoTo StartNextWeek          'Go to next Row of data
        End If
        
        'New for Viya writeback
        'dm(gPeriodDimCount, 3) = cube.ServerHierarchies("CALENDAR").LeafMembers(dm(gPeriodDimCount, 1)).ID
        dm(gPeriodDimCount, 3) = table.ServerHierarchies("CALENDAR").LeafMembers(dm(gPeriodDimCount, 1)).ID
'        If dm(gPeriodDimCount, 3) = "" Then
'            MsgBox "Period Member: " & dm(gPeriodDimCount, 1) & " has blank API Code"
'        Else
'            Debug.Print dm(gPeriodDimCount, 3)
'        End If
        
        For iModMemberCount = 0 To modMax
            'Creates a crossing record and places the dm array values in that record
            Set Crossings = cube.Crossings
            
            If blnMod = True Then       'Add the MOD DP_MATERIAL members to the dm Array
                dm(gProdDimCount, 1) = ModMemberArray(iModMemberCount, 0) 'Add the product to the array
                dm(gProdDimCount, 3) = cube.ServerHierarchies("PRODUCT").LeafMembers(dm(gProdDimCount, 1)).ID 'get the member id after the DP_MATERIAL has been added to the array
            End If
            
            'assigns the zupload value to the newValue member of the crossing record we just created
            'newDM = dmArray(dm)
            Set Crossing = cube.Crossing(dm)

            If blnMod = True Then
                If IsNumeric(Cells(iRowCount, iColumnCount) * ModMemberArray(iModMemberCount, 1)) Then
                    If dm(gCurrencyDimCount, 1) = "GSV" Then
                        GSVTotal = gGetModDetailGSV(dm(gMaterialGroupDimCount, 1))
                        PctValue = ModMemberArray(iModMemberCount, 1) / GSVTotal
                        newValue = Cells(iRowCount, iColumnCount) * PctValue  'Multiply the pal/gsv * the product multiplier
                    Else
                        newValue = Cells(iRowCount, iColumnCount) * ModMemberArray(iModMemberCount, 1)  'Multiply the pal/gsv * the product multiplier
                    End If
                Else
                    Call WriteException(dm, "Crossing not Writeable")
                    Exit For
                End If
            Else
                If Len(Trim(Cells(iRowCount, iColumnCount).Value)) > 0 Then
                    If IsNumeric(Cells(iRowCount, iColumnCount)) Then
                        newValue = Cells(iRowCount, iColumnCount)
                    Else
                        Call WriteException(dm, "Crossing not Writeable")
                        Exit For
                    End If
                    
                End If
            End If
            
            'Check to ensure the weekly crossing is writeable, if so then add the value.  If the crossing is not writeable we call the WriteException method to deal with it
            If Crossing Is Nothing Then                             'This handles bad member data, and moves on to the next record keep
                If errorReason = "" Then
                    'Check the dimension members to see if i can determine which is bad
                    For iDm = LBound(dm) To UBound(dm)
                        If dm(iDm, 3) = "" Then
                            If dm(iDm, 3) <> "CALENDAR" Then    'If a member other than time is invalid then we shouldn't run through the rest of the time columns.
                                bStopTime = True
                            End If
                            errorReason = errorReason & "Dimension: " & dm(iDm, 0) & " Dimension Member: " & dm(iDm, 1) & " is not valid" & vbCrLf
                        Else
                            If errorReason = "" Then
                                errorReason = "Bad Member " & vbCrLf
                            End If
                        End If
                    Next iDm
                    
                End If
                Call WriteException(dm, errorReason)
                errorReason = ""   'Reset it back to zero for next error
            
                If blnMod = True Then dm(gProdDimCount, 1) = ""
                
                If bStopTime = True Then   'If a member other than time is bad then we will skip all the rest of the weeks
                    GoTo StartNextWeek
                Else
                    Exit For    'It was just a time member that was bad so we will skip that member and move to the next
                End If
            Else
                If Crossing.Writeable = True Then
                    If iDataCol > UBound(dataDM, 2) Then        'If our dataDM is too small we resize it accordingly.
                        ReDim Preserve dataDM(UBound(dataDM), UBound(dataDM, 2) + 1)
                    End If
                    For index = LBound(dm, 1) To UBound(dm, 1) + 1
                        If index = UBound(dm, 1) + 1 Then
                            dataDM(index, iDataCol) = newValue
                            If Crossing.Value = "-nan(ind)" Then
                                dataDM(index + 1, iDataCol) = 0
                            Else
                                dataDM(index + 1, iDataCol) = Crossing.Value
                            End If
                                
                        Else
                            dataDM(index, iDataCol) = dm(index, 3)
                        End If
                    Next
                    
                    iDataCol = iDataCol + 1
                    If blnMod = True Then
                        dm(gProdDimCount, 1) = ""         'Clears out the current product to ensure only new products are added.
                        dm(gProdDimCount, 3) = ""
'                    Else
'                        dm(gPeriodDimCount, 3) = ""
                    End If
                Else                                                        'This handles crossings that are not writeable, and moves on to the next record
                    Call WriteException(dm, "Crossing not Writeable")
                    If blnMod = True Then dm(gProdDimCount, 1) = ""
                    Exit For
                End If
            End If
        
        Next iModMemberCount
            
        Do Until Len(Trim(Cells(iRowCount, iColumnCount + 1).Value)) > 0
            If iColumnCount > iEndColumn Then Exit Do
            iColumnCount = iColumnCount + 1
        Loop
        dm(gPeriodDimCount, 3) = ""
        newValue = Nothing
    Next iColumnCount
    
    
StartNextWeek:
Next iRowCount

Debug.Print "DM and dataDM Array Creation", Timer - rowTime

t = Timer

'Check that we have members to writeback


If iDataCol = 0 Then
    MsgBox "There are no records to write back!"
    
    Call Crossings.Clear
    cube.ClearQuery
    Erase dm()
    Erase dataDM()
    Set addin = Nothing
    Set conn = Nothing
    
    Call TurnOnEvents
    Exit Sub
End If

'As long as we have the writebackToken we can skip the other calls but it is better to get our own token id
dimArray = getArray(dm) 'This just gets the dimensions and their order.  The missing calendar value will not matter
token_body = getTokenBody(dimArray)

writebackToken = setWritebackToken(token, FolderId, planningAreaId, token_body)
p_body = getWritebackBody(dataDM, dm, writebackToken)

token = getTokenAuth()  'For some reasone the writeback likes to have a new token
writebackSuccess = postSimpleWriteback(token, p_body)       'This is the writeback call.  It will return a true or false based on it's success

If writebackSuccess = True Then     'The return is hasError.  So if it has error it would be true
    MsgBox "Writeback Failed.  Please contact your SAS administrator or Check the Zupload File and try again"
    'We could write out the failure response code to a page.
End If

'Clear out the crossings and Array's just to clear up memory
Call Crossings.Clear
cube.ClearQuery
Erase dm()
Erase dataDM()
Erase dimArray()

Debug.Print "MainWriteback Complete", Timer - t
t = Timer

addin.RefreshAll            'Refresh all tables in the FM workbook

Debug.Print "RefreshAll", Timer - t
Debug.Print "Total Run Time", Timer - tTime
MsgBox "Writeback Complete at: " & Now()        'Testing only.  Remove for production

errorFail:
'Turn events back on
Call TurnOnEvents
Set addin = Nothing
Set conn = Nothing
gVirtualCubeId = ""
gFolderHeader = ""
planningAreaId = ""
FolderId = ""
gFormsetId = ""


End Sub






