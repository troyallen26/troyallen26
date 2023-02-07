Attribute VB_Name = "m_Menu"
Option Explicit

Sub AddMenus()

'************************************************************************************
' Add's the Zupload to the menu list when users right click in the worksheets
' Sub Menu's are added to the zupload selection for users to choose which adjustment
' key figure they want to add the adjustments too
' 8/14/2015
'************************************************************************************
Dim cButMainMenu As CommandBar
Dim cBut, cBut1 As CommandBarControl
Dim DebugFlag As String
Dim ssMenu As Object
Dim SS As CommandBarControl
Dim x As String
Dim bAFH As Boolean
Dim bCAN As Boolean
Dim strWBName As String
 

 'Code to check that excel is connected ot the SAS add-in for excel, if not connected already a connection is made
' If addin Is Nothing Then
'    Set conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
'    Set addin = conn.FMAddIn
' End If

strWBName = ActiveWorkbook.Name

If InStr(1, strWBName, "AFH") > 0 Then
    bAFH = True
Else
    bAFH = False
End If

On Error Resume Next

'Checks to see if the Z Upload menu has been created and if so delete (we add it back in next step, this just makes sure we don't have multiple instances of the menu item
For Each SS In Application.CommandBars("cell").Controls
    If SS.Caption = "Z Upload" Then SS.Delete
    If SS.Caption = "Write Weeks for Zupload" Then SS.Delete
Next SS

'create the z upload menu option
Set cButMainMenu = Application.CommandBars("Cell")
Set cBut = cButMainMenu.Controls.Add(Type:=msoControlPopup)
Set cBut1 = cButMainMenu.Controls.Add(Type:=msoControlPopup)

cBut1.Caption = "Write Weeks for Zupload"
cBut.Caption = "Z Upload"

'Create the Write Weeks sub menu option
With cBut1.Controls.Add(Type:=msoControlButton)
    .Caption = "Write Weeks"
    .FaceId = 190
    .OnAction = "WriteWeeks"
End With

'Create the Z upload sub menu options (these are the writeable adjustment key figures)
With cBut.Controls.Add(Type:=msoControlButton)
    .Caption = "ZDPCAUF1 (Innovation / PWR)"
    .FaceId = 190
    .OnAction = "Caller1"
End With

'   With cBut.Controls.Add(Type:=msoControlButton)
'    .Caption = "ZDPEXKF1 (Sales Adjustment)"
'    .FaceId = 190
'    .OnAction = "Caller3"
'   End With

'   With cBut.Controls.Add(Type:=msoControlButton)
'    .Caption = "ZDPMKTFC (Marketing Adjustment)"
'    .FaceId = 190
'    .OnAction = "Caller4"
'   End With

'   With cBut.Controls.Add(Type:=msoControlButton)
'    .Caption = "ZDPCAUF2 (Cannibalization Adjustment)"
'    .FaceId = 190
'    .OnAction = "Caller5"
'   End With

With cBut.Controls.Add(Type:=msoControlButton)
    .Caption = "ZDPCAUF4 (Other Adjustment)"
    .FaceId = 190
    .OnAction = "Caller6"
End With

With cBut.Controls.Add(Type:=msoControlButton)
    .Caption = "ZDPSTFC3 (Planner Adjustment)"
    .FaceId = 190
    .OnAction = "Caller7"
End With

With cBut.Controls.Add(Type:=msoControlButton)
    .Caption = "ZDPFSACC (All In Stat Accepted)"
    .FaceId = 190
    .OnAction = "Caller9"
End With


If bAFH = True Then
    With cBut.Controls.Add(Type:=msoControlButton)
        .Caption = "ZDPCUSTF (Customer Forecast)"
        .FaceId = 190
        .OnAction = "Caller8"
    End With
End If

On Error GoTo 0
End Sub
 
Sub DeleteMenu()
'*************************************************
' Delete the z-upload menu from the command bar
'*************************************************
On Error Resume Next
Application.CommandBars("Cell").Controls("Z Upload").Delete
Application.CommandBars("Cell").Controls("Write Weeks for Zupload").Delete
On Error GoTo 0

End Sub



