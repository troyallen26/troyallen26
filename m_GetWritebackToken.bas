Attribute VB_Name = "m_GetWritebackToken"
Option Explicit

Function setWritebackToken(token, FolderId, planningAreaId, p_body)

Dim req As New MSXML2.ServerXMLHTTP60
Dim reqURL, tokenId As String

tokenId = gFormsetId            'Token is typically the formsetId

reqURL = "https://" & gServer & "/planning/planningAreas/" & planningAreaId & "/folders/" & FolderId & "/tokens/" & tokenId
'Debug.Print reqURL
req.Open "POST", reqURL, False
req.setRequestHeader "Content-Type", "application/vnd.sas.planning.data.protection.query+json"
req.setRequestHeader "Accept", "application/vnd.sas.planning.token+json"
req.setRequestHeader "Authorization", "bearer " & token

req.Send p_body

If req.Status <> 201 Then   'Unable to get writeback token
   MsgBox "Unable to get Writeback Token.  Please contact System Administrator"
   Exit Function
End If

'Debug.Print req.ResponseText
Dim response As Object '
Set response = JsonConverter.ParseJson(req.ResponseText)

setWritebackToken = response("token")

End Function

Function getFolder(token, planningAreaId) As String
'MOD Naming Convention must be Zupload <Folder Name (i.e. Frozen)>-MOD
'Non Mod Naming Convention must be Zupload <Folder Name>

Dim req As New MSXML2.ServerXMLHTTP60
Dim reqURL As String
Dim p_body As String, sSearchExpression As String, dimName As String, myFolderHeader As String
Dim strFormSetName As String, myFormsetSearch As String
Dim response, dimensions, dimension, myFormsets, formset
Dim modit As Boolean

reqURL = "https://" & gServer & "/planning/planningAreas/" & planningAreaId & "/folders"
req.Open "GET", reqURL, False
req.setRequestHeader "Content-Type", "application/vnd.sas.collection+json"
req.setRequestHeader "Accept", "application/vnd.sas.collection+json"
req.setRequestHeader "Authorization", "bearer " & token

req.Send

If req.Status <> 200 Then
   MsgBox "Unable to get Folder for Writeback.  Please contact your System Administrator"
   Exit Function
End If

'Debug.Print req.ResponseText
modit = False
Set response = JsonConverter.ParseJson(req.ResponseText)

sSearchExpression = "-MOD"              'If this is a MOD then we extract the folder name further
If InStr(1, gFolderHeader, sSearchExpression, vbTextCompare) Then
    myFolderHeader = Mid(gFolderHeader, 1, InStr(1, gFolderHeader, sSearchExpression, vbTextCompare) - 1)
    'myFormsetSearch = "Zupload" & " " & Mid(gFolderHeader, 1, Len(gFolderHeader) - 5)
    myFormsetSearch = "Zupload" & " " & myFolderHeader & sSearchExpression
    modit = True
Else
    If InStr(1, gFolderHeader, "-", vbTextCompare) Then         'This would be for the the forms
        myFolderHeader = Trim(Mid(gFolderHeader, 1, InStr(1, gFolderHeader, "-", vbTextCompare) - 1))
        myFormsetSearch = "Zupload" & " " & myFolderHeader
    Else    'This would be for formset
        myFolderHeader = Trim(Mid(gFolderHeader, 1, InStr(1, gFolderHeader, ".", vbTextCompare) - 1))
        myFormsetSearch = "Zupload" & " " & myFolderHeader
    End If
End If

Set dimensions = response("items")  'This will get each member under the Items section of the return.

For Each dimension In dimensions
    dimName = dimension("name")
    If dimName = UCase(myFolderHeader) Then
    'If InStr(1, dimName, myFolderHeader, vbTextCompare) Then
        getFolder = dimension("id")
        For Each formset In dimension("formSets")
            If formset("name") = myFormsetSearch Then
            'If InStr(1, formset("name"), myFormsetSearch, vbTextCompare) Then
                gFormsetId = formset("id")
                Exit For
            End If
        Next
        Exit For
    End If
    
Next dimension
End Function

Function getPlanningAreaId(token) As String
Dim req As New MSXML2.ServerXMLHTTP60
Dim reqURL As String
Dim p_body, planningAreaName As String
Dim response, dimensions, dimension

'For the Zupload this will be in the "Detailed" Planning Area.
planningAreaName = "Detailed"

reqURL = "https://" & gServer & "/planning/planningAreas/"
req.Open "GET", reqURL, False
req.setRequestHeader "Content-Type", "application/vnd.sas.planning.planning.area+json"
req.setRequestHeader "Accept", "application/vnd.sas.collection+json"
req.setRequestHeader "Authorization", "bearer " & token

req.Send

If req.Status <> 200 Then
   MsgBox "Unable to get Planning Area Id.  Please contact your System Administrator"
   Exit Function
End If

'Debug.Print req.ResponseText

Set response = JsonConverter.ParseJson(req.ResponseText)

Set dimensions = response("items")  'This will get each member under the Items section of the return.

For Each dimension In dimensions
    If dimension("name") = planningAreaName Then
        getPlanningAreaId = dimension("id")
        Exit For
    End If
Next dimension

End Function

'Function getFormsetId(token, planningAreaId, FolderId) As String
''MOD Naming Convention must be Zupload <Folder Name (i.e. Frozen)>-MOD
''Non Mod Naming Convention must be Zupload <Folder Name>
'
'Dim req As New MSXML2.ServerXMLHTTP60
'Dim reqURL As String
'Dim p_body, sSearchExpression, dimName As String
'Dim response, dimensions, dimension
'
'reqURL = "https://" & gServer & "/planningAreas/" & planningAreaId & "/folder/" & FolderId & "/formsets"
'Debug.Print reqURL
'req.Open "GET", reqURL, False
'req.setRequestHeader "Content-Type", "application/vnd.sas.collection+json"
'req.setRequestHeader "Accept", "application/vnd.sas.collection+json"
'req.setRequestHeader "Authorization", "bearer " & token
'
'req.Send
'
'If req.Status <> 200 Then
'   MsgBox "Unable to get Formset ID for Writeback.  Please contact your System Administrator"
'   Exit Function
'End If
'
''Debug.Print req.ResponseText
'
'Set response = JsonConverter.ParseJson(req.ResponseText)
'
'sSearchExpression = "-MOD"              'If this is a MOD then we extract the folder name further
'If InStr(1, gFolderHeader, sSearchExpression, vbTextCompare) Then
'    gFolderHeader = Mid(gFolderHeader, 1, InStr(1, gFolderHeader, sSearchExpression, vbTextCompare) - 1)
'End If
'
'Set dimensions = response("items")  'This will get each member under the Items section of the return.
''Debug.Print req.ResponseText
'For Each dimension In dimensions
'    dimName = dimension("name")
'    If InStr(1, dimName, gFolderHeader, vbTextCompare) Then
'        getFormsetId = dimension("id")
'        Exit For
'    End If
'
'Next dimension
'End Function

Function postSimpleWriteback(token, p_body) As String

Dim req As New MSXML2.ServerXMLHTTP60
Dim reqURL As String
Dim response As Object
Dim dimensions, dimension
'Token is typically the formsetId

reqURL = "https://" & gServer & "/planning/writebacks/cubeWritebacks"
req.Open "POST", reqURL, False
req.setRequestHeader "Content-Type", "application/vnd.sas.planning.cube.writeback+json"
req.setRequestHeader "Accept", "application/vnd.sas.collection+json"
req.setRequestHeader "Authorization", "bearer " & token

req.Send p_body

If req.Status <> 201 Then
   MsgBox "The Writeback Failed: " & vbCrLf & "Reason Code for Failure: " & vbCrLf & req.ResponseText
   postSimpleWriteback = "True"
   Exit Function
End If

'Debug.Print req.ResponseText
Set response = JsonConverter.ParseJson(req.ResponseText)

Set dimensions = response("items")  'This will get each member under the Items section of the return.

For Each dimension In dimensions
    postSimpleWriteback = dimension("hasError")
Next

End Function


