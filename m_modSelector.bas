Attribute VB_Name = "m_modSelector"
Option Explicit
Sub modSelector(modCollection As Collection)
    'This changes the Material Group hiearchy to only those members the user is uploading. This makes the table much smaller and therefore fasteer
    Dim addin As FMAddIn
    Dim conn As Connect
    Dim table As FMTable
    Dim fm_hierarchy As FMHierarchy
    Dim cSheet As String
    Dim prodSheet As String
    Dim DefaultReadMember As FMMember
    Dim member As FMMember
    Dim fm_mod As Variant
    Dim sMainTableName As String
    Dim result As FMMember
    
    
    cSheet = "Zupload"
    prodSheet = "Product"
    
    If addin Is Nothing Then
        Set conn = Application.COMAddIns.Item("SASSESExcelAddIn.Connect").Object
        Set addin = conn.FMAddIn
    End If
    
    If addin.IsLoggedIn = False Then
        MsgBox "You must be logged in to SAS Financial Management."
        Exit Sub
    End If
    
    Sheets(prodSheet).Select
    Set table = GetTableNameFromSheet(prodSheet)

    If table Is Nothing Then
       MsgBox "There must be a table in the formset on the " & prodSheet & " tab"
       Call TurnOnEvents
       Exit Sub
    Else
       sMainTableName = table.Code
    End If
    
    Set fm_hierarchy = table.ServerHierarchies("MATERIAL_GROUP")
    
    'Reset the table so there aren't any missing members
    For Each member In fm_hierarchy.Members
        member.SelectionRule = fmSelection_NoRule
    Next
    
    If modCollection.Count > 0 Then     'Making sure there are members in the collection
    
        For Each fm_mod In modCollection    'Loop through the collection for the mods
            Set result = fm_hierarchy.Members(fm_mod)   'Checking the member agains the fm_hiearchy members for the Material Group
            If result Is Nothing Then
                'If member is bad then I already check this in the MainWriteback so i'm not goint to do anything further.
                'I keep this hear so it doesn't try and add it to the hierarchy, this just skips it
            Else
                fm_hierarchy.Members(result.Code).SelectionRule = fmSelection_Member    'Add member to the table
            End If
               
        Next fm_mod
    
    End If
    
    addin.Refresh
    Sheets(cSheet).Select
    
End Sub
