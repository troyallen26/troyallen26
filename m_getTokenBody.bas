Attribute VB_Name = "m_getTokenBody"
Option Explicit

Function getTokenBody(dimArray As Variant)
'**************************************************
'Writes the writeback token body
'5/2/2022
'**************************************************
Dim string1 As String
Dim appTypeValue As String
Dim fileToWrite, bitMaskValue As String   'This is the file name for
Dim index As Long

appTypeValue = "3"
bitMaskValue = "32767"
fileToWrite = "getTokenRequest.txt" 'TH - Need to put in the entire directory or in the create and write use a directory everyone can access

string1 = ""
string1 = "[" & vbCrLf
string1 = string1 & "{" & vbCrLf
string1 = string1 & Chr(34) & "vCubeId" & Chr(34) & ":" & Chr(34) & gVirtualCubeId & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "appType" & Chr(34) & ":" & Chr(34) & appTypeValue & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "bitMask" & Chr(34) & ":" & Chr(34) & bitMaskValue & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "dimTypeIdsAllowNonLeafMemberInRange" & Chr(34) & ":[]," & vbCrLf
string1 = string1 & Chr(34) & "filterMemberCombinationRules" & Chr(34) & ":" & Chr(34) & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "floatingTimeEndOffset" & Chr(34) & ":" & Chr(34) & "0" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "floatingTimeOn" & Chr(34) & ":" & Chr(34) & "False" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "floatingTimeStartOffset" & Chr(34) & ":" & Chr(34) & "0" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "formSetId" & Chr(34) & ":" & Chr(34) & gFormsetId & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "hierarchyRules" & Chr(34) & ":[" & vbCrLf

For index = LBound(dimArray, 1) To UBound(dimArray, 1)
    string1 = string1 & "{" & Chr(34) & "dimensionTypeId" & Chr(34) & ":" & Chr(34) & dimArray(index) & Chr(34) & "," & vbCrLf
    string1 = string1 & Chr(34) & "includeVirtualChildren" & Chr(34) & ":" & "false" & "," & vbCrLf
    string1 = string1 & Chr(34) & "memberSelectionRuleMemberIds" & Chr(34) & ":[]," & vbCrLf
    string1 = string1 & Chr(34) & "memberPropertyFilterRules" & Chr(34) & ":[]," & vbCrLf
    string1 = string1 & Chr(34) & "memberSelectionRuleTypes" & Chr(34) & ":[]," & vbCrLf
    If index <> UBound(dimArray, 1) Then
        string1 = string1 & Chr(34) & "protectedMemberIds" & Chr(34) & ":[]}," & vbCrLf
    Else
        string1 = string1 & Chr(34) & "protectedMemberIds" & Chr(34) & ":[]}]," & vbCrLf
    End If
Next

string1 = string1 & Chr(34) & "ignoreRuleBitMask" & Chr(34) & ":" & Chr(34) & "0" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "readOnlyTable" & Chr(34) & ":" & Chr(34) & "False" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "reconcileDimTypeId" & Chr(34) & ":[]," & vbCrLf
string1 = string1 & Chr(34) & "reconcileHierarchyMaxLevel" & Chr(34) & ":[]," & vbCrLf
string1 = string1 & Chr(34) & "rollupsWriteable" & Chr(34) & ":" & Chr(34) & "False" & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "tableId" & Chr(34) & ":" & Chr(34) & sMainTableName & Chr(34) & "," & vbCrLf
string1 = string1 & Chr(34) & "analysisHierarchyRulesOnFormSet" & Chr(34) & ":{" & vbCrLf
string1 = string1 & Chr(34) & "dimensionTypeId" & Chr(34) & ":" & "null" & "," & vbCrLf
string1 = string1 & Chr(34) & "includeVirtualChildren" & Chr(34) & ":" & "false" & "," & vbCrLf
string1 = string1 & Chr(34) & "memberSelectionRuleMemberIds" & Chr(34) & ":" & "null" & "," & vbCrLf
string1 = string1 & Chr(34) & "memberPropertyFilterRules" & Chr(34) & ":" & "null" & "," & vbCrLf
string1 = string1 & Chr(34) & "memberSelectionRuleTypes" & Chr(34) & ":" & "null" & "," & vbCrLf
string1 = string1 & Chr(34) & "protectedMemberIds" & Chr(34) & ":" & "null" & "}," & vbCrLf
string1 = string1 & Chr(34) & "dimTypeInRuleToIgnore" & Chr(34) & ":[]}]" & vbCrLf


Call FSOCreateAndWriteToTextFile(string1, fileToWrite)           'Use this for testing the output

getTokenBody = string1
string1 = ""

End Function

Sub FSOCreateAndWriteToTextFile(myString, fileName)
    Dim FSO As New FileSystemObject
    Dim FiletoCreate
    Dim myDir As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    myDir = Environ$("USERPROFILE") & "\Downloads\"
    Set FiletoCreate = FSO.CreateTextFile(myDir & fileName)
 
    FiletoCreate.Write myString
    FiletoCreate.Close
 
End Sub

