Attribute VB_Name = "m_getWritebackBody"
Option Explicit

Function getWritebackBody(dataArray As Variant, dimArray As Variant, tokenValue) As String

Dim virtualCubeIdValue As String
'Dim tokenValueas As String
Dim objectIdValue As String
Dim formulaObjectIdValue As String
Dim tableIdValue As String
Dim dimensionIdValue As String
Dim memberIdValue As String
Dim currencyIdValue As String
Dim newValueValues As String
Dim formsetIdValue As String
Dim oldValuesValue As String
Dim writeToParentValue As String
Dim goalSeekValue As String
Dim indirectFormulaValue As String
Dim reconcilValue As String
Dim objectTypeNameValue As String
Dim excludedDimIdValue As String
Dim excludedMemberIdValue As String
Dim useQueryFilterValue As String
Dim formIdValue As String
Dim nullValue As String
Dim versionValue As String
Dim linksValue, fileToWrite, string1 As String



memberIdValue = Join2D(dataArray, "Data")
newValueValues = Join2D(dataArray, "New")
oldValuesValue = Join2D(dataArray, "Old")
currencyIdValue = Join2D(dataArray, "Currency")
dimensionIdValue = writeDimensionString(dimArray)

If tokenValue = "" Then
    tokenValue = "8"
End If

'Values that we get from other parts of the system
virtualCubeIdValue = gVirtualCubeId
objectIdValue = gFormsetId
formulaObjectIdValue = gFormsetId
tableIdValue = sMainTableName
formsetIdValue = gFormsetId

'Hardcoded values that shouldn't change
writeToParentValue = "false"
goalSeekValue = "false"
indirectFormulaValue = "false"
reconcilValue = "false"
objectTypeNameValue = "fms/formset"
excludedDimIdValue = ""             'This one should be blank
excludedMemberIdValue = "[]"
useQueryFilterValue = "false"
formIdValue = "null"
nullValue = "null"
versionValue = "0"
linksValue = ""

fileToWrite = "getWritebackBody.txt" 'TH - Need to put in the entire directory or in the create and write use a directory everyone can access

string1 = ""
string1 = "{" & Chr(34) & "type" & Chr(34) & ":" & Chr(34) & "simpleCubeWriteback" & Chr(34) & ","
string1 = string1 & Chr(34) & "applicationId" & Chr(34) & ":" & Chr(34) & "FM_TEMPLATE" & Chr(34) & ","
string1 = string1 & Chr(34) & "token" & Chr(34) & ":" & tokenValue & ","
string1 = string1 & Chr(34) & "writetoParentEnabled" & Chr(34) & ":" & writeToParentValue & ","
string1 = string1 & Chr(34) & "goalSeekingEnabled" & Chr(34) & ":" & goalSeekValue & ","
string1 = string1 & Chr(34) & "indirectFormulaDependancyEnabled" & Chr(34) & ":" & indirectFormulaValue & ","
string1 = string1 & Chr(34) & "reconcilliationEnabled" & Chr(34) & ":" & reconcilValue & ","
string1 = string1 & Chr(34) & "writebacks" & Chr(34) & ":" & "[{"
string1 = string1 & Chr(34) & "virtualCubeId" & Chr(34) & ":" & Chr(34) & virtualCubeIdValue & Chr(34) & ","
string1 = string1 & Chr(34) & "objectTypeName" & Chr(34) & ":" & Chr(34) & objectTypeNameValue & Chr(34) & ","
string1 = string1 & Chr(34) & "objectId" & Chr(34) & ":" & Chr(34) & objectIdValue & Chr(34) & ","
string1 = string1 & Chr(34) & "formulaObjectId" & Chr(34) & ":" & Chr(34) & "-" & formulaObjectIdValue & Chr(34) & ","
string1 = string1 & Chr(34) & "tableId" & Chr(34) & ":" & Chr(34) & tableIdValue & Chr(34) & ","
string1 = string1 & Chr(34) & "dimensionIds" & Chr(34) & ":" & dimensionIdValue & ","
string1 = string1 & Chr(34) & "excludedDimensionIds" & Chr(34) & ":[" & excludedDimIdValue & "],"
string1 = string1 & Chr(34) & "memberIds" & Chr(34) & ":[" & memberIdValue & "],"
string1 = string1 & Chr(34) & "readMemberIds" & Chr(34) & ":[" & memberIdValue & "],"
string1 = string1 & Chr(34) & "excludedMemberIds" & Chr(34) & ":[" & excludedMemberIdValue & "],"
string1 = string1 & Chr(34) & "currencyIds" & Chr(34) & ":" & currencyIdValue & ","
string1 = string1 & Chr(34) & "newValues" & Chr(34) & ":" & newValueValues & ","
string1 = string1 & Chr(34) & "useQueryFilter" & Chr(34) & ":" & useQueryFilterValue & ","
string1 = string1 & Chr(34) & "formSetId" & Chr(34) & ":" & Chr(34) & formsetIdValue & Chr(34) & ","
string1 = string1 & Chr(34) & "formId" & Chr(34) & ":" & formIdValue & ","
string1 = string1 & Chr(34) & "oldValues" & Chr(34) & ":" & oldValuesValue & "}],"
string1 = string1 & Chr(34) & "disaggregationParams" & Chr(34) & ":" & nullValue & ","
string1 = string1 & Chr(34) & "evenDisaggregationParams" & Chr(34) & ":" & nullValue & ","
string1 = string1 & Chr(34) & "proportionDisaggregationParams" & Chr(34) & ":" & nullValue & ","
string1 = string1 & Chr(34) & "reconcilliationOptions" & Chr(34) & ":" & nullValue & ","
string1 = string1 & Chr(34) & "version" & Chr(34) & ":" & versionValue & ","
string1 = string1 & Chr(34) & "indirectFormulaDependencyEnabled" & Chr(34) & ":" & indirectFormulaValue & ","
string1 = string1 & Chr(34) & "playpenName" & Chr(34) & ":" & nullValue & ","
string1 = string1 & Chr(34) & "links" & Chr(34) & ":" & linksValue & "[]}"

Call FSOCreateAndWriteToTextFile(string1, fileToWrite)
getWritebackBody = string1
string1 = ""

End Function


