Attribute VB_Name = "m_getMaterialGroupMembers"
Option Explicit

Function ModCollSplit(myCollection As Collection, matGrpKey) As Variant()
'##################################################################################################################
'Pass in Material Group and this function will use that key to find the Material Group Members from the collection
'and break them out along with their values into an array that will be sent back to the caller
'##################################################################################################################
Dim MaterialGroupArray() As Variant, iCollCount As Integer, prodValue As Variant, prodSplit, iProdCount

iCollCount = myCollection.Count - 1
ReDim MaterialGroupArray(iCollCount, 2)

prodSplit = Split(myCollection(matGrpKey), "_")     'Break out all the products for the given material group

ReDim MaterialGroupArray(UBound(prodSplit), 1)
For iProdCount = 0 To UBound(prodSplit)
    prodValue = Split(prodSplit(iProdCount), ";")
    'Add the product and it's value to the Array
    MaterialGroupArray(iProdCount, 0) = prodValue(0)    'Product
    MaterialGroupArray(iProdCount, 1) = prodValue(1)    'Comp Cases per Header
Next iProdCount

ModCollSplit = MaterialGroupArray
End Function

Function Exists(coll As Collection, key As String) As Boolean

    On Error GoTo EH

    IsObject (coll.Item(key))
    
    Exists = True
EH:
End Function



