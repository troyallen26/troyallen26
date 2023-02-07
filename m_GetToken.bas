Attribute VB_Name = "m_GetToken"
Option Explicit

Function getTokenAuth()
Dim req As New MSXML2.ServerXMLHTTP60
Dim reqURL As String
Dim p_body As String
Dim uName, uPwd As String

uName = "func_svc_klg_zupload"
uPwd = "H%26KG%21ZODQk0Ivu3W%5E7Sc%23VpB4"
'gDeleteFileFailed = False

reqURL = "https://" & gServer & "/SASLogon/oauth/token"
req.Open "POST", reqURL, False
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.setRequestHeader "Authorization", "Basic c2FzLmVjOg=="

p_body = "grant_type=password&username=" & uName & "&password=" & uPwd

req.Send p_body

If req.Status <> 200 Then
   MsgBox "Unable to get Authorization Token.  Please contact your System Administrator"
   Exit Function
End If

'Debug.Print req.ResponseText
Dim response As Object
Set response = JsonConverter.ParseJson(req.ResponseText)

Dim token As String
getTokenAuth = response("access_token")

End Function


