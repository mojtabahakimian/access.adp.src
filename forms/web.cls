Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Function http_Resp(ByVal sReq As String) As String

    Dim byteData() As Byte
    Dim XMLHTTP As Object

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")

    XMLHTTP.Open "GET", sReq, False
    XMLHTTP.Send
    byteData = XMLHTTP.ResponseBody

    Set XMLHTTP = Nothing

    http_Resp = StrConv(byteData, vbUnicode)

End Function
Public Sub Command1_Click()
    'For API
'    Dim reader As New XMLHTTP60
'    Set reader = New XMLHTTP60
'    reader.Open "GET", "http://192.168.1.210\asa\index.php", False
'    reader.setRequestHeader "Accept", "application/json"
'    reader.send
'
'
'
'    Do Until reader.readyState = 4
'        DoEvents
'    Loop

'    If reader.Status = 200 Then
'        MsgBox (reader.responseText)
'    Else
'        MsgBox "Unable to import data."
'    End If
End Sub

Private Sub Command2_Click()
    Call http_Resp("http://192.168.1.210\asa\index.php")
End Sub
