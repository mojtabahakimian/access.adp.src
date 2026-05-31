Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
'Call main
PORTP = "COM9"
        If PORTP = "" Then
        Else
                Dim BB(37) As Byte, ST As String, i As Integer
                Open PORTP & ":9600,N,8,1" For Random As #1 Len = 38
                ST = "{" & """ServiceCode""" & ":" & """1""" & "," & """Amount""" & ":" & """11111""" & "}"
'                ST = "{" & """ServiceCode""" & ":" & """1""" & "," & """Amount""" & ":" & """1000""" & "," & """PayerId""" & ":" & """108""" & "," & """PcID""" & ":" & """1234""" & "}"
                For i = 1 To Len(ST)
                  BB(i - 1) = Asc(MID(ST, i, 1))
                Next i
              '  BB(0) = 48
              '  BB(1) = 48
           '     BB(2) = 56
           '     BB(3) = 49
                Put #1, , BB
                Close #1
       End If

   Call Shell("C:\S.EXE  com9 11111", vbMaximizedFocus)
End Sub
