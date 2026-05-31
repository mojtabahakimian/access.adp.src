Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
 Dim cn As String, i As Integer, NCON As New ADODB.Connection
   If Not (Me.SERV = "" Or IsNull(Me.SERV)) Then
      cn = CurrentProject.BaseConnectionString

      i = InStr(1, cn, "DATA SOURCE")
    Err.Clear
      NCON.connectionString = Left(cn, i - 1) & "Data Source=" & Me.SERV & ";" & RIGHT(cn, Len(cn) - InStr(i, cn, ";"))
      NCON.Open
      If Err <> 0 Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , , acDialog, "نام وارد شده صحيح نيست يا اينكه شبكه قطع است."
      End If
   End If
   DoCmd.Close acForm, Me.NAME
 End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Quit
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
