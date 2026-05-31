Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub CANBAR_AfterUpdate()
                    On Error Resume Next
   ANBAR = CANBAR
End Sub
Private Sub ANBAR_AfterUpdate()
                    On Error Resume Next
   CANBAR = ANBAR
End Sub

Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String, SHART  As String
  Dim rst, RST2 As Recordset, PATH As String, i As Integer
  If IsNull(Me.F2) Or IsNull(Me.F1) Or IsNull(Me.ANBAR) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  SHART = "(ANBAR = " & Me.ANBAR & " OR ANBAR BETWEEN " & Me.F1 & " AND " & Me.F2 & ")"
  If Me.VF <> "" Then
     For i = 1 To Len(Me.VF)
        If Trim(RIGHT(SHART, 11)) <> "or ANBAR =" Then
            SHART = SHART & " or ANBAR = "
        End If
        While IsNumeric(MID(Me.VF, i, 1))
           SHART = SHART & MID(Me.VF, i, 1)
           i = i + 1
        Wend
     Next
     If Trim(RIGHT(SHART, 11)) = "or ANBAR =" Then
            SHART = Left(SHART, Len(SHART) - 11)
     End If
  End If
  DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST_MERG", acFormDS, , , , , SHART
  DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Command6_Click()
                     On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
           Forms![BASEKNOW]![hhwin] = Me.hWnd
   If Me.OpenArgs = "RK" Then
    Me.DTT = 1
   Else
     Me.DTT = 2
   End If
End Sub
