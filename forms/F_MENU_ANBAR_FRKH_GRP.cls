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
  SHART = "INVO_LST.NUMBER = "
  If Me.VF <> "" Then
     For i = 1 To Len(Me.VF)
        While IsNumeric(MID(Me.VF, i, 1))
           SHART = SHART & MID(Me.VF, i, 1)
           i = i + 1
        Wend
        If Trim(RIGHT(SHART, 22)) <> "or INVO_LST.NUMBER =" Then
            SHART = SHART & " or INVO_LST.NUMBER = "
        End If
     Next
     If Trim(RIGHT(SHART, 22)) = "or INVO_LST.NUMBER =" Then
            SHART = Left(SHART, Len(SHART) - 22)
     End If
   Else
     SHART = ""
  End If
  DoCmd.OpenReport "LIST_FROOSH_ANBARS_HAVALA", acViewPreview, , , , SHART
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
