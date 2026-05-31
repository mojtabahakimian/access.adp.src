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
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT2) Then
     Me.DT2 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If Me.OpenArgs = "FD" Or Me.OpenArgs = "TANBGRP" Then
     If IsNull(Me.ANBAR) Then
        MANBAR = "%"
     End If
  End If
  If IsNull(Me.ANBAR) And Me.OpenArgs <> "TANBGRP" Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  Else
    MANBAR = Me.ANBAR
  End If
  Select Case Me.OpenArgs
    Case "R":
'            If IsNull(Me.ANBAR) Then
                DoCmd.OpenReport "R_TARAZ_ANBAR_KHAS", acViewPreview
'            Else
'                DoCmd.OpenReport "R_TARAZ_ANBAR_KHAS", acViewPreview, , "ANBAR = " & Me.ANBAR
'            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "F":
'            If IsNull(Me.ANBAR) Then
                DoCmd.OpenForm "TARAZ_ANBAR_KHAS", acFormDS
'            Else
'                DoCmd.OpenForm "TARAZ_ANBAR_KHAS", acFormDS, , "ANBAR = " & Me.ANBAR
'            End If
    Case "FD":
'            If IsNull(Me.ANBAR) Then
                DoCmd.OpenForm "C_TARAZ_ANBAR_KHAS", acFormDS
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "TANBGRP":
'            If IsNull(Me.ANBAR) Then
                DoCmd.OpenReport "R_TARAZ_ANBAR_KHAS_GRP", acViewPreview
            DoCmd.OpenForm Me.NAME, , , , , acHidden
  End Select
            DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Command6_Click()

                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  If Me.OpenArgs = "TANBGRP" Then
     Me.ANBAR.Visible = False
     Me.CANBAR.Visible = False
  End If
End Sub
