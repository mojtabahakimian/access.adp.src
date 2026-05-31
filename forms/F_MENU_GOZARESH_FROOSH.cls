Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                 On Error Resume Next
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.SHIFT) Then
    SSHIFT = "%"
  Else
    SSHIFT = SHIFT
  End If
  If Me.OpenArgs = "FR" Then
        If IsNull(Me.USERR) Then
            Me.USERR = "%"
        End If
  End If
  If IsNull(Me.DT1) Or IsNull(Me.DT2) Or IsNull(Me.DEPART) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If

  Select Case Me.OpenArgs
    Case "F":
             If Me.SSHIFT = "%" Then
                DoCmd.OpenReport "GOZARESH_FROOSH_USER3", acViewPreview
             Else
                DoCmd.OpenReport "GOZARESH_FROOSH_USER", acViewPreview
             End If
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FR":  DoCmd.OpenReport "LIST_FROOSH_ANBARS_DTL", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "FK":
                If Me.DTL Then
                    If IsNull(Me.USERR) Then
                       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "نام کاربر وارد نشده!"
                      Exit Sub
                    End If
                     DoCmd.OpenReport "R_FROOSH_DAYLY2", acViewPreview
                Else
                     DoCmd.OpenReport "GOZARESH_FROOSH_USER3", acViewPreview
                End If

               DoCmd.OpenForm Me.NAME, , , , , acHidden
  End Select
End Sub

Private Sub Command6_Click()
                     On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub DTL_Click()
    If Me.DTL Then
        Me.SHIFT.Visible = True
        Me.USERR.Visible = True
    Else
        Me.SHIFT.Visible = False
        Me.USERR.Visible = False
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
        Forms![BASEKNOW]![hhwin] = Me.hWnd
        Me.DEPART = [Forms]![Default]![TFSAZMAN]
        If Forms![BASEKNOW]![UGRP] = 3 Then
            Me.DEPART.Enabled = False
        End If
        If Me.OpenArgs = "FK" Then
            Me.SHIFT.Visible = False
            Me.USERR.Visible = False
        End If
End Sub
