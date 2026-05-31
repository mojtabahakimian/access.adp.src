Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub CANBAR_AfterUpdate()
                    On Error Resume Next
   ANBAR = CANBAR
   Me.KALA.Requery
   Me.KALA.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF INNER JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (((STUF_FSK.ANBAR)=" & [Forms]![F_MENU_ANBAR_FRKH]![ANBAR] & ")) GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME;"
End Sub
Private Sub ANBAR_AfterUpdate()
                    On Error Resume Next
   CANBAR = ANBAR
   Me.KALA.Requery
  Me.KALA.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF INNER JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (((STUF_FSK.ANBAR)=" & [Forms]![F_MENU_ANBAR_FRKH]![ANBAR] & ")) GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME;"

End Sub

Private Sub Command5_Click()
                   On Error Resume Next
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.DT2) Or IsNull(Me.DT1) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If IsNull(Me.ANBAR) Then
     MANBAR = "%"
  Else
    MANBAR = Me.ANBAR
  End If
  If IsNull(Me.KALA) Then
     MKALA = "%"
  Else
    MKALA = Me.KALA
  End If
  Select Case Me.OpenArgs
    Case "RF":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "LIST_FROOSH_ANBARS", acViewPreview
            Else
                DoCmd.OpenReport "LIST_FROOSH_ANBARS", acViewPreview, , "code = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RK":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "LIST_KHAREED_ANBARS", acViewPreview
            Else
                DoCmd.OpenReport "LIST_KHAREED_ANBARS", acViewPreview, , "code = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "F":
            If IsNull(Me.KALA) Then
                DoCmd.OpenForm "LIST_KHAREED_ANBARS", acFormDS
            Else
                DoCmd.OpenForm "LIST_KHAREED_ANBARS", acFormDS, , "code = '" & Me.KALA & "'"
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

Private Sub KALA_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.KALA.Text = "-" Or Me.KALA.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.KALA.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
     Me.KALA = NewData
  End If
  Response = acDataErrContinue
End Sub
