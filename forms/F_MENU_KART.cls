Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CANBAR_AfterUpdate()
   ANBAR = CANBAR
   Me.KALA.RowSource = "SELECT  STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE AS Expr1 FROM   STUF_DEF INNER JOIN  STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (STUF_FSK.ANBAR = " & Me.ANBAR & ")GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME"
   Me.KALA.Requery
End Sub
Private Sub ANBAR_AfterUpdate()
                     On Error Resume Next
   CANBAR = ANBAR
   Me.KALA.RowSource = "SELECT  STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE AS Expr1 FROM   STUF_DEF INNER JOIN  STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (STUF_FSK.ANBAR = " & Me.ANBAR & ")GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME"
   Me.KALA.Requery
End Sub

Private Sub Command5_Click()
                      On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT2) Or IsNull(Me.ANBAR) Then
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
    Case "R":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "R_KA_KALA", acViewPreview
            Else
                DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "KARTR2":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "R_KA_KALA2", acViewPreview
            Else
                DoCmd.OpenReport "R_KA_KALA2", acViewPreview, , "code = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "F":
            If IsNull(Me.KALA) Then
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST", acFormDS
            Else
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST", acFormDS, , "code = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
  End Select
  BringWindowToTop (Forms![BASEKNOW]![hhwin])
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
           Forms![BASEKNOW]![hhwin] = Me.hWnd
  Me.FT = CLng(Forms![BASEKNOW]![YEA] & "0101")

End Sub

Private Sub KALA_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.KALA.Text = "-" Or Me.KALA.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.KALA.Text = ""
    DoCmd.OpenForm "SERCHK2"
  Else
    Me.KALA = NewData
  End If
  Response = acDataErrContinue

End Sub
