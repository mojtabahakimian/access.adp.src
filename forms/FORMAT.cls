Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command123_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "CHK_DATEN"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Command84_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "CHK_DATE"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Command85_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "CHK_ALPHA"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Command86_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "CHK_VAJH"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Command87_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "CHK_NUM"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Command88_Click()
                      On Error Resume Next
   DoCmd.GoToRecord acActiveDataObject, , acNext
   DoCmd.GoToRecord acActiveDataObject, , acLast
End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
   If IsLoadedR("TEJA50") Then
      DoCmd.Close acReport, "TEJA50"
      DoCmd.OpenReport "TEJA50", acViewPreview
   Else
      If IsLoadedR("CHAPCHEK") Then
            DoCmd.Close acReport, "CHAPCHEK"
            DoCmd.OpenReport "CHAPCHEK", acViewPreview
      End If
   End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   Me.DATE_COLOR.BackColor = Me.DATE_COLOR
   Me.ALPHA_COLOR.BackColor = Me.ALPHA_COLOR
   Me.VAJH_COLOR.BackColor = Me.VAJH_COLOR
   Me.NUM_COLOR.BackColor = Me.NUM_COLOR
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
        Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
   If IsLoadedR("TEJA50") Then
        rst.Open "SELECT COD_HESAB.BANK, CHECK_BOOK.KIND FROM COD_HESAB INNER JOIN CHECK_BOOK ON COD_HESAB.CODE = CHECK_BOOK.COD_HES WHERE (((COD_HESAB.MOIN) = " & [Forms]![HED_SODUR]![CHK_SODUR].[Form]![N_TAF] & "))GROUP BY COD_HESAB.BANK, CHECK_BOOK.KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        COBANK = rst.Fields("BANK")
        COKIND = rst.Fields("KIND")
        Me.Filter = "((BANK = " & COBANK & ") AND KIND = " & COKIND & ")"
        Me.FilterOn = True
        Me.Command88.Visible = True
   Else
        If IsLoadedR("CHAPCHEK") Then
             COBANK = [Forms]![CHAPCHEK]![MOIN]
             COKIND = [Forms]![CHAPCHEK]![kindp]
             Me.Filter = "((BANK = " & COBANK & ") AND KIND = " & COKIND & ")"
             Me.FilterOn = True
             Me.Command88.Visible = True
        End If
   End If
       Dim wd As Object, fontID As Variant

    Set wd = CreateObject("Word.Application")

    For Each fontID In wd.FontNames
        DATE_FONT.AddItem fontID
        DATE_FONTN.AddItem fontID
        ALPHA_FONT.AddItem fontID
        VAJH_FONT.AddItem fontID
        NUM_FONT.AddItem fontID

'        Sheet1.cmbFonts.AddItem fontID
    Next
    wd.Quit
    Set wd = Nothing

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
