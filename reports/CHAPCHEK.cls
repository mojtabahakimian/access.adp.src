Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Report_Close()
If Not IsLoaded("format") Then
        DoCmd.OpenForm "chapchek"
End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
'  RST.Open "SELECT COD_HESAB.BANK, CHECK_BOOK.KIND FROM COD_HESAB INNER JOIN CHECK_BOOK ON COD_HESAB.CODE = CHECK_BOOK.COD_HES WHERE (((COD_HESAB.MOIN) = " & [Forms]![CHAPCHEK]![MOIN] & "))GROUP BY COD_HESAB.BANK, CHECK_BOOK.KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'   If RST.RecordCount > 0 Then
         COBANK = [Forms]![CHAPCHEK]![MOIN]
        COKIND = [Forms]![CHAPCHEK]![kindp]
  '      RST.Close

        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM FORMAT WHERE BANK = " & COBANK & " AND KIND = " & COKIND, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.DATE.Top = rst.Fields("DATE_TOP") * 567
            Me.DATE.Left = rst.Fields("DATE_LEFT") * 567
            If Not IsNull(rst.Fields("DATE_HEIGHT")) Then
                Me.DATE.Height = rst.Fields("DATE_HEIGHT") * 567
            End If
            Me.DATE.FontName = rst.Fields("DATE_FONT")
            Me.DATE.FontSize = rst.Fields("DATE_SIZE")
            Me.DATE.FontBold = rst.Fields("DATE_B")
            Me.DATE.FontItalic = rst.Fields("DATE_I")
            Me.DATE.FontUnderline = rst.Fields("DATE_U")
            Me.DATE.ForeColor = rst.Fields("DATE_color")
            Me.DATE.Visible = rst.Fields("DATE_p")

            Me.DATEN.Top = rst.Fields("DATE_TOPN") * 567
            Me.DATEN.Left = rst.Fields("DATE_LEFTN") * 567
            If Not IsNull(rst.Fields("DATE_HEIGHTN")) Then
                Me.DATEN.Height = rst.Fields("DATE_HEIGHTN") * 567
            End If
            Me.DATEN.Width = rst.Fields("DATE_WITHN") * 567
            Me.DATEN.FontName = rst.Fields("DATE_FONTN")
            Me.DATEN.FontSize = rst.Fields("DATE_SIZEN")
            Me.DATEN.FontBold = rst.Fields("DATE_BN")
            Me.DATEN.FontItalic = rst.Fields("DATE_IN")
            Me.DATEN.FontUnderline = rst.Fields("DATE_UN")
            Me.DATEN.ForeColor = rst.Fields("DATE_colorN")
            Me.DATEN.Visible = rst.Fields("DATE_pN")
            Me.DATEN.Vertical = rst.Fields("DATE_VERN")

            Me.TEXTMABL.Top = rst.Fields("ALPHA_TOP") * 567
            Me.TEXTMABL.Left = rst.Fields("ALPHA_LEFT") * 567
            Me.TEXTMABL.FontName = rst.Fields("ALPHA_FONT")
            Me.TEXTMABL.FontSize = rst.Fields("ALPHA_SIZE")
            Me.TEXTMABL.FontBold = rst.Fields("ALPHA_B")
            Me.TEXTMABL.FontItalic = rst.Fields("ALPHA_I")
            Me.TEXTMABL.FontUnderline = rst.Fields("ALPHA_U")
            Me.TEXTMABL.ForeColor = rst.Fields("ALPHA_color")
            Me.TEXTMABL.Width = rst.Fields("ALPHA_WITH") * 567
            If Not IsNull(rst.Fields("ALPHA_HEIGHT")) Then
                Me.TEXTMABL.Height = rst.Fields("ALPHA_HEIGHT") * 567
            End If
            Me.TEXTMABL.Visible = rst.Fields("ALPHA_p")

            Me.VAJH.Top = rst.Fields("VAJH_TOP") * 567
            Me.VAJH.Left = rst.Fields("VAJH_LEFT") * 567
            Me.VAJH.FontName = rst.Fields("VAJH_FONT")
            Me.VAJH.FontSize = rst.Fields("VAJH_SIZE")
            Me.VAJH.FontBold = rst.Fields("VAJH_B")
            Me.VAJH.FontItalic = rst.Fields("VAJH_I")
            Me.VAJH.FontUnderline = rst.Fields("VAJH_U")
            Me.VAJH.ForeColor = rst.Fields("VAJH_color")
            Me.VAJH.Visible = rst.Fields("VAJH_p")
            Me.VAJH.Width = rst.Fields("VAJH_WITH") * 567
            If Not IsNull(rst.Fields("VAJH_HEIGHT")) Then
                Me.VAJH.Height = rst.Fields("VAJH_HEIGHT") * 567
            End If

            Me.mabl.Top = rst.Fields("NUM_TOP") * 567
            Me.mabl.Left = rst.Fields("NUM_LEFT") * 567
            Me.mabl.FontName = rst.Fields("NUM_FONT")
            Me.mabl.FontSize = rst.Fields("NUM_SIZE")
            Me.mabl.FontBold = rst.Fields("NUM_B")
            Me.mabl.FontItalic = rst.Fields("NUM_I")
            Me.mabl.FontUnderline = rst.Fields("NUM_U")
            Me.mabl.ForeColor = rst.Fields("NUM_color")
            Me.mabl.Visible = rst.Fields("NUM_p")
            If Not IsNull(rst.Fields("NUM_HEIGHT")) Then
                Me.mabl.Height = rst.Fields("NUM_HEIGHT") * 567
            End If

        End If
    rst.Close
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
' Else
   '     DoCmd.OpenForm "mesag", , , , , acDialog, "معين معادل حساب بانك در بخش تعريف دسته چك و حساب مشخص نشده است "
   '     DoCmd.Close acReport, Me.NAME
 '  End If
End Sub
