Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
    Me.RecordSource = "SELECT     CHK_SODUR.DATE_S, CHK_SODUR.MABL, CHK_SODUR.NAME_TAH FROM CHK_SODUR WHERE     (((CHK_SODUR.N_SERI) = " & [Forms]![HED_SODUR]![CHK_SODUR].[Form]![N_SERI] & ") AND ((CHK_SODUR.N_TAF) =" & [Forms]![HED_SODUR]![CHK_SODUR].[Form]![N_TAF] & ")) GROUP BY CHK_SODUR.DATE_S, CHK_SODUR.MABL, CHK_SODUR.NAME_TAH"
    Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
   rst.Open "SELECT COD_HESAB.BANK, CHECK_BOOK.KIND FROM COD_HESAB INNER JOIN CHECK_BOOK ON COD_HESAB.CODE = CHECK_BOOK.COD_HES WHERE (((COD_HESAB.MOIN) = " & [Forms]![HED_SODUR]![CHK_SODUR].[Form]![N_TAF] & "))GROUP BY COD_HESAB.BANK, CHECK_BOOK.KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        COBANK = rst.Fields("BANK")
        COKIND = rst.Fields("KIND")
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "FORMAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = ("BANK = " & COBANK & " AND KIND = " & COKIND)
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
            Me.DATE.Top = rst.Fields("DATE_TOP") * 567
            Me.DATE.Left = rst.Fields("DATE_LEFT") * 567
            Me.DATE.FontName = rst.Fields("DATE_FONT")
            Me.DATE.FontSize = rst.Fields("DATE_SIZE")
            Me.DATE.FontBold = rst.Fields("DATE_B")
            Me.DATE.FontItalic = rst.Fields("DATE_I")
            Me.DATE.FontUnderline = rst.Fields("DATE_U")
            Me.DATE.ForeColor = rst.Fields("DATE_color")
            Me.DATE.Visible = rst.Fields("DATE_p")

            Me.TEXTMABL.Top = rst.Fields("ALPHA_TOP") * 567
            Me.TEXTMABL.Left = rst.Fields("ALPHA_LEFT") * 567
            Me.TEXTMABL.FontName = rst.Fields("ALPHA_FONT")
            Me.TEXTMABL.FontSize = rst.Fields("ALPHA_SIZE")
            Me.TEXTMABL.FontBold = rst.Fields("ALPHA_B")
            Me.TEXTMABL.FontItalic = rst.Fields("ALPHA_I")
            Me.TEXTMABL.FontUnderline = rst.Fields("ALPHA_U")
            Me.TEXTMABL.ForeColor = rst.Fields("ALPHA_color")
            Me.TEXTMABL.Width = rst.Fields("ALPHA_WITH") * 567
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

            Me.mabl.Top = rst.Fields("NUM_TOP") * 567
            Me.mabl.Left = rst.Fields("NUM_LEFT") * 567
            Me.mabl.FontName = rst.Fields("NUM_FONT")
            Me.mabl.FontSize = rst.Fields("NUM_SIZE")
            Me.mabl.FontBold = rst.Fields("NUM_B")
            Me.mabl.FontItalic = rst.Fields("NUM_I")
            Me.mabl.FontUnderline = rst.Fields("NUM_U")
            Me.mabl.ForeColor = rst.Fields("NUM_color")
            Me.mabl.Visible = rst.Fields("NUM_p")
        End If
    rst.Close
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
 Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "معين معادل حساب بانك در بخش تعريف دسته چك و حساب مشخص نشده است "
        DoCmd.Close acReport, Me.NAME
   End If
End Sub
