Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a, textm, CPODEP, NHR, TAGH, i, K
   rst.Open "SELECT MEL_C FROM TCOD_BANKS WHERE CODE = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 1 Then
      TEJA = rst.Fields("MEL_C")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.MABL, " & TEJA & " AS tej, dbo.PAY_GETD.LIST_NO, dbo.PAY_GETD.BANK AS bkoh, dbo.PAY_GETD.N_SERI , dbo.COD_HESAB.CODE, dbo.COD_HESAB.SHOBEH, dbo.CHRE_LSPH.RADIF FROM         dbo.PAY_GETD INNER JOIN  dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN RIGHT OUTER JOIN  dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND   dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S WHERE     (dbo.CHRE_LSPH.RADIF = " & Forms![VBP_CHECK]![LIST_NO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set a = fs.CREATETextFile(Me.DBPATH, True)
   While Not rst.EOF()
        If IsNull(rst.Fields("bkoh")) Then
              DoCmd.OpenForm "mesag", , , , , acDialog, "در تعريف دسته چك و حساب معين معادل بانك ملي تعريف نشده است لطفا تعريف كنيد"
              Exit Sub
        End If
        RST2.Open "SELECT MEL_C FROM TCOD_BANKS WHERE CODE = " & rst.Fields("bkoh"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount = 1 Then
           TEJA = RST2.Fields("MEL_C")
        End If
        RST2.Close
        Set RST2 = New ADODB.Recordset
        a.writeline ("12150101" & RIGHT("             " & rst.Fields("N_HESAB"), 13) & rst.Fields("DATE_S") & RIGHT("                 " & rst.Fields("MABL"), 17) & RIGHT("          " & rst.Fields("N_SERI"), 10) & RIGHT("  " & TEJA, 2) & "26200305" & RIGHT("     " & rst.Fields("LIST_NO"), 5) & "                    " & RIGHT("0000000000000" & rst.Fields("code"), 13) & "      ")
        rst.MoveNext
   Wend
   DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال اطلاعات و تهيه فايل  با موفقيت انجام شد "
   rst.Close
   a.Close
    DoCmd.Close acForm, "BUN"
End Sub

Private Sub Command1_Click()
    DoCmd.Close
End Sub
