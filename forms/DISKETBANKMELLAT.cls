Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a, textm, CPODEP, NHR, TAGH, i, K, JAM
   rst.Open "SELECT MLA_C FROM TCOD_BANKS WHERE CODE = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 1 Then
      TEJA = rst.Fields("MLA_C")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   JAM = 0
   rst.Open "SELECT     dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.MABL, " & TEJA & " AS tej, dbo.PAY_GETD.LIST_NO, dbo.PAY_GETD.BANK AS bkoh, dbo.PAY_GETD.N_SERI , dbo.COD_HESAB.CODE, dbo.COD_HESAB.SHOBEH, dbo.CHRE_LSPH.RADIF FROM         dbo.PAY_GETD INNER JOIN  dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN RIGHT OUTER JOIN  dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND   dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S WHERE     (dbo.CHRE_LSPH.RADIF = " & Forms![VBP_CHECK]![LIST_NO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        JAM = JAM + rst.Fields("MABL")
        rst.MoveNext
   Wend
   rst.MoveFirst
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set a = fs.CREATETextFile(Me.DBPATH & rst.Fields("code") & ".TXT", True)
   a.writeline ("13" & FARSIDATE(Now()) & RIGHT("0000000000" & rst.Fields("code"), 10) & RIGHT("0000" & rst.RecordCount, 4))
   While Not rst.EOF()
        If IsNull(rst.Fields("bkoh")) Then
              DoCmd.OpenForm "mesag", , , , , acDialog, "در تعريف دسته چك و حساب معين معادل بانك ملي تعريف نشده است لطفا تعريف كنيد"
              Exit Sub
        End If
        RST2.Open "SELECT MLA_C FROM TCOD_BANKS WHERE CODE = " & rst.Fields("bkoh"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount = 1 Then
           TEJA = RST2.Fields("MLA_C")
        End If
        RST2.Close
        Set RST2 = New ADODB.Recordset
        a.writeline ("00" & "000000" & RIGHT("00000000" & rst.Fields("N_SERI"), 8) & RIGHT("000000000000000000" & rst.Fields("N_HESAB"), 18) & RIGHT("00" & TEJA, 2) & RIGHT("0000000" & rst.Fields("LIST_NO"), 7) & "1" & "13" & rst.Fields("DATE_S") & RIGHT("000000000000000" & rst.Fields("MABL"), 17) & "000000000000000")
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
