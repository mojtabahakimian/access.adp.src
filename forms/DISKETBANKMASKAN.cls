Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a, textm, CPODEP, NHR, TAGH, i, K
   rst.Open "SELECT MAS_C FROM TCOD_BANKS WHERE CODE = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 1 Then
      TEJA = rst.Fields("MAS_C")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.MABL, " & TEJA & " AS tej, dbo.PAY_GETD.LIST_NO, dbo.PAY_GETD.BANK AS bkoh, dbo.PAY_GETD.N_SERI , dbo.COD_HESAB.CODE, dbo.COD_HESAB.SHOBEH, dbo.CHRE_LSPH.RADIF FROM         dbo.PAY_GETD INNER JOIN  dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN RIGHT OUTER JOIN  dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND   dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S WHERE     (dbo.CHRE_LSPH.RADIF = " & Forms![VBP_CHECK]![LIST_NO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set a = fs.CREATETextFile(Me.DBPATH, True)
   a.writeline ("DC_DEPYY DC_DEPMM DC_DEPDD   DC_SER  DC_BANK DC_BRNCH   DC_ACCNO  DC_DUYY  DC_DUMM  DC_DUDD        DC_VAL  DC_MAJOR DC_MINOR DC_STAT")
  While Not rst.EOF()
        If IsNull(rst.Fields("bkoh")) Then
              DoCmd.OpenForm "mesag", , , , , acDialog, "در تعريف دسته چك و حساب معين معادل بانك مسكن تعريف نشده است لطفا تعريف كنيد"
              Exit Sub
        End If
        RST2.Open "SELECT MAS_C FROM TCOD_BANKS WHERE CODE = " & rst.Fields("bkoh"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount = 1 Then
           TEJA = RST2.Fields("MAS_C")
        End If
        RST2.Close
        Set RST2 = New ADODB.Recordset
        a.writeline (RIGHT("        " & UYear(Forms![VBP_CHECK]![DTS]), 8) & " " & RIGHT("        " & UMonth(Forms![VBP_CHECK]![DTS]), 8) & " " & RIGHT("        " & UDay(Forms![VBP_CHECK]![DTS]), 8) & " " & RIGHT("        " & rst.Fields("N_SERI"), 8) & " " & RIGHT("        " & TEJA, 8) & " " & RIGHT("        " & rst.Fields("LIST_NO"), 8) & " " & RIGHT("          " & rst.Fields("N_HESAB"), 10) & RIGHT("        " & UYear(rst.Fields("DATE_S")), 8) & " " & RIGHT("        " & UMonth(rst.Fields("DATE_S")), 8) & " " & RIGHT("        " & UDay(rst.Fields("DATE_S")), 8) & " " & RIGHT("             " & rst.Fields("MABL"), 13) & " " & RIGHT("        " & "8728", 8) & " " & RIGHT("          " & "42", 8) & " " & RIGHT("          " & "8", 10))
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

Private Sub Command165_Click()
Call RelinkTablepic
Me.Refresh

End Sub


Private Sub Text10_AfterUpdate()
End Sub
