Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a, textm, CPODEP, NHR, TAGH, i, K
If Me.BANK = 6 Then
   rst.Open "SELECT TEJ_C FROM TCOD_BANKS WHERE CODE = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 1 Then
      TEJA = rst.Fields("TEJ_C")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.MABL, " & TEJA & " AS tej, dbo.PAY_GETD.LIST_NO, dbo.PAY_GETD.BANK AS bkoh, dbo.PAY_GETD.N_SERI , dbo.COD_HESAB.CODE, dbo.COD_HESAB.SHOBEH, dbo.CHRE_LSPH.RADIF FROM         dbo.PAY_GETD INNER JOIN  dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN RIGHT OUTER JOIN  dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND   dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S WHERE     (dbo.CHRE_LSPH.RADIF = " & Forms![VBP_CHECK]![LIST_NO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set a = fs.CREATETextFile(Me.DBPATH, True)
   While Not rst.EOF()
        RST2.Open "SELECT TEJ_C FROM TCOD_BANKS WHERE CODE = " & rst.Fields("bkoh"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount = 1 Then
           TEJA = RST2.Fields("TEJ_C")
        End If
        RST2.Close
        Set RST2 = New ADODB.Recordset
        a.writeline (RIGHT("000000" & Trim(rst.Fields("SHOBEH")), 6) & RIGHT("0000000000" & rst.Fields("CODE"), 10) & RIGHT("000000" & rst.Fields("N_SERI"), 6) & RIGHT("00" & TEJA, 2) & RIGHT("00000000" & rst.Fields("LIST_NO"), 8) & RIGHT("00" & rst.Fields("TEJ"), 2) & RIGHT("0000000000000" & rst.Fields("MABL"), 13) & RIGHT("000000" & rst.Fields("DATE_S"), 6) & RIGHT("0000000000" & rst.Fields("N_HESAB"), 10))
        rst.MoveNext
   Wend
  DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال اطلاعات و تهيه فايل  با موفقيت انجام شد "
End If
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
