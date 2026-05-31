Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command3_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command4_Click()
             On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer, TEDAD As Integer, rptname As String, RST2 As New ADODB.Recordset
  rst.Open "SELECT   sum(mbav) as mbb, C0 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 0 THEN mbav ELSE 0 END  )/ sum(mbav)*100, C1 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 1 THEN mbav ELSE 0 END  )/sum(mbav)*100, C2 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 2 THEN mbav ELSE 0 END  )/sum(mbav)*100, C3 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 3 THEN mbav ELSE 0 END  )/sum(mbav)*100, C4 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 4 THEN mbav ELSE 0 END  )/sum(mbav)*100, C5 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 5 THEN mbav ELSE 0 END  )/sum(mbav)*100, C6 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 6 THEN mbav ELSE 0 END  )/sum(mbav)*100, C7 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 7 THEN mbav ELSE 0 END  )/sum(mbav)*100, C8 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 8 THEN mbav ELSE 0 END  )/sum(mbav)*100, C9 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 9 THEN mbav ELSE 0 END  )/sum(mbav)*100, " & _
  " C10 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 10 THEN mbav ELSE 0 END  )/sum(mbav)*100, C11 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 11 THEN mbav ELSE 0 END  )/sum(mbav)*100, C12 =  sum( CASE (ys - yd) * 12 + (ms - md) WHEN 12 THEN mbav ELSE 0 END  )/sum(mbav)*100 FROM  (SELECT        dbo.Umonth(DATE_S) AS MS, dbo.Uyear(DATE_S) AS YS, dbo.Umonth(DATE) AS MD, dbo.Uyear(DATE) AS YD, SUM(MABL) AS mbav FROM " & Me.dbnam & ".dbo.PAY_GETD WHERE        (N_KOL <> 911 OR  N_KOL IS NULL) AND (DATE BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") GROUP BY dbo.Umonth(DATE_S), dbo.Uyear(DATE_S), dbo.Umonth(DATE), dbo.Uyear(DATE) ) as p", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Forms![buget_defaul]![BGPCH0] = rst.Fields("C0")
        Forms![buget_defaul]![BGPCH1] = rst.Fields("C1")
        Forms![buget_defaul]![BGPCH2] = rst.Fields("C2")
        Forms![buget_defaul]![BGPCH3] = rst.Fields("C3")
        Forms![buget_defaul]![BGPCH4] = rst.Fields("C4")
        Forms![buget_defaul]![BGPCH5] = rst.Fields("C5")
        Forms![buget_defaul]![BGPCH6] = rst.Fields("C6")
        Forms![buget_defaul]![BGPCH7] = rst.Fields("C7")
        Forms![buget_defaul]![BGPCH8] = rst.Fields("C8")
        Forms![buget_defaul]![BGPCH9] = rst.Fields("C9")
        Forms![buget_defaul]![BGPCH10] = rst.Fields("C10")
        Forms![buget_defaul]![BGPCH11] = rst.Fields("C11")
        Forms![buget_defaul]![BGPCH12] = rst.Fields("C12")
    Else
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(" & Me.dbnam & ".dbo.DEED_DTL.BED) AS MAB FROM         " & Me.dbnam & ".dbo.DEED_DTL INNER JOIN    " & Me.dbnam & ".dbo.DEED_HED ON " & Me.dbnam & ".dbo.DEED_DTL.N_S = " & Me.dbnam & ".dbo.DEED_HED.N_S WHERE     (" & Me.dbnam & ".dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")  AND (" & Me.dbnam & ".dbo.DEED_DTL.HES = N'" & Me.FRCHASH & " ')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MAB")) Then
        Dim CASH As Double
        CASH = rst.Fields("MAB")
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM(" & Me.dbnam & ".dbo.DEED_DTL.BES - " & Me.dbnam & ".dbo.DEED_DTL.BED) AS MAB FROM         " & Me.dbnam & ".dbo.DEED_DTL INNER JOIN     " & Me.dbnam & ".dbo.DEED_HED ON " & Me.dbnam & ".dbo.DEED_DTL.N_S = " & Me.dbnam & ".dbo.DEED_HED.N_S  WHERE     (" & Me.dbnam & ".dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (" & Me.dbnam & ".dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![FROSH] & " or " & Me.dbnam & ".dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![MFROSH] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Forms![buget_defaul]![BGPCASH] = CASH / rst.Fields("MAB") * 100
        End If
    Else
    End If
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub dt1_AfterUpdate()
    If UYear(DT1) <> Forms![BASEKNOW]![YEA] Then
        Me.dbnam.Visible = True
    Else
        Me.dbnam.Visible = False
    End If
End Sub

Private Sub dt2_AfterUpdate()
    If UYear(DT2) <> Forms![BASEKNOW]![YEA] Then
        Me.dbnam.Visible = True
    Else
        Me.dbnam.Visible = False
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

End Sub
