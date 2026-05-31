Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
Dim rst As New ADODB.Recordset
If IsNull(Me.MMO) Then
  DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "پارامترها كافي نيست..."
Else
  Me.MODATE = Forms![BASEKNOW]![YEA] * 10000 + Forms![F_MENU_BIM]![MMO] * 100 + 31
  rst.Open "SELECT     TENDAR FROM dbo.WORKHEAD WHERE (dbo.UMonth(WDATE) = " & Me.MMO & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
  Dim PSK1, PSM1, PST1, PST2, PST3, PST4
  Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
  Select Case Me.OpenArgs
       Case "RK":
          If rst.Fields("TENDAR") Then
              DoCmd.OpenReport "R_LIST_BIMEH1_10", acViewPreview, , "MM = " & MMO
          Else
            If Me.pap = 1 Then
              DoCmd.OpenReport "R_LIST_BIMEH1", acViewPreview, , "MM = " & MMO
            Else
              DoCmd.OpenReport "R_LIST_BIMEH", acViewPreview, , "MM = " & MMO
            End If
          End If
    Case "F1":

            DoCmd.RunSQL "DELETE FROM dbo.SALARY_BESTT "
            DoCmd.RunSQL "DELETE FROM dbo.SALARY_BEDTT "
            If IsNull(PST2) Then
                DoCmd.RunSQL "INSERT INTO dbo.SALARY_BESTT   (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, '" & UCurrentUser() & "'  AS USR FROM    dbo.SALARY_BEST( " & Forms![F_MENU_BIM]![MMO] & "," & PSK1 & "," & Forms![BASEKNOW]![YEA] & ") SALARY_BEST "
                DoCmd.RunSQL "INSERT INTO dbo.SALARY_BEDTT   (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, '" & UCurrentUser() & "'  AS USR FROM    dbo.SALARY_BED( " & Forms![F_MENU_BIM]![MMO] & "," & PSK1 & "," & Forms![BASEKNOW]![YEA] & ") SALARY_BED "
            Else
                DoCmd.RunSQL "INSERT INTO dbo.SALARY_BESTT   (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     HES_K, HES_M,  HES_T2, BED, BES, MM, DATE_S, SHR, '" & UCurrentUser() & "'  AS USR FROM    dbo.SALARY_BEST2( " & Forms![F_MENU_BIM]![MMO] & "," & Forms![BASEKNOW]![YEA] & "," & PSK1 & "," & PSM1 & "," & PST1 & ")  SALARY_BEST2 "
                DoCmd.RunSQL "INSERT INTO dbo.SALARY_BEDTT   (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     HES_K, HES_M, HES_T2, BED, BES, MM, DATE_S, SHR, '" & UCurrentUser() & "'  AS USR FROM    dbo.SALARY_BED2( " & Forms![F_MENU_BIM]![MMO] & "," & Forms![BASEKNOW]![YEA] & "," & PSK1 & "," & PSM1 & "," & PST1 & ") SALARY_BED2 "

            End If
                If IsNull(Me.PCODE) Then
                    DoCmd.OpenReport "FISH_HOGHUGH", acViewPreview
                Else
                    DoCmd.OpenReport "FISH_HOGHUGH", acViewPreview, , "CODE = " & Me.PCODE
                End If
    Case "F2":
             DoCmd.RunSQL "DELETE FROM dbo.SALARY_BESTT "
            DoCmd.RunSQL "DELETE FROM dbo.SALARY_BEDTT "
            Set rst = New ADODB.Recordset
            rst.Open "SELECT  N_S FROM dbo.WORKHEAD WHERE (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                If IsNull(PST2) Then
                    DoCmd.RunSQL "INSERT INTO dbo.SALARY_BESTT (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     TOP 100 PERCENT HES_K, HES_M, HES_T AS Expr4, BED, BES, 0 AS Expr1, 0 AS Expr2, dbo.FISHSTR(ISNULL(SHARH, ' ')) AS SHARH, ' ' AS Expr3 FROM dbo.DEED_DTL WHERE (N_S = " & rst.Fields("N_S") & ") and (HES_K = " & PSK1 & ") AND (HES_M = " & PSM1 & ") AND (BES > 0)"
                    DoCmd.RunSQL "INSERT INTO dbo.SALARY_BEDTT (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     TOP 100 PERCENT HES_K, HES_M, HES_T AS Expr4, BED, BES, 0 AS Expr1, 0 AS Expr2, dbo.FISHSTR(ISNULL(SHARH, ' ')) AS SHARH, ' ' AS Expr3 FROM dbo.DEED_DTL WHERE (N_S = " & rst.Fields("N_S") & ") and (HES_K = " & PSK1 & ") AND (HES_M = " & PSM1 & ") AND (BED > 0)"
                Else
                    DoCmd.RunSQL "INSERT INTO dbo.SALARY_BESTT (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     TOP 100 PERCENT HES_K, HES_M, HES_T2 , BED, BES, 0 AS Expr1, 0 AS Expr2, dbo.FISHSTR(ISNULL(SHARH, ' ')) AS SHARH, ' ' AS Expr3 FROM dbo.DEED_DTL WHERE (N_S = " & rst.Fields("N_S") & ") and (HES_K = " & PSK1 & ") AND (HES_M = " & PSM1 & ")AND (HES_T = " & PST1 & ") AND (BES > 0)"
                    DoCmd.RunSQL "INSERT INTO dbo.SALARY_BEDTT (HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, SHR, USR) SELECT     TOP 100 PERCENT HES_K, HES_M, HES_T2 , BED, BES, 0 AS Expr1, 0 AS Expr2, dbo.FISHSTR(ISNULL(SHARH, ' ')) AS SHARH, ' ' AS Expr3 FROM dbo.DEED_DTL WHERE (N_S = " & rst.Fields("N_S") & ") and (HES_K = " & PSK1 & ") AND (HES_M = " & PSM1 & ")AND (HES_T = " & PST1 & ")  AND (BED > 0)"
                End If
                If IsNull(Me.PCODE) Then
                    DoCmd.OpenReport "FISH_HOGHUGH2", acViewPreview
                Else
                    DoCmd.OpenReport "FISH_HOGHUGH2", acViewPreview, , "CODE = " & Me.PCODE
                End If
            End If
    Case "SA":   DoCmd.OpenReport "R_LIST_salary", acViewPreview
    Case "SA2":   DoCmd.OpenReport "R_LIST_salary2", acViewPreview
    Case "MK":   DoCmd.OpenReport "R_LIST_maliat", acViewPreview, , "MM = " & MMO
   End Select
      DoCmd.OpenForm Me.NAME, , , , , acHidden
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "براي اين ماه كاربرگ حقوق ثبت نشده است"
        DoCmd.Close acForm, Me.NAME
   End If
End If

End Sub

Private Sub Command6_Click()
                     On Error Resume Next
   DoCmd.Close
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
