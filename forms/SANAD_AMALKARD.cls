Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()
              On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, max_ns As Long, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, ii As Double
  Call GETTAF3(Me.AMAL, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  DoCmd.OpenForm "BUN"
  Set rst = New ADODB.Recordset
  rst.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("N_S") = max_ns
    rst.Fields("DATE_S") = FARSIDATE(Now())
    rst.Fields("SHARH_S") = Left(" صدور سند عملكرد " & Format(FARSIDATE(Now()), "####/##/##"), 255)
    rst.Fields("GHATEI") = 0
    rst.Fields("NO_S") = 0
    rst.Fields("OKF") = True
    rst.update
    Set rst = New ADODB.Recordset
   RST2.Open "DEED_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Open "SELECT     SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T,  dbo.DEED_DTL.HES_T2 , dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4, dbo.DEED_DTL.HES FROM         dbo.TOTA_HES INNER JOIN      dbo.DEED_DTL ON dbo.TOTA_HES.NUMBER = dbo.DEED_DTL.HES_K  WHERE (dbo.TOTA_HES.M_D = 2) GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4,  dbo.DEED_DTL.HES HAVING      (SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF()
      If rst.Fields("MAND") > 0 Then
         RST2.AddNew
         RST2.Fields("N_S") = max_ns
         RST2.Fields("HES_K") = rst.Fields("HES_K")
         RST2.Fields("HES_M") = rst.Fields("HES_M")
         RST2.Fields("HES_T") = rst.Fields("HES_T")
         RST2.Fields("HES_T2") = rst.Fields("HES_T2")
         RST2.Fields("HES_T3") = rst.Fields("HES_T3")
         RST2.Fields("HES_T4") = rst.Fields("HES_T4")
         RST2.Fields("BES") = rst.Fields("MAND")
         RST2.Fields("hes") = rst.Fields("HES")
         RST2.Fields("SHARH") = "انتقال به حساب عملكرد"
         RST2.update
         Forms![BUN].[num].CAPTION = rst.Fields("HES")
         Forms![BUN].Form.Refresh
         RST2.AddNew
         RST2.Fields("N_S") = max_ns
         RST2.Fields("HES_K") = CKOL
         RST2.Fields("HES_M") = CMOIN
         RST2.Fields("HES_T") = CTAF
         RST2.Fields("HES_T2") = CTAF2
         RST2.Fields("HES_T3") = CTAF3
         RST2.Fields("HES_T4") = CTAF4
         RST2.Fields("hes") = Me.AMAL
         RST2.Fields("BED") = rst.Fields("MAND")
         RST2.Fields("SHARH") = "انتقال از حساب " & rst.Fields("HES")
         RST2.update
       '  Forms![bun].[num].CAPTION = Me.AMAL
       '  Forms![bun].Form.Refresh
      Else
         RST2.AddNew
         RST2.Fields("N_S") = max_ns
         RST2.Fields("HES_K") = rst.Fields("HES_K")
         RST2.Fields("HES_M") = rst.Fields("HES_M")
         RST2.Fields("HES_T") = rst.Fields("HES_T")
         RST2.Fields("HES_T2") = rst.Fields("HES_T2")
         RST2.Fields("HES_T3") = rst.Fields("HES_T3")
         RST2.Fields("HES_T4") = rst.Fields("HES_T4")
         RST2.Fields("BED") = rst.Fields("MAND") * -1
         RST2.Fields("hes") = rst.Fields("HES")
         RST2.Fields("SHARH") = "انتقال به حساب عملكرد"
         RST2.update
         Forms![BUN].[num].CAPTION = rst.Fields("HES")
         Forms![BUN].Form.Refresh
         RST2.AddNew
         RST2.Fields("N_S") = max_ns
         RST2.Fields("HES_K") = CKOL
         RST2.Fields("HES_M") = CMOIN
         RST2.Fields("HES_T") = CTAF
         RST2.Fields("HES_T2") = CTAF2
         RST2.Fields("HES_T3") = CTAF3
         RST2.Fields("HES_T4") = CTAF4
         RST2.Fields("hes") = Me.AMAL
         RST2.Fields("BES") = rst.Fields("MAND") * -1
         RST2.Fields("SHARH") = "انتقال از حساب " & rst.Fields("HES_K") & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_T")
         RST2.update
       '  Forms![bun].[num].CAPTION = Me.AMAL
       '  Forms![bun].Form.Refresh
      End If
      rst.MoveNext
   Wend
   DoCmd.Close acForm, Me.NAME
   DoCmd.Close acForm, "bun"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
