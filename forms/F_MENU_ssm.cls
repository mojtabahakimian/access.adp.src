Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command24_Click()
    Dim rst As New ADODB.Recordset, SDRST As New ADODB.Recordset, SND As Long, JAMBED As Double
                On Error Resume Next
    DoCmd.OpenForm "BUN"
    If Not IsNull(Me.MMO) Then
        rst.Open ("SELECT     SCAID, CUST_NO, SCADATE, SCAJAMF, SCARASFR, SCARASTAV, SCARASPAY, SCATAFAVOT, SCAMABNADATE, SCAMABLA, SCAMABP, CTIM, USERCO, Month FROM dbo.SSM_CUST_ANALYS WHERE ((Month = " & Me.MMO & ")  AND (SCATAFAVOT > " & Forms![BASEKNOW]![SSMTRTAKM] & ")) OR ((MONTH = " & Me.MMO & ") AND (SCATAFAVOT <" & Forms![BASEKNOW]![SSMTRTAGM] * -1 & "))"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            SND = Createsanad(Forms![BASEKNOW]![YEA] & RIGHT("00" & Me.MMO, 2) & "30", "تاخير و تعجيل در پرداخت مشتريان در ماه " & Me.MMO, False, 0, True, UCurrentUser())
            DoCmd.OpenForm "BUN"
            SDRST.Open "SELECT * FROM deed_DTL WHERE N_S = " & SND, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst.EOF
                Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
                Call GETTAF3(rst.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                If rst.Fields("SCAMABLA") <> 0 Then
                    SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
                    SDRST.Fields("N_S") = SND
                    SDRST.Fields("HES_K") = CKOL
                    SDRST.Fields("HES_M") = CMOIN
                    SDRST.Fields("HES_T") = CTAF
                    SDRST.Fields("HES_T2") = CTAF2
                    SDRST.Fields("HES_T3") = CTAF3
                    SDRST.Fields("HES_T4") = CTAF4
                    SDRST.Fields("hes") = rst.Fields("CUST_NO")
                    If rst.Fields("SCAMABLA") > 0 Then
                        SDRST.Fields("SHARH") = Left("تاخير پرداخت به مدت " & rst.Fields("SCATAFAVOT") & " روز راس فاکتورها: " & rst.Fields("SCARASFR") & " راس پرداختها: " & rst.Fields("SCARASPAY") & " راس توافق شده: " & rst.Fields("SCARASTAV") & " تاريخ مبنا:" & Format(rst.Fields("SCAMABNADATE"), "####/##/##") & " جمع فاکتورها:" & Format(rst.Fields("SCAJAMF"), "#,##0") & " درصد کاهش تخفيف روزانه :" & rst.Fields("SCAMABP"), 255)
                        SDRST.Fields("BED") = rst.Fields("SCAMABLA")
                    Else
                        SDRST.Fields("SHARH") = Left("تعجيل پرداخت به مدت " & Abs(rst.Fields("SCATAFAVOT")) & " روز راس فاکتورها: " & rst.Fields("SCARASFR") & " راس پرداختها: " & rst.Fields("SCARASPAY") & " راس توافق شده: " & rst.Fields("SCARASTAV") & " تاريخ مبنا:" & Format(rst.Fields("SCAMABNADATE"), "####/##/##") & " جمع فاکتورها:" & Format(rst.Fields("SCAJAMF"), "#,##0") & " درصد افزايش تخفيف روزانه :" & rst.Fields("SCAMABP"), 255)
                        SDRST.Fields("BES") = Abs(rst.Fields("SCAMABLA"))
                    End If
                    JAMBED = JAMBED + rst.Fields("SCAMABLA")
                    SDRST.update
                End If
                rst.MoveNext
                Forms![BUN].[num].CAPTION = GETTAFNAME(rst.Fields("cust_no"))
                Forms![BUN].Form.Refresh
            Wend
            Call GETTAF3(Forms![BASEKNOW]![HDARKASRTAKHF], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
            SDRST.AddNew
            SDRST.Fields("N_S") = SND
            SDRST.Fields("HES_K") = CKOL
            SDRST.Fields("HES_M") = CMOIN
            SDRST.Fields("HES_T") = CTAF
            SDRST.Fields("HES_T2") = CTAF2
            SDRST.Fields("HES_T3") = CTAF3
            SDRST.Fields("HES_T4") = CTAF4
            SDRST.Fields("hes") = Forms![BASEKNOW]![HDARKASRTAKHF]
            If JAMBED < 0 Then
                SDRST.Fields("SHARH") = Left(" تعجيل پرداخت ماه " & Me.MMO, 255)
                SDRST.Fields("BED") = Abs(JAMBED)
            Else
                SDRST.Fields("SHARH") = Left(" تاخير پرداخت ماه " & Me.MMO, 255)
                SDRST.Fields("BES") = Abs(JAMBED)
            End If
            SDRST.update
            DoCmd.OpenForm "MESAG", , , , , acDialog, "سند صادر شد"
            DoCmd.Close acForm, "BUN"
            DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & SND, , acDialog
        End If
    End If
    DoCmd.Close acForm, "BUN"
 End Sub

Private Sub Command5_Click()
            On Error Resume Next
Dim rst As New ADODB.Recordset, i As Integer
If IsNull(Me.MMO) Then
  DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "ماه را انتخاب کنيد..."
Else
    If Me.SSM_CUST_ANALYS_form.Form.RecordsetClone.RecordCount = 0 Then
        If UMonth(FARSIDATE(DATE)) > Me.MMO + IIf(IsNull(Forms![BASEKNOW]![SSMTBMON]), 0, Forms![BASEKNOW]![SSMTBMON]) Then
            DoCmd.OpenForm "bun"
            rst.Open "SELECT     CUST_NO FROM         dbo.HEAD_LST WHERE     (TAG = 13) AND (dbo.Umonth(DATE_N) = " & Me.MMO & ") GROUP BY CUST_NO", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                While Not rst.EOF
                    Forms![BUN].[num].CAPTION = GETTAFNAME(rst.Fields("cust_no"))
                    Forms![BUN].Form.Refresh
                    i = HesAnalysis(rst.Fields("cust_no"), Me.MMO)
                    rst.MoveNext
                Wend
            Else
               DoCmd.Close acForm, "bun"
               DoCmd.OpenForm "MESAG", , , , , acDialog, "براي اين ماه فروشي ثبت نشده"
               Exit Sub
             '  DoCmd.Close acForm, Me.NAME
            End If
            DoCmd.Close acForm, "bun"
        Else
          DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "با توجه به تنضيمات براي ماه مورد نظر قابل محاسبه نيست حداقل  بايد " & Forms![BASEKNOW]![SSMTBMON] & "  ماه با تاخير محاسبه گردد"
        End If
    Else
      DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اول محاسبات قبلي را حذف کنيد سپس تاييد را بزنيد "
    End If
End If
Me.SSM_CUST_ANALYS_form.Requery
End Sub

Private Sub Command6_Click()
                     On Error Resume Next
   DoCmd.Close
End Sub


Private Sub ESLAH_Click()
  Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.MMO) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("SSM_CUST_ANALYS", "(MONTH = " & Me.MMO & ")", dt, 1)
    End If
    Me.SSM_CUST_ANALYS_form.Form.AllowDeletions = True
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MMO_AfterUpdate()
        Me.SSM_CUST_ANALYS_form.Form.AllowDeletions = False
End Sub
