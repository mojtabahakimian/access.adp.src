Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
            On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
              On Error Resume Next
     Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, JST As New ADODB.Recordset, JAMF As Double, i  As Integer, jammoin As Double, jammointkh As Double
     i = 0
     jammoin = 0
     jammointkh = 0
     If KeyAscii = 27 Then
       DoCmd.Close
     End If
     If KeyAscii = 13 Then
        If IsLoaded("HEAD_LST_FROOSH22") Then
            JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Forms![HEAD_LST_FROOSH22]![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
                JAMF = JST.Fields(0)
                If RIGHT(Me.Text0, 1) = "%" Then
                  Me.Text0 = Round(JAMF * Val(Left(Me.Text0, Len(Me.Text0) - 1)) / 100)
                End If
                Forms![HEAD_LST_FROOSH22]![TAKHFIF] = Me.Text0
                rst.Open "select * from invo_LST WHERE TAG = 2 AND NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    i = i + 1
                    If i = rst.RecordCount Then
                        rst.Fields("N_MOIN") = Me.Text0 - jammoin
                        If rst.Fields("MABL_K") <> 0 Then
                            rst.Fields("N_KOL") = (Me.Text0 - jammoin) / rst.Fields("MABL_K") * 100
                        Else
                            rst.Fields("N_KOL") = 0
                        End If
                    Else
                        rst.Fields("N_MOIN") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0)
                        If rst.Fields("MABL_K") <> 0 Then
                            rst.Fields("N_KOL") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0) / rst.Fields("MABL_K") * 100
                        Else
                            rst.Fields("N_KOL") = 0
                        End If
                    End If
                    jammoin = jammoin + Round(rst.Fields("MABL_K") / JAMF * Me.Text0)
                    jammointkh = jammointkh + Round(rst.Fields("MABL_K") / JAMF * Me.Text0) + Round((rst.Fields("MABL_K") - Round(rst.Fields("MABL_K") / JAMF * Me.Text0)) * rst.Fields("TKHN") / 100)
                    If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
                       RST2.Open "select CMBAA ,code from STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If RST2.RecordCount > 0 Then
                           If RST2.Fields("CMBAA") Then
                               If rst.Fields("IMBAA") <> Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100) Then
                                    rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100)
                               End If
                           Else
                              If rst.Fields("IMBAA") <> 0 Then
                                   Forms![BASEKNOW]![Text44] = False
                                   DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                                   If Forms![BASEKNOW]![Text44] Then
                                       rst.Fields("IMBAA") = 0
                                   End If
                              End If
                           End If
                       End If
                       Set RST2 = New ADODB.Recordset
                    End If
                    rst.update
                    rst.MoveNext
                Wend
            Else
               JAMF = 0
            End If
            Set RST2 = New ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.Open "select sum(N_MOIN) as moin,sum(IMBAA)  as mbaa  from invo_lst  where tag=2 and number  = " & Forms![HEAD_LST_FROOSH22]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            RST2.Open "select TAKHFIF,MBAA from head_lst  where tag=13 and number  = " & Forms![HEAD_LST_FROOSH22]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 1 Then
                If IsNull(rst.Fields("moin")) Then
                    RST2.Fields("TAKHFIF") = 0
                    RST2.Fields("MBAA") = 0
                Else
                    RST2.Fields("TAKHFIF") = rst.Fields("moin")
                    RST2.Fields("MBAA") = rst.Fields("MBAA")
                End If
                RST2.update
                RST2.Close
            End If
            Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Requery
        Else
          If IsLoaded("head_lst_froosh2") Then
            JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Forms![HEAD_LST_FROOSH2]![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
               JAMF = JST.Fields(0)
                If RIGHT(Me.Text0, 1) = "%" Then
                  Me.Text0 = Round(JAMF * Val(Left(Me.Text0, Len(Me.Text0) - 1)) / 100)
                End If
                Forms![HEAD_LST_FROOSH2]![TAKHFIF] = Me.Text0
                rst.Open "select * from invo_LST WHERE TAG = 2 AND NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    rst.Fields("N_MOIN") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0)
                    If rst.Fields("MABL_K") <> 0 Then
                        rst.Fields("N_KOL") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0) / rst.Fields("MABL_K") * 100
                    Else
                        rst.Fields("N_KOL") = 0
                    End If
                    If Forms![HEAD_LST_FROOSH2]![TICMBAA] Then
                       RST2.Open "select CMBAA ,code from STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If RST2.RecordCount > 0 Then
                           If RST2.Fields("CMBAA") Then
                               If rst.Fields("IMBAA") <> Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100) Then
                                    rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100)
                               End If
                           Else
                              If rst.Fields("IMBAA") <> 0 Then
                                   Forms![BASEKNOW]![Text44] = False
                                   DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                                   If Forms![BASEKNOW]![Text44] Then
                                       rst.Fields("IMBAA") = 0
                                   End If
                              End If
                           End If
                       End If
                       Set RST2 = New ADODB.Recordset
                    End If
                    rst.update
                    rst.MoveNext
                Wend
            Else
               JAMF = 0
            End If
            Set RST2 = New ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.Open "select sum(N_MOIN) as moin,sum(IMBAA)  as mbaa  from invo_lst  where tag=2 and number  = " & Forms![HEAD_LST_FROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            RST2.Open "select TAKHFIF,MBAA from head_lst  where tag=13 and number  = " & Forms![HEAD_LST_FROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 1 Then
                If IsNull(rst.Fields("moin")) Then
                    RST2.Fields("TAKHFIF") = 0
                    RST2.Fields("MBAA") = 0
                Else
                    RST2.Fields("TAKHFIF") = rst.Fields("moin")
                    RST2.Fields("MBAA") = rst.Fields("MBAA")
                End If
                RST2.update
                RST2.Close
            End If
            Forms![HEAD_LST_FROOSH2]![INVO_LST_sub].Requery
        Else
          If IsLoaded("HEAD_LST_PISHFROOSH2") Then
            JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & " ) AND ((INVO_LST.TAG)=20))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
               JAMF = JST.Fields(0)
                If RIGHT(Me.Text0, 1) = "%" Then
                  Me.Text0 = Round(JAMF * Val(Left(Me.Text0, Len(Me.Text0) - 1)) / 100)
                End If
                Forms![HEAD_LST_PISHFROOSH2]![TAKHFIF] = Me.Text0
                rst.Open "select * from invo_LST WHERE TAG = 20 AND NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    rst.Fields("N_MOIN") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0)
                    If rst.Fields("MABL_K") <> 0 Then
                        rst.Fields("N_KOL") = Round(rst.Fields("MABL_K") / JAMF * Me.Text0) / rst.Fields("MABL_K") * 100
                    Else
                        rst.Fields("N_KOL") = 0
                    End If
                    If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
                       RST2.Open "select CMBAA ,code from STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If RST2.RecordCount > 0 Then
                           If RST2.Fields("CMBAA") Then
                               If rst.Fields("IMBAA") <> Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100) Then
                                    rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(RST2.Fields("code")) / 100)
                               End If
                           Else
                              If rst.Fields("IMBAA") <> 0 Then
                                   Forms![BASEKNOW]![Text44] = False
                                   DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                                   If Forms![BASEKNOW]![Text44] Then
                                       rst.Fields("IMBAA") = 0
                                   End If
                              End If
                           End If
                       End If
                       Set RST2 = New ADODB.Recordset
                    End If
                    rst.update
                    rst.MoveNext
                Wend
            Else
               JAMF = 0
            End If
            Set RST2 = New ADODB.Recordset
            Set rst = New ADODB.Recordset
            rst.Open "select sum(N_MOIN) as moin,sum(IMBAA)  as mbaa  from invo_lst  where tag=20 and number  = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            RST2.Open "select TAKHFIF,MBAA from head_lst  where tag=20 and number  = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 1 Then
                If IsNull(rst.Fields("moin")) Then
                    RST2.Fields("TAKHFIF") = 0
                    RST2.Fields("MBAA") = 0
                Else
                    RST2.Fields("TAKHFIF") = rst.Fields("moin")
                    RST2.Fields("MBAA") = rst.Fields("MBAA")
                End If
                RST2.update
                RST2.Close
            End If
            Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_sub].Requery
        End If
        End If
        End If
       DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   If LETSGO("TFTMLOCK") Then
    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ستون تخفيفات قفل است و شما نمي توانيد تخفيف بزنيد"
    DoCmd.Close acForm, Me.NAME

  End If
If MID(Forms![BASEKNOW]![OPTIONSS], 47, 1) = "5" Then
    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "باتوجه  به فعال بودن ستون تخفيفات نقدي نميتوانيد تخفيف کلي بزنيد"
    DoCmd.Close acForm, Me.NAME
End If

If IsLoaded("HEAD_LST_FROOSH22") Then
    If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
        DoCmd.Close acForm, Me.NAME
    End If
End If
If IsLoaded("HEAD_LST_FROOSH2") Then
    If Not Forms![HEAD_LST_FROOSH2].Form.AllowEdits Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
        DoCmd.Close acForm, Me.NAME
    End If
End If
End Sub

Private Sub Text0_KeyPress(KeyAscii As Integer)
                   On Error Resume Next
 sendkeyu (KeyAscii)
End Sub
Private Sub Form_Load()
                   On Error Resume Next
  If Forms![BASEKNOW]![TKHF] = 2 Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "تخفيفات مصوب بصورت دستي قابل تغيير نيست"
     DoCmd.Close acForm, Me.NAME
  End If
End Sub
