Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETS As Boolean, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, Meidnum As Double
Private Sub SANAD()
                 On Error Resume Next
 Dim arst As New ADODB.Recordset
 arst.Open "SELECT  CODE, KIND FROM dbo.TCOD_ANBAR WHERE (CODE = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 If arst.RecordCount > 0 Then
    If arst.Fields("kind") = 1 Or arst.Fields("kind") = 2 Then
       If Not (Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT])) Then
            DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
            Exit Sub
       End If
    End If
 End If
 arst.Close
 If Forms![BASEKNOW]![SANAD] <> 3 Then
  DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
'  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  Err.Clear
  SHART = "NO_S = 10 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = Left(" حواله انتقالي مواد شماره " & Me.NUMBER & "-" & Me.FNUMCO & " از انبار " & Me.ANBAR & " به " & Me.ANBARF & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 10
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.update
    MABNA = SHRST.Fields("BASE")
  Else
    max_ns = SHRST.Fields("N_S")
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = Left(" حواله انتقالي مواد شماره " & Me.NUMBER & "-" & Me.FNUMCO & " از انبار " & Me.ANBAR & " به " & Me.ANBARF & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 10
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    MABNA = SHRST.Fields("BASE")
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  LETS = True
  DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
'   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  JST.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, STUF_DEF.NAME, INVO_LST.ANBAR, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.ANBARF FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
       If JST.Fields("MABL_K") <> 0 Then
        If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE")), JST.Fields(2))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![MOGODIA] & "-" & CLng(JST.Fields("ANBAR")) & "-" & CLng(JST.Fields("CODE")) & " - " & JST.Fields(2)
              LETS = False
           End If
        End If
       SDRST.AddNew 'از انبار
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
        SDRST.Fields("HES_M") = JST.Fields("ANBAR")
        SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
        SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
        SDRST.Fields("SHARH") = Left("حواله انتقالي شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
        SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
       End If
       If JST.Fields("MABL_K") <> 0 Then
        If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBARF")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBARF")), CLng(JST.Fields("CODE")), JST.Fields(2))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![MOGODIA] & "-" & CLng(JST.Fields("ANBARF")) & "-" & CLng(JST.Fields("CODE")) & " - " & JST.Fields(2)
              LETS = False
           End If
        End If
        SDRST.AddNew 'به انبار
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
        SDRST.Fields("HES_M") = JST.Fields("ANBARF")
        SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
        SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBARF") & "-" & CLng(JST.Fields("code"))
        SDRST.Fields("SHARH") = Left("حواله انتقالي شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & "  بابت " & JST.Fields("name"), 255)
        SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
       End If
        JST.MoveNext
  Wend
  If Not LETS Then
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
  End If
  End If
End Sub
Private Sub ANBAR_AfterUpdate()
                      On Error Resume Next
    Me.ANBARF.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) and (TCOD_ANBAR.CODE <> " & Me.ANBAR & ")  ORDER BY TCOD_ANBAR.CODE"
    Me.ANBARF.Requery
End Sub

Private Sub ANBAR_BeforeUpdate(CANCEL As Integer)
                          On Error Resume Next
      Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, min As Double
  rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, INVO_LST.ANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.CUST_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "برگه داراي سطر كالا ميباشد نميتوانيد انبار را تغيير دهيد اول كالاهها را پاك كنيد"
        CANCEL = True
        GoTo ED
   End If

    If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
       Set rst = New ADODB.Recordset
        rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, INVO_LST.ANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.CUST_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            RST2.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst3.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst.EOF()
                RST2.Filter = "code = '" & rst.Fields("code") & "' and anbar = " & Me.ANBAR
                Set RST2 = RST2
                If RST2.RecordCount = 0 Then
                    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كالا متعلق به انبار انتخاب شده نيست ابتدا در بخش تعريف كالا تعلق كالا به انبار را ايجاد كنيد.كد كالا :" & rst.Fields("CODE")
                    CANCEL = True
                    GoTo ED
                End If
                rst3.Filter = "CODE = '" & rst.Fields("CODE") & "'"
                Set rst3 = rst3
                If rst3.RecordCount = 0 Then
                Else
                   min = rst3.Fields("MIN_M")
                End If
                RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & Me.ANBAR
                Set RST2 = RST2
                If RST2.RecordCount = 0 Then
                    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                    CANCEL = True
                    GoTo ED
                Else
                   If (RST2.Fields("MOGODI") + RST2.Fields("MOGODI_A")) - rst.Fields("MEGHk") < min And Me.ANBAR > 0 Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "انتقال  كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "كد كالا : " & rst.Fields("code") & "حداقل موجودي تعريف شده در اف دو :" & min
                        CANCEL = True
                        GoTo ED
                   End If
                End If
                rst.MoveNext
            Wend
            rst.MoveFirst
            While Not rst.EOF()
                 rst3.Filter = "CODE = '" & rst.Fields("CODE") & "'"
                 Set rst3 = rst3
                 If rst3.RecordCount = 0 Then
                 Else
                    min = rst3.Fields("MIN_M")
                 End If
                 RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & Me.ANBAR
                 Set RST2 = RST2
                 If RST2.RecordCount = 0 Then
                    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                    CANCEL = True
                    GoTo ED
                 Else

                    RST2.Fields("MOGODI") = RST2.Fields("MOGODI") - rst.Fields("MEGHk")
                    RST2.update
    '//////////////////////////////////////////////////////////////////////////////////////
                    RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & rst.Fields("anbar")
                    Set RST2 = RST2
                    If RST2.RecordCount = 0 Then
                        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                        CANCEL = True
                        GoTo ED
                    Else

                        RST2.Fields("MOGODI") = RST2.Fields("MOGODI") + rst.Fields("MEGHk")
                        RST2.update
                    End If
               End If

               rst.Fields("ANBAR") = Me.ANBAR
               rst.update
               rst.MoveNext
           Wend
        End If
    End If
ED:
End Sub

Private Sub ANBAR_Enter()
                      On Error Resume Next
     Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY TCOD_ANBAR.CODE"
End Sub

Private Sub ANBARF_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
      Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, min As Double
  rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, INVO_LST.ANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.CUST_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "برگه داراي سطر كالا ميباشد نميتوانيد انبار را تغيير دهيد اول كالاهها را پاك كنيد"
        CANCEL = True
        GoTo ED
   End If
       Set rst = New ADODB.Recordset
    rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, INVO_LST.ANBAR, INVO_LST.RADIF, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MANDAH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.FROM_A, INVO_LST.N_RASID, INVO_LST.MEGH_R, INVO_LST.RADAH, INVO_LST.SANAD_NO, INVO_LST.CUST_NO, INVO_LST.ANBARF, INVO_LST.VAHED_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        RST2.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst3.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF()
            RST2.Filter = "code = '" & rst.Fields("code") & "' and anbar = " & Me.ANBARF
            Set RST2 = RST2
            If RST2.RecordCount = 0 Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كالا متعلق به انبار انتخاب شده نيست ابتدا در بخش تعريف كالا تعلق كالا به انبار را ايجاد كنيد.كد كالا :" & rst.Fields("CODE")
                CANCEL = True
                GoTo ED
            End If
            rst3.Filter = "CODE = '" & rst.Fields("CODE") & "'"
            Set rst3 = rst3
            If rst3.RecordCount = 0 Then
            Else
               min = rst3.Fields("MIN_M")
            End If
            RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & rst.Fields("ANBARF")
            Set RST2 = RST2
            If rst.RecordCount = 0 Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                CANCEL = True
                GoTo ED
            Else
               If (RST2.Fields("MOGODI") + RST2.Fields("MOGODI_A")) - rst.Fields("MEGHk") < min And Me.ANBAR > 0 Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "انتقال  كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "كد كالا : " & rst.Fields("code") & "حداقل موجودي تعريف شده در اف دو :" & min
                    CANCEL = True
                    GoTo ED
               End If
            End If
            rst.MoveNext
        Wend
        rst.MoveFirst
        While Not rst.EOF()
             rst3.Filter = "CODE = '" & rst.Fields("CODE") & "'"
             Set rst3 = rst3
             If rst3.RecordCount = 0 Then
             Else
                min = rst3.Fields("MIN_M")
             End If
             RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & rst.Fields("ANBARF")
             Set RST2 = RST2
             If RST2.RecordCount = 0 Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                CANCEL = True
                GoTo ED
             Else

                RST2.Fields("MOGODI") = RST2.Fields("MOGODI") - rst.Fields("MEGHk")
                RST2.update
'//////////////////////////////////////////////////////////////////////////////////////
                RST2.Filter = "CODE = '" & rst.Fields("CODE") & "' AND ANBAR = " & Me.ANBARF
                Set RST2 = RST2
                If RST2.RecordCount = 0 Then
                    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                    CANCEL = True
                    GoTo ED
                Else
                    RST2.Fields("MOGODI") = RST2.Fields("MOGODI") + rst.Fields("MEGHk")
                    RST2.update
                End If
           End If

           rst.Fields("ANBARF") = Me.ANBARF
           rst.update
           rst.MoveNext
       Wend
    End If
ED:

End Sub

Private Sub ANBARF_Enter()
                     On Error Resume Next
    Me.ANBARF.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) and (TCOD_ANBAR.CODE <> " & Me.ANBAR & ")  ORDER BY TCOD_ANBAR.CODE"
    Me.ANBARF.Requery
End Sub

Private Sub Command100_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, NOTPR As Boolean, min As Double
  If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
        NOTPR = False
        DoCmd.OpenForm "BUN"
        RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 5", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
              Forms![BUN].Form.Refresh
              '  RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              '  If RST3.RecordCount = 1 Then
              '    If IsNull(RST3.Fields("MIN_M")) Then
                    min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
              '    Else
              '       min = RST3.Fields("MIN_M")
              '    End If
              '  End If
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                     If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                      NOTPR = True
                   End If
                End If
                Set rst = New ADODB.Recordset
                Forms![BUN].Form.Refresh
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBARf") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                       If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                        NOTPR = True
                     End If
                End If
                RST2.MoveNext
                Set rst = New ADODB.Recordset
                Set rst3 = New ADODB.Recordset
        Wend
        DoCmd.Close acForm, "BUN"
  End If
  If NOTPR = False Then
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    DoCmd.Save acForm, Me.NAME
    DoCmd.OpenReport "BARGAH_ENTEGHAL", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_ENTEGHAL_SUB.Enabled = False
         Me.ESLAH.Enabled = True
    End If
 End If
End Sub

Private Sub custprint_Click()
                 On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, NOTPR As Boolean, min As Double
  If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
        NOTPR = False
        DoCmd.OpenForm "BUN"
        RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 5", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
              Forms![BUN].Form.Refresh
             '   RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             '   If RST3.RecordCount = 1 Then
             '     If IsNull(RST3.Fields("MIN_M")) Then
                    min = Getmin(Me.ANBAR, RST2.Fields("CODE"))
             '     Else
             '        min = RST3.Fields("MIN_M")
             '     End If
             '   End If
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                     If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                      NOTPR = True
                   End If
                End If
                Set rst = New ADODB.Recordset
                Forms![BUN].Form.Refresh
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBARf") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                       If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                        NOTPR = True
                     End If
                End If
                RST2.MoveNext
                Set rst = New ADODB.Recordset
                Set rst3 = New ADODB.Recordset
        Wend
        DoCmd.Close acForm, "BUN"
  End If
  If NOTPR = False Then
      Dim prt As Printer
      Dim frm As Form
      If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
           rptname = "BI" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
      Else
           rptname = "BI1"
      End If
'      DoCmd.RunCommand acCmdSaveRecord
      DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =5", , 2
      If Not Me.OKF Then
            Me.OKDATE = FARSIDATE(DATE)
            Me.OKTIME = GTFS()
      End If
      If Forms![BASEKNOW]![LOCKFAP] Then
              Me.OKF = True
      End If
      If Me.OKF Then
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_ENTEGHAL_SUB.Enabled = False
            Me.ESLAH.Enabled = True
      End If
   End If
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
                 On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
    dt = Now()
    Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & " ) and (TAG = 5)", dt, 1)
    Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & " ) and (TAG = 5)", dt, 2)
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me.INVO_LST_ENTEGHAL_SUB.Enabled = True
    If Me.SGN1 Or SGN2 Or SGN3 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
        Me.INVO_LST_ENTEGHAL_SUB.Locked = True
        Me.DATE_N.Locked = True
        Me.NUMBER.Locked = True
        Me.MOLAH.Locked = True
        Me.ANBAR.Locked = True
        Me.ANBARF.Locked = True
        Me.AllowEdits = True
    Else
        Me.INVO_LST_ENTEGHAL_SUB.Locked = False
        Me.NUMBER.Locked = False
        Me.MOLAH.Locked = False
        Me.DATE_N.Locked = False
        Me.ANBAR.Locked = False
        Me.ANBARF.Locked = False
        Me.AllowEdits = True
    End If
    Call SETSECURITY(Me.NAME, "ESWAP", 3)
    End If
End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
  '  If Me.USER_NAME <> UCurrentUser() Then
  '      Me.USER_NAME = UCurrentUser()
  '  End If
    If IsNull(Me.ANBAR) Or IsNull(Me.ANBARF) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "انبارها نمي تواند داراي مقدار خالي باشد"
            Me.INVO_LST_ENTEGHAL_SUB.Locked = True
            Exit Sub
    Else
            Me.INVO_LST_ENTEGHAL_SUB.Locked = False
    End If
'    rst.Open "SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE (((TCOD_ANBAR.CODE)=" & Me.ANBAR & ") AND ((TCOD_ANBAR.KIND)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'    If rst.RecordCount > 0 Then
'        rst.Close
'        Set rst = New ADODB.Recordset
'        rst.Open "SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE (((TCOD_ANBAR.CODE)=" & Me.ANBARF & ") AND ((TCOD_ANBAR.KIND)<>0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        If rst.RecordCount > 0 Then
'            DoCmd.OpenForm "mesageform", , , , , acDialog, "انبار مواد اوليه تنها به انبار مواد اوليه ديگر مي تواند كالا انتقال دهد"
'            Me.ANBAR = Null
'            Me.ANBARF = Null
'            Me.INVO_LST_ENTEGHAL_SUB.Locked = True
'            Exit Sub
'        Else
            Me.INVO_LST_ENTEGHAL_SUB.Locked = False
'        End If
'    End If
'    rst.Close
'    Set rst = New ADODB.Recordset
'    rst.Open "SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE (((TCOD_ANBAR.CODE)=" & Me.ANBAR & ") AND ((TCOD_ANBAR.KIND)<>0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'    If rst.RecordCount > 0 Then
'        rst.Close
'        Set rst = New ADODB.Recordset
'        rst.Open "SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE (((TCOD_ANBAR.CODE)=" & Me.ANBARF & ") AND ((TCOD_ANBAR.KIND)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        If rst.RecordCount > 0 Then
'            DoCmd.OpenForm "mesageform", , , , , acDialog, "انبار مواد اوليه تنها به انبار مواد اوليه ديگر مي تواند كالا انتقال دهد"
'            Me.ANBAR = Null
'            Me.ANBARF = Null
'            Me.INVO_LST_ENTEGHAL_SUB.Locked = True
'            Exit Sub
'        Else
            Me.INVO_LST_ENTEGHAL_SUB.Locked = False
'         End If
'    End If
  Call SANAD
  LETSANAD = False
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 6, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
If IsNull(Me.ANBAR) Or IsNull(Me.ANBARF) Then
   DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "هر دو انبار بايد مشخص شود...!"
   CANCEL = True
   Exit Sub
End If
  If Me.NUMBER = 0 Or IsNull(Me.NUMBER) Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STENT]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.INVO_LST_ENTEGHAL_SUB.Locked = False
  Me.Refresh
End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                     On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                     On Error Resume Next
    Me.PERSONEL.Visible = True
     Me.lsanad.ForeColor = 65535
  Set frm = Me![INVO_LST_ENTEGHAL_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
  Else
    Me.Command100.Enabled = False
  End If
   Me.ANBARF.Requery
   If Me.NewRecord Then
        Me.INVO_LST_ENTEGHAL_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.MOLAH.TAG = ""
    MABNA = Null
   Else
        Me.INVO_LST_ENTEGHAL_SUB.Locked = False
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
       MABNA = rst.Fields("BASE")
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_ENTEGHAL_SUB.Locked = True
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowAdditions = False
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowDeletions = False
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowEditing = False
            Me.lsanad.ForeColor = 125
        Else
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_ENTEGHAL_SUB.Locked = False
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowAdditions = True
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowDeletions = True
            Me![INVO_LST_ENTEGHAL_SUB].Form.AllowEditing = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
   End If
  If Forms![BASEKNOW]![SIGN] Then
       If Me.SGN1 And Me.SGN2 Or Me.SGN3 Then
            Me.Command100.Enabled = True
            Me.custprint.Enabled = True
       Else
            Me.Command100.Enabled = False
            Me.custprint.Enabled = False
       End If
  End If
    If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_ENTEGHAL_SUB.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 6, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
  Err.Clear
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_ENTEGHAL_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub


Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                     On Error Resume Next
   Response = acDataErrContinue
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                  On Error Resume Next
                  On Error Resume Next
 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
   Set SDRST = Forms![recording].Form.RecordsetClone
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 End If
 Call SETSECURITY(Me.NAME, "ESWAP", 3)
 DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
      If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
   If Not LETSGO("ESLAHE") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
    If Not LETSGO("SEEENT") Then
        Me.ServerFilter = "(USER_NAME = N'" & UCurrentUser() & "')"
        Me.Refresh
    End If
    If Not IsNull(Me.OpenArgs) Then
        Me.ServerFilter = Me.OpenArgs
        Me.Refresh
    Else
        DoCmd.GoToRecord acActiveDataObject, , acNewRec
    End If
If Forms![BASEKNOW]![SIGN] Then
    Me.SGN1.Visible = True
    Me.SGN2.Visible = True
    Me.SGN3.Visible = True
    Me.SGN1usid.Visible = True
    Me.sgn2usid.Visible = True
    Me.sgn3usid.Visible = True
    If IsLoaded("main") Then
        If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 6) Then
          Meidnum = Forms![main]![TASKS]![IDNUM]
          If GETSIGNER(6, Forms![BASEKNOW]![USERCOD]) > 0 Then
             ' Me.answerbtn.Visible = True
            End If
        End If
    End If
   ' Me.ersalbtn.Visible = True
End If
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
End Sub

Private Sub INVO_LST_ENTEGHAL_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
   Dim frm As Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
  Else
    Me.Command100.Enabled = False
  End If
   If Forms![BASEKNOW]![SIGN] Then
       If Me.SGN1 And Me.SGN2 Or Me.SGN3 Then
            Me.Command100.Enabled = True
            Me.custprint.Enabled = True
       Else
            Me.Command100.Enabled = False
            Me.custprint.Enabled = False
       End If
  End If
 If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub
Private Sub answerbtn_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "6-" & Me.NUMBER & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME
End Sub



Private Sub ersalbtn_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.NUMBER, 12) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين  سند انتقال  قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "' سند انتقال شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  از انبار: " & Me.ANBAR & "  به انبار: " & Me.ANBARF & "','" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(6, Me.NUMBER, Me.PERSONEL, "'انتقالي شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  از انبار: " & Me.ANBAR & "','توسط :" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'")
        DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub


Private Sub SGN1_Click()
    Dim MIDDU As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
  Dim rptname As String, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, NOTPR As Boolean, min As Double
  If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
        NOTPR = False
        DoCmd.OpenForm "BUN"
        RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 5", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
              Forms![BUN].Form.Refresh
              '  RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              '  If RST3.RecordCount = 1 Then
              '    If IsNull(RST3.Fields("MIN_M")) Then
                    min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
              '    Else
              '       min = RST3.Fields("MIN_M")
              '    End If
              '  End If
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                     If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                      NOTPR = True
                   End If
                End If
                Set rst = New ADODB.Recordset
                Forms![BUN].Form.Refresh
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBARf") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBARf") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                       If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                        NOTPR = True
                     End If
                End If
                RST2.MoveNext
                Set rst = New ADODB.Recordset
                Set rst3 = New ADODB.Recordset
        Wend
        DoCmd.Close acForm, "BUN"
  End If
  If NOTPR = False Or Not SGN1 Then
         MIDDU = Gettaskid(Me.NUMBER, 6)
         SHARH = "'انتقالي شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  از انبار: " & Me.ANBAR & "','توسط :" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
         If MIDDU > 0 Then
             CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
             CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MIDDU) & ",STATUS = 1 WHERE IDNUM = " & MIDDU)
         Else
             td = Now()
             CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
             MIDDU = Gettaskid(Me.NUMBER, 6)
             CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
        End If
        Call PERSONELUpdate(6, Me.NUMBER, Me.PERSONEL, SHARH)
        Me.PERSONEL.Visible = True
        Meidnum = MIDDU:
        Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
 Else
    Me.SGN1 = False
 End If

End Sub

Private Sub SGN2_Click()
    Dim MIDDU As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    MIDDU = Gettaskid(Me.NUMBER, 6)
    SHARH = "'انتقالي شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  از انبار: " & Me.ANBAR & "','توسط :" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
    If MIDDU > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MIDDU) & ",STATUS = 1 WHERE IDNUM = " & MIDDU)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
        MIDDU = Gettaskid(Me.NUMBER, 6)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
   End If
   Meidnum = MIDDU: If Not Me.OKF Then Me.OKF = True
   Call PERSONELUpdate(6, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
End Sub

Private Sub SGN3_Click()
    Dim MIDDU As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    MIDDU = Gettaskid(Me.NUMBER, 6)
    If MIDDU > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MIDDU) & ",STATUS = 1 WHERE IDNUM = " & MIDDU)
    Else
        td = Now()
        SHARH = "'انتقالي شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  از انبار: " & Me.ANBAR & "','توسط :" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
        MIDDU = Gettaskid(Me.NUMBER, 6)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDDU & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",6," & Me.NUMBER & ",6 )")
   End If
   Meidnum = MIDDU: If Not Me.OKF Then Me.OKF = True
   Call PERSONELUpdate(6, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
End Sub
