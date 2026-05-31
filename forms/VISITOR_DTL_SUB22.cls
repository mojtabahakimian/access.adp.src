Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CUST_NO_AfterUpdate()
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
       rst.Open "SELECT  TOZIH frOM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            If IsNumeric(rst.Fields("tozih")) Then
                Me.DARSAD = rst.Fields("tozih")
                Me.PURSANT = Round((Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * Me.DARSAD / 100)
                If Me.STAT Then
                  Me.STAT = False
                End If
            End If
       End If
End Sub

Private Sub CUST_NO_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
    If Me.CUST_NO = "" Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "نام ويزيتور نميتواند خالي باشد!"
       CANCEL = True
    End If
End Sub

Private Sub CUST_NO_DblClick(CANCEL As Integer)
                 On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset
If BLOCKED(GETKOL(Me.CUST_NO), GETMOIN(Me.CUST_NO), GETTAF(Me.CUST_NO)) Then
   DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
   Exit Sub
End If
        If IsLoaded("R_DAFTAR_MOIN_LIST") Then
              DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
        End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.CUST_NO & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With

End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
    If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "41"
    Else
        Err.Clear
        rst.Open "SELECT    hes FROM dbo.CUST_HESAB WHERE     (hes = N'" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Err = 0 Then
            If rst.RecordCount = 1 Then
                Me.CUST_NO = NewData
            End If
        End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT  TOZIH frOM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         If IsNumeric(rst.Fields("tozih")) Then
             Me.DARSAD = rst.Fields("tozih")
             Me.PURSANT = Round((Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * Me.DARSAD / 100)
             If Me.STAT Then
               Me.STAT = False
             End If
         End If
    End If
    Response = acDataErrContinue
End Sub

Private Sub DARSAD_AfterUpdate()
                  On Error Resume Next
 Me.PURSANT = Round((Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * Me.DARSAD / 100)
  If Me.STAT Then
    Me.STAT = False
  End If
End Sub


Private Sub DARSAD_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
  If Not IsNull(Me.PORID) Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "باتوجه به اينكه اين سطر دراي الگوي پرداخت پورسانت ميباشد با زدن درصد الگوي آن حذف ميشود آيا از ادامه عمليات اطمينان داريد."
    If Forms![BASEKNOW]![Text44] Then
        Me.PORID = Null
    Else
        CANCEL = True
    End If
  End If
End Sub

Private Sub DARSAD_DblClick(CANCEL As Integer)
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Me.NewRecord Then
       rst.Open "SELECT     TOP 100 PERCENT dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT FROM         dbo.HEAD_LST INNER JOIN   dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG WHERE     (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.NUMBER <> " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ")GROUP BY dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT,  dbo.HEAD_LST.NUMBER ORDER BY dbo.HEAD_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            Me.CUST_NO = rst.Fields("CUST_NO")
            Me.DARSAD = rst.Fields("DARSAD")
            Me.PURSANT = Round((Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * rst.Fields("DARSAD") / 100)
            Me.STAT = False
            Me.TOZIH = rst.Fields("TOZIH")
       End If
    End If
End Sub

Private Sub Form_AfterUpdate()
   Forms![HEAD_LST_FROOSH22]![LETSANAD] = True
   Dim rst As New ADODB.Recordset
   rst.Open "select NUMBER1  from head_lst where NUMBER1 = " & Me.NUMBER & " and tag = 4", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "توجه  ! توجه  : براي اين فاكتور مرجوعي ثبت شده است اگر ويزيتور آنرا تغيير ميدهيد لازم است در فاكتور برگشت فروش هم ويزيتور آنرا اصلاح كنيد "
   End If
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                 On Error Resume Next
   If Me.OpenArgs = 1 Then
        If IsNull(Me.NUMBER) Then
            Me.NUMBER = Forms![HEAD_LST_FROOSH22]![NUMBER]
        End If
   End If
End Sub

Private Sub Form_Close()
                 On Error Resume Next
    If Me.OpenArgs = 1 And IsLoaded("HEAD_LST_FROOSH22") Then
        Forms![HEAD_LST_FROOSH22]![VISITOR_DTL_SUB].Form.Requery
    End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
   Forms![HEAD_LST_FROOSH22]![LETSANAD] = True
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    If IsLoaded("HEAD_LST_FROOSH22") Then
        If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
              DoCmd.Close acForm, Me.NAME
        Else
           If Me.OpenArgs = 1 Then
                Me.RecordSource = "select * from VISITOR_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER]
           End If
        End If
  Else
    DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub PORID_AfterUpdate()
                 On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, prs As Long, MBK As Long
    prs = 0
    If Not IsNull(PORID) Then
        rst.Open "select code ,mabl_k - n_moin as mablk from invo_lst where tag = 2 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            RST2.Open "SELECT     PORSANT FROM dbo.VISITORS_PORSANT_KALA WHERE     (PORID = " & Me.PORID & ") and (code = '" & rst.Fields("code") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 1 Then
                prs = prs + Round(rst.Fields("mablk") * RST2.Fields("PORSANT") / 100)
                MBK = MBK + rst.Fields("mablk")
            Else
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "تذكر مهم :اين كالا فاقد الگو براي اين ويزيتور است و پورسانت محاسبه نشد.درصورت لزوم براي آن تعريف كنيد و همينجا مجددا الگو را انتخاب كنيد  : " & GETKALANAME(rst.Fields("code"))
            End If
            rst.MoveNext
            RST2.Close
        Wend
        Me.PURSANT = Round(prs)
        If MBK > 0 And prs > 0 Then
            Me.DARSAD = Me.PURSANT / MBK * 100
        Else
            Me.DARSAD = 0
        End If
    Else
    End If
End Sub

Private Sub PORID_DblClick(CANCEL As Integer)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Me.NewRecord Then
       rst.Open "SELECT     TOP 100 PERCENT dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT, dbo.VISITOR_DTL.PORID FROM         dbo.HEAD_LST INNER JOIN   dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG WHERE     (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.NUMBER <> " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ")GROUP BY dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT,  dbo.HEAD_LST.NUMBER, dbo.VISITOR_DTL.PORID ORDER BY dbo.HEAD_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 And Not IsNull(rst.Fields("porid")) Then
            Me.CUST_NO = rst.Fields("CUST_NO")
            Me.STAT = True
            Me.TOZIH = rst.Fields("TOZIH")
            Me.PORID = rst.Fields("PORID")
            Call PORID_AfterUpdate
       End If
    End If

End Sub

Private Sub PURSANT_AfterUpdate()
                 On Error Resume Next
If (Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF] + Forms![HEAD_LST_FROOSH22]![MBAA]) <> 0 Then
    Me.DARSAD = Me.PURSANT / (Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * 100
  Else
    Me.DARSAD = 0
  End If
  If Not Me.STAT Then
    Me.STAT = True
  End If

End Sub

Private Sub PURSANT_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
  If Not IsNull(Me.PORID) Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "باتوجه به اينكه اين سطر دراي الگوي پرداخت پورسانت ميباشد با زدن مبلغ الگوي آن حذف ميشود آيا از ادامه عمليات اطمينان داريد."
    If Forms![BASEKNOW]![Text44] Then
        Me.PORID = Null
    Else
        CANCEL = True
    End If
  End If

End Sub

Private Sub PURSANT_DblClick(CANCEL As Integer)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Me.NewRecord Then
       rst.Open "SELECT     TOP 100 PERCENT dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT FROM         dbo.HEAD_LST INNER JOIN   dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG WHERE     (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.NUMBER <> " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ")GROUP BY dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT,  dbo.HEAD_LST.NUMBER ORDER BY dbo.HEAD_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            Me.CUST_NO = rst.Fields("CUST_NO")
            If (Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) <> 0 Then
                Me.DARSAD = rst.Fields("PURSANT") / (Forms![HEAD_LST_FROOSH22]![JF] - Forms![HEAD_LST_FROOSH22]![TAKHFIF]) * 100
            Else
                Me.DARSAD = 0
            End If
            Me.PURSANT = rst.Fields("PURSANT")
            Me.STAT = True
            Me.TOZIH = rst.Fields("TOZIH")
       End If
    End If
End Sub

Private Sub PURSANT_KeyPress(KeyAscii As Integer)
sendkeyu (KeyAscii)

End Sub
