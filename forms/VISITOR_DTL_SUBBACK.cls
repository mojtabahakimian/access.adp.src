Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

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
    Dim rst As New ADODB.Recordset
    If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "52"
    Else
        Err.Clear
         Me.CUST_NO = NewData
    End If
       Response = acDataErrContinue
End Sub

Private Sub DARSAD_AfterUpdate()
  Me.PURSANT = Round((Forms![head_lst_froosh_back2]![JF] - Forms![head_lst_froosh_back2]![TAKHFIF]) * Me.DARSAD / 100)
  If Me.STAT Then
    Me.STAT = False
  End If
End Sub


Private Sub DARSAD_BeforeUpdate(CANCEL As Integer)
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

Private Sub Form_AfterUpdate()
   Forms![head_lst_froosh_back2]![LETSANAD] = True
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
   If Me.OpenArgs = 1 Then
        If IsNull(Me.NUMBER) Then
            Me.NUMBER = Forms![head_lst_froosh_back2]![NUMBER]
        End If
   End If
End Sub

Private Sub Form_Close()
                 On Error Resume Next
    If Me.OpenArgs = 1 And IsLoaded("HEAD_LST_FROOSH_back2") Then
        Forms![head_lst_froosh_back2]![VISITOR_DTL_SUB].Form.Requery
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    If IsLoaded("HEAD_LST_FROOSH_back2") Then
        If Not Forms![head_lst_froosh_back2].Form.AllowEdits Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
              DoCmd.Close acForm, Me.NAME
        Else
           If Me.OpenArgs = 1 Then
                Me.RecordSource = "select * from VISITOR_DTL where tag = 4 and NUMBER = " & Forms![head_lst_froosh_back2]![NUMBER]
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
        rst.Open "select code ,MEGH_mar * mabl + (megh_MAR/MEGH)*(imbaa - n_moin) as mablk from invo_lst where tag = 2 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        Me.PURSANT = prs
        If MBK > 0 And prs > 0 Then
            Me.DARSAD = Me.PURSANT / MBK * 100
        Else
            Me.DARSAD = 0
        End If
    Else
    End If
End Sub

Private Sub PURSANT_AfterUpdate()
                 On Error Resume Next
If (Forms![head_lst_froosh_back2]![JF] - Forms![head_lst_froosh_back2]![TAKHFIF]) <> 0 Then
    Me.DARSAD = Me.PURSANT / (Forms![head_lst_froosh_back2]![JF] - Forms![head_lst_froosh_back2]![TAKHFIF]) * 100
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

Private Sub PURSANT_KeyPress(KeyAscii As Integer)
sendkeyu (KeyAscii)

End Sub
