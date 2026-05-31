Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim KINDF As String
Private Sub CAMIUN_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    Me.OTHER_DTL_SUB_SUB.Requery
End Sub

Private Sub CAMIUN_NUM_AfterUpdate()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT   TOP 100 PERCENT * FROM OTHER_DTL WHERE CAMIUN_NUM = '" & Me.CAMIUN_NUM & "' ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("DRIVER_MOB") <> "" Then
            Me.DRIVER_MOB = rst.Fields("DRIVER_MOB")
        End If
        If rst.Fields("CAMIUN") <> "" Then
            Me.CAMIUN = rst.Fields("CAMIUN")
        End If
        If rst.Fields("driver") <> "" Or IsNull(rst.Fields("driver")) Then
            Me.driver = rst.Fields("driver")
        End If
        If rst.Fields("CAM_KHALY") > 0 Then
            Me.CAM_KHALY = rst.Fields("CAM_KHALY")
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
    Call INSERTVAZN
    Me.OTHER_DTL_SUB_SUB.Requery
    Me.CAMIUN.SetFocus
End Sub

Private Sub Command26_Click()
                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
     If Not (IsNull(Me.driver) And IsNull(CAMIUN_NUM)) Then
        Select Case KINDF
            Case "PISH":
                DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, RADIF) SELECT     NUMBER, TAG, CODE, RADIF  FROM         dbo.INVO_LST WHERE  TAG = 20 AND NUMBER = " & Me.NUMBER
            Case "HAV":
                DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, RADIF) SELECT     NUMBER, TAG, CODE, RADIF  FROM         dbo.INVO_LST WHERE  TAG = 2 AND NUMBER = " & Me.NUMBER
            Case "FROOSH22":
                DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, RADIF) SELECT     NUMBER, TAG, CODE, RADIF  FROM         dbo.INVO_LST WHERE  TAG = 2 AND NUMBER = " & Me.NUMBER
            Case "FROOSH2":
                DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, RADIF) SELECT     NUMBER, TAG, CODE, RADIF  FROM         dbo.INVO_LST WHERE  TAG = 2 AND NUMBER = " & Me.NUMBER
            Case "RASID":
                DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, RADIF) SELECT     NUMBER, TAG, CODE, RADIF  FROM         dbo.INVO_LST WHERE  TAG = 1 AND NUMBER = " & Me.NUMBER
        End Select
        Me.OTHER_DTL_SUB_SUB.Requery
    Else
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ابتدا مشخصات راننده و كاميون را وارد كنيد "
    End If
End Sub

Private Sub Command27_Click()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
     If Not (IsNull(Me.driver) And IsNull(CAMIUN_NUM)) Then
        Select Case KINDF
        Case "PISH":
            rst.Open "SELECT   * FROM OTHER_DTL_sub WHERE tagg= 20 and MEGHk > 0 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 While Not rst.EOF
                    RST2.Open "select * from invo_lst where tag = 20 and NUMBER = " & Me.NUMBER & " and code = '" & rst.Fields("code") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount = 1 Then
                         RST2.Fields("megh") = rst.Fields("meghk")
                         RST2.Fields("meghk") = rst.Fields("meghk")
                         RST2.Fields("mabl_k") = Round(rst.Fields("meghk") * RST2.Fields("mabl"))
                         RST2.Fields("n_moin") = Round(rst.Fields("meghk") * RST2.Fields("mabl") * RST2.Fields("n_kol") / 100)
                         RST2.update
                    End If
                    rst.MoveNext
                    Set RST2 = New ADODB.Recordset
                 Wend
                 If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
                      Set rst = New ADODB.Recordset
                      rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      While Not rst.EOF
                          RST2.Open "select CMBAA ,code FROM STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                          If RST2.RecordCount > 0 Then
                              If RST2.Fields("CMBAA") Then
                                   rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                                   SMBAA = SMBAA + Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                              Else
                                  rst.Fields("IMBAA") = 0
                              End If
                          End If
                          RST2.Close
                          Set RST2 = New ADODB.Recordset
                          rst.update
                          rst.MoveNext
                      Wend
                      If SMBAA <> Forms![HEAD_LST_PISHFROOSH2]![MBAA] And SMBAA > 0 Then
                        Forms![HEAD_LST_PISHFROOSH2]![MBAA] = SMBAA
                      End If
                  End If
             End If
             Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Requery
        Case "FROOSH22":
            rst.Open "SELECT   * FROM OTHER_DTL_sub WHERE tagg= 2 and MEGHk > 0 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 While Not rst.EOF
                    RST2.Open "select * from invo_lst where tag = 2 and NUMBER = " & Me.NUMBER & " and code = '" & rst.Fields("code") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount = 1 Then
                         RST2.Fields("megh") = rst.Fields("meghk")
                         RST2.Fields("meghk") = rst.Fields("meghk")
                         RST2.Fields("mabl_k") = Round(rst.Fields("meghk") * RST2.Fields("mabl"))
                         RST2.Fields("n_moin") = Round(rst.Fields("meghk") * RST2.Fields("mabl") * RST2.Fields("n_kol") / 100)
                         RST2.update
                    End If
                    rst.MoveNext
                    Set RST2 = New ADODB.Recordset
                 Wend
                 If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
                      Set rst = New ADODB.Recordset
                      rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      While Not rst.EOF
                          RST2.Open "select CMBAA ,code FROM STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                          If RST2.RecordCount > 0 Then
                              If RST2.Fields("CMBAA") Then
                                   rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                                   SMBAA = SMBAA + Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                              Else
                                  rst.Fields("IMBAA") = 0
                              End If
                          End If
                          RST2.Close
                          Set RST2 = New ADODB.Recordset
                          rst.update
                          rst.MoveNext
                      Wend
                      If SMBAA <> Forms![HEAD_LST_FROOSH22]![MBAA] And SMBAA > 0 Then
                        Forms![HEAD_LST_FROOSH22]![MBAA] = SMBAA
                        Forms![HEAD_LST_FROOSH22]![HMBAA] = Forms![BASEKNOW]![HESMBAA]
                      End If
                  End If
             End If
             Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Requery
        Case "RASID":
             rst.Open "SELECT   * FROM OTHER_DTL_sub WHERE tagg= 1 and MEGHk > 0 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 While Not rst.EOF
                    RST2.Open "select * from invo_lst where tag = 1 and NUMBER = " & Me.NUMBER & " and code = '" & rst.Fields("code") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount = 1 Then
                         RST2.Fields("megh") = rst.Fields("VAZNH")
                         RST2.Fields("meghk") = rst.Fields("VAZNH")
                         RST2.Fields("megh_R") = rst.Fields("VAZNH")
                         RST2.Fields("mabl_k") = Round(rst.Fields("VAZNH") * RST2.Fields("mabl"))
                         RST2.update
                    End If
                    rst.MoveNext
                    Set RST2 = New ADODB.Recordset
                 Wend
              End If
              Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Requery
        End Select
        DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub DRIVER_AfterUpdate()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT   TOP 100 PERCENT * FROM OTHER_DTL WHERE DRIVER = '" & Me.driver & "' ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("DRIVER_MOB") <> "" Then
            Me.DRIVER_MOB = rst.Fields("DRIVER_MOB")
        End If
        If rst.Fields("CAMIUN_NUM") <> "" Then
            Me.CAMIUN_NUM = rst.Fields("CAMIUN_NUM")
        End If
        If rst.Fields("CAMIUN") <> "" Or IsNull(rst.Fields("CAMIUN")) Then
            Me.CAMIUN = rst.Fields("CAMIUN")
        End If
        If rst.Fields("CAM_KHALY") > 0 Then
            Me.CAM_KHALY = rst.Fields("CAM_KHALY")
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
    Me.OTHER_DTL_SUB_SUB.Requery
    Me.DRIVER_MOB.SetFocus
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                     On Error Resume Next
   If Me.OpenArgs = 1 Then
        If IsNull(Me.NUMBER) Then
          If IsLoaded("HEAD_LST_PISHFROOSH2") Then
             Me.NUMBER = Forms![HEAD_LST_PISHFROOSH2]![NUMBER]
             Me!TAG = 20
          Else
            If IsLoaded("HEAD_LST_HAVL") Then
               Me.NUMBER = Forms![head_lst_havl]![NUMBER]
               Me!TAG = 2
            Else
                If IsLoaded("HEAD_LST_FROOSH22") Then
                    Me.NUMBER = Forms![HEAD_LST_FROOSH22]![NUMBER]
                   Me!TAG = 2
                Else
                    If IsLoaded("HEAD_LST_FROOSH2") Then
                        Me.NUMBER = Forms![HEAD_LST_FROOSH2]![NUMBER]
                       Me!TAG = 2
                    Else
                        Me.NUMBER = Forms![head_lst_rasid]![NUMBER]
                        Me!TAG = 1
                    End If
             End If
            End If
          End If
        End If
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
  Me.OTHER_DTL_SUB_SUB.Form![VAZNH].ColumnHidden = True
  If IsLoaded("HEAD_LST_PISHFROOSH2") Then
    DoCmd.Close acForm, "HEAD_LST_FROOSH22"
    If Me.OpenArgs = 1 Then
        Me.RecordSource = "select * from OTHER_DTL where tag = 20 and NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER]
         rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") And (dbo.INVO_LST.TAG = 20) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         RST2.Open "select * from OTHER_DTL_sub where tagg = 20 and NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         While Not rst.EOF
            RST2.AddNew
            RST2.Fields("tagg") = 20
            RST2.Fields("code") = rst.Fields("code")
            RST2.Fields("radif") = rst.Fields("radif")
            RST2.Fields("NUMBER") = Forms![HEAD_LST_PISHFROOSH2]![NUMBER]
            RST2.update
            rst.MoveNext
         Wend
         DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 20) AND (NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") AND (TAG = 20))))"
         Me.OTHER_DTL_SUB_SUB.Requery
    End If
    KINDF = "PISH"
  Else
    If IsLoaded("HEAD_LST_HAVL") Then
      DoCmd.Close acForm, "HEAD_LST_FROOSH22"
      If Me.OpenArgs = 1 Then
           Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![head_lst_havl]![NUMBER]
      End If
      KINDF = "HAV"
    Else
      If IsLoaded("HEAD_LST_FROOSH22") Then
          KINDF = "FROOSH22"
          If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
                DoCmd.Close acForm, Me.NAME
          Else
              If Me.OpenArgs = 1 Then
                   Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER]
                   rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") And (dbo.INVO_LST.TAG = 2) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   RST2.Open "select * from OTHER_DTL_sub where tagg = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rst.EOF
                      RST2.AddNew
                      RST2.Fields("tagg") = 2
                      RST2.Fields("code") = rst.Fields("code")
                      RST2.Fields("radif") = rst.Fields("radif")
                      RST2.Fields("NUMBER") = Forms![HEAD_LST_FROOSH22]![NUMBER]
                      RST2.update
                      rst.MoveNext
                   Wend
                   DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 2) AND (NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") AND (TAG = 2))))"
                   Me.OTHER_DTL_SUB_SUB.Requery
              End If
          End If
      Else
          If IsLoaded("HEAD_LST_RASID") Then
              KINDF = "RASID"
              Me.OTHER_DTL_SUB_SUB.Form![VAZNH].ColumnHidden = False
              If Not Forms![head_lst_rasid].Form.AllowEdits Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا رسيد قابل اصلاح باشد"
                    DoCmd.Close acForm, Me.NAME
              Else
                  If Me.OpenArgs = 1 Then
                       Me.RecordSource = "select * from OTHER_DTL where tag = 1 and NUMBER = " & Forms![head_lst_rasid]![NUMBER]
                       rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") And (dbo.INVO_LST.TAG = 1) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       RST2.Open "select * from OTHER_DTL_sub where tagg = 1 and NUMBER = " & Forms![head_lst_rasid]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rst.EOF
                          RST2.AddNew
                          RST2.Fields("tagg") = 1
                          RST2.Fields("code") = rst.Fields("code")
                          RST2.Fields("radif") = rst.Fields("radif")
                          RST2.Fields("NUMBER") = Forms![head_lst_rasid]![NUMBER]
                          RST2.update
                          rst.MoveNext
                       Wend
                       DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 1) AND (NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") AND (TAG = 1))))"
                       Me.OTHER_DTL_SUB_SUB.Requery
                  End If
              End If
          Else
                If IsLoaded("HEAD_LST_FROOSH2") Then
                   KINDF = "FROOSH2"
                   If Not Forms![HEAD_LST_FROOSH2].Form.AllowEdits Then
                         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
                         DoCmd.Close acForm, Me.NAME
                   Else
                       If Me.OpenArgs = 1 Then
                            Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER]
                            rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER] & ") And (dbo.INVO_LST.TAG = 2) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            RST2.Open "select * from OTHER_DTL_sub where tagg = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            While Not rst.EOF
                               RST2.AddNew
                               RST2.Fields("tagg") = 2
                               RST2.Fields("code") = rst.Fields("code")
                               RST2.Fields("radif") = rst.Fields("radif")
                               RST2.Fields("NUMBER") = Forms![HEAD_LST_FROOSH2]![NUMBER]
                               RST2.update
                               rst.MoveNext
                            Wend
                            DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 2) AND (NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER] & ") AND (TAG = 2))))"
                            Me.OTHER_DTL_SUB_SUB.Requery
                       End If
                   End If
          Else
             DoCmd.Close acForm, Me.NAME
          End If
      End If
      End If
    End If
  End If
End Sub

Private Sub MAGHSAD_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "TCODE_CODING", acFormDS, , , , acDialog, 1
End Sub
Private Sub INSERTVAZN()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
  Me.OTHER_DTL_SUB_SUB.Form![VAZNH].ColumnHidden = True
  If IsLoaded("HEAD_LST_PISHFROOSH2") Then
    DoCmd.Close acForm, "HEAD_LST_FROOSH22"
    If Me.OpenArgs = 1 Then
        Me.RecordSource = "select * from OTHER_DTL where tag = 20 and NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER]
         rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") And (dbo.INVO_LST.TAG = 20) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         RST2.Open "select * from OTHER_DTL_sub where tagg = 20 and NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         While Not rst.EOF
            RST2.AddNew
            RST2.Fields("tagg") = 20
            RST2.Fields("code") = rst.Fields("code")
            RST2.Fields("radif") = rst.Fields("radif")
            RST2.Fields("NUMBER") = Forms![HEAD_LST_PISHFROOSH2]![NUMBER]
            RST2.update
            rst.MoveNext
         Wend
         DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 20) AND (NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![HEAD_LST_PISHFROOSH2]![NUMBER] & ") AND (TAG = 20))))"
         Me.OTHER_DTL_SUB_SUB.Requery
    End If
    KINDF = "PISH"
  Else
    If IsLoaded("HEAD_LST_HAVL") Then
      DoCmd.Close acForm, "HEAD_LST_FROOSH22"
      If Me.OpenArgs = 1 Then
           Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![head_lst_havl]![NUMBER]
      End If
      KINDF = "HAV"
    Else
      If IsLoaded("HEAD_LST_FROOSH22") Then
          KINDF = "FROOSH22"
          If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
                DoCmd.Close acForm, Me.NAME
          Else
              If Me.OpenArgs = 1 Then
                   Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER]
                   rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") And (dbo.INVO_LST.TAG = 2) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   RST2.Open "select * from OTHER_DTL_sub where tagg = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rst.EOF
                      RST2.AddNew
                      RST2.Fields("tagg") = 2
                      RST2.Fields("code") = rst.Fields("code")
                      RST2.Fields("radif") = rst.Fields("radif")
                      RST2.Fields("NUMBER") = Forms![HEAD_LST_FROOSH22]![NUMBER]
                      RST2.update
                      rst.MoveNext
                   Wend
                   DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 2) AND (NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") AND (TAG = 2))))"
                   Me.OTHER_DTL_SUB_SUB.Requery
              End If
          End If
      Else
          If IsLoaded("HEAD_LST_RASID") Then
              KINDF = "RASID"
              Me.OTHER_DTL_SUB_SUB.Form![VAZNH].ColumnHidden = False
              If Not Forms![head_lst_rasid].Form.AllowEdits Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا رسيد قابل اصلاح باشد"
                    DoCmd.Close acForm, Me.NAME
              Else
                  If Me.OpenArgs = 1 Then
                       Me.RecordSource = "select * from OTHER_DTL where tag = 1 and NUMBER = " & Forms![head_lst_rasid]![NUMBER]
                       rst.Open "SELECT  dbo.INVO_LST.CODE ,dbo.INVO_LST.radif FROM  dbo.INVO_LST LEFT OUTER JOIN   dbo.OTHER_DTL_SUB ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND   dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE WHERE (dbo.INVO_LST.NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") And (dbo.INVO_LST.TAG = 1) And (dbo.OTHER_DTL_SUB.NUMBER Is Null)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       RST2.Open "select * from OTHER_DTL_sub where tagg = 1 and NUMBER = " & Forms![head_lst_rasid]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rst.EOF
                          RST2.AddNew
                          RST2.Fields("tagg") = 1
                          RST2.Fields("code") = rst.Fields("code")
                          RST2.Fields("radif") = rst.Fields("radif")
                          RST2.Fields("NUMBER") = Forms![head_lst_rasid]![NUMBER]
                          RST2.update
                          rst.MoveNext
                       Wend
                       DoCmd.RunSQL "DELETE FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 1) AND (NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") AND (NOT (CODE IN   (SELECT     CODE  FROM dbo.INVO_LST   WHERE     (NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") AND (TAG = 1))))"
                       Me.OTHER_DTL_SUB_SUB.Requery
                  End If
              End If
          Else
              DoCmd.Close acForm, Me.NAME
          End If
      End If
    End If
  End If
End Sub
