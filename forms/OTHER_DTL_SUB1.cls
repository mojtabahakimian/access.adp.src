Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

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
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                     On Error Resume Next
   If Me.OpenArgs = 1 Then
        If IsNull(Me.NUMBER) Then
          If IsLoaded("HEAD_LST_HAVL") Then
             Me.NUMBER = Forms![head_lst_havl]![NUMBER]
          Else
            If IsLoaded("HEAD_LST_FROOSH2") Then
                Me.NUMBER = Forms![HEAD_LST_FROOSH2]![NUMBER]
            Else
                Me.NUMBER = Forms![HEAD_LST_FROOSH22]![NUMBER]
           End If
          End If
        End If
   End If
End Sub

Private Sub Form_Close()
                     On Error Resume Next
    If Me.OpenArgs = 1 And IsLoaded("HEAD_LST_FROOSH2") Then
         Forms![HEAD_LST_FROOSH2]![OTHER_DTL_SUB].Form.Requery
    Else
        If Me.OpenArgs = 1 And IsLoaded("HEAD_LST_FROOSH22") Then
           Forms![HEAD_LST_FROOSH22]![OTHER_DTL_SUB].Form.Requery
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
  If IsLoaded("HEAD_LST_HAVL") Then
    DoCmd.Close acForm, "HEAD_LST_FROOSH2"
    DoCmd.Close acForm, "HEAD_LST_FROOSH22"
    If Me.OpenArgs = 1 Then
         Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![head_lst_havl]![NUMBER]
    End If
  Else
    If IsLoaded("HEAD_LST_FROOSH2") Then
      If Not Forms![HEAD_LST_FROOSH2].Form.AllowEdits Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
            DoCmd.Close acForm, Me.NAME
      Else
          If Me.OpenArgs = 1 Then
               Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER]
          End If
       End If
    Else
        If IsLoaded("HEAD_LST_FROOSH22") Then
          If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
                DoCmd.Close acForm, Me.NAME
          Else
              If Me.OpenArgs = 1 Then
                   Me.RecordSource = "select * from OTHER_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH22]![NUMBER]
              End If
           End If
        Else
       DoCmd.Close acForm, Me.NAME
    End If
    End If
  End If
End Sub

Private Sub MAGHSAD_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "TCODE_CODING", acFormDS, , , , acDialog, 1
End Sub
