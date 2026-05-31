Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Combo34_AfterUpdate()
  If Not IsNull(Combo34) Then
        Me.Combo36 = Me.Combo34
        Me.Combo36.Requery
   End If
End Sub

Private Sub Combo34_NotInList(NewData As String, Response As Integer)
 Dim rst As New ADODB.Recordset
    If Combo34.Text = "-" Or Combo34.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "39"
       Response = acDataErrContinue
    Else
      If Len(NewData) > 7 Then
          Err.Clear
          rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount > 0 Then
              Me.Combo36 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
              Me.Combo34 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
               Response = acDataErrContinue
            End If
          End If
      Else
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                Me.Combo34 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                Response = acDataErrContinue
            End If
          End If
      End If
    End If

End Sub

Private Sub Combo39_AfterUpdate()
  If Not IsNull(Combo39) Then
        Me.Combo41 = Me.Combo39
        Me.Combo41.Requery
   End If
End Sub

Private Sub Combo39_NotInList(NewData As String, Response As Integer)
 Dim rst As New ADODB.Recordset
    If Combo39.Text = "-" Or Combo39.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "40"
       Response = acDataErrContinue
    Else
      If Len(NewData) > 7 Then
          Err.Clear
          rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount > 0 Then
              Me.Combo39 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
              Me.Combo41 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
               Response = acDataErrContinue
            End If
          End If
      Else
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                Me.Combo39 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                Response = acDataErrContinue
            End If
          End If
      End If
    End If
End Sub

Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String, MAND As Double, SBES As Double, HESM As String, CHK As Boolean
  Dim rst  As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, PATH As String, i As Integer
  If (IsNull(Combo34) Or IsNull(Combo39)) And IsNull(Me.visit) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  MAND = 0
  HESM = -1
  CHK = False
  If Me.BEDBES = 1 Then
        DoCmd.RunSQL "DELETE FROM dbo.AI_DAFTAR"
        RST2.Open "AI_DAFTAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(Me.visit) Then
            rst.Open "SELECT AI.* FROM dbo.AI( '" & Me.Combo34 & "','" & Me.Combo39 & "') AI", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Else
            rst.Open "SELECT AI2.* FROM dbo.AI2( N'" & Me.visit & "') AI2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        While Not rst.EOF
              If rst.Fields("HES") <> HESM Then
                MAND = 0
                rst3.Open "SELECT     SUM(BES) AS sbes FROM dbo.DEED_DTL WHERE     (HES = N'" & rst.Fields("HES") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                SBES = rst3.Fields("SBES")
                Set rst3 = New ADODB.Recordset
                HESM = rst.Fields("HES")
              End If
              SBES = SBES - rst.Fields("BED")
              If SBES <= 0 Then
                    RST2.AddNew
                    RST2.Fields("N_S") = rst.Fields("N_S")
                    RST2.Fields("HES_T") = rst.Fields("HES_T")
                    RST2.Fields("HES_M") = rst.Fields("HES_M")
                    RST2.Fields("DATE_S") = rst.Fields("DATE_S")
                    RST2.Fields("HES_K") = rst.Fields("HES_K")
                    RST2.Fields("SHARH") = RIGHT("مانده تا اين تاريخ - " & rst.Fields("SHARH"), 255)
                    RST2.Fields("BED") = Abs(SBES)
                    RST2.Fields("BES") = rst.Fields("BES")
                    RST2.Fields("HES") = rst.Fields("HES")
                    MAND = Abs(SBES) + MAND
                    RST2.Fields("MAND") = Round(Abs(MAND))
                    RST2.Fields("TS") = IIf(MAND >= 0, 1, -1)
                    RST2.Fields("DIFF") = DIFF(rst.Fields("DATE_S"), FARSIDATE(DATE))
                    RST2.update
                    rst.MoveNext
                    While Not rst.EOF And rst.Fields("HES") = HESM
                        If rst.Fields("BES") = 0 Then
                            RST2.AddNew
                            RST2.Fields("N_S") = rst.Fields("N_S")
                            RST2.Fields("HES_T") = rst.Fields("HES_T")
                            RST2.Fields("HES_M") = rst.Fields("HES_M")
                            RST2.Fields("DATE_S") = rst.Fields("DATE_S")
                            RST2.Fields("HES_K") = rst.Fields("HES_K")
                            RST2.Fields("SHARH") = rst.Fields("SHARH")
                            RST2.Fields("BED") = rst.Fields("BEd")
                            RST2.Fields("BES") = rst.Fields("BES")
                            RST2.Fields("HES") = rst.Fields("HES")
                            MAND = rst.Fields("BEd") + MAND
                            RST2.Fields("MAND") = Round(Abs(MAND))
                            RST2.Fields("TS") = IIf(MAND >= 0, 1, -1)
                            RST2.Fields("DIFF") = DIFF(rst.Fields("DATE_S"), FARSIDATE(DATE))
                            RST2.update
                        End If
                        rst.MoveNext
                        If rst.EOF Then
                          GoTo b1
                        End If
                   Wend
b1:                rst.MovePrevious
             End If
             rst.MoveNext
        Wend
        DoCmd.RunSQL "DELETE FROM dbo.AI_DAFTAR WHERE     (BED = 0) AND (BES = 0)"
        DoCmd.OpenReport "R_DAFTAR_AI", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
  Else
        RST2.Open "AI_DAFTARMOIN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        DoCmd.RunSQL "DELETE FROM dbo.AI_DAFTARMOIN"
        If IsNull(Me.visit) Then
        RST2.Open "AI_DAFTAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Else
            rst.Open "SELECT AI2.* FROM dbo.AI2( ,N'" & Me.visit & "') AI2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        While Not rst.EOF
              If rst.Fields("HES") <> HESM Then
                MAND = 0
                HESM = rst.Fields("HES")
              End If
              RST2.AddNew
              RST2.Fields("N_S") = rst.Fields("N_S")
              RST2.Fields("HES_T") = rst.Fields("HES_T")
              RST2.Fields("HES_M") = rst.Fields("HES_M")
              RST2.Fields("DATE_S") = rst.Fields("DATE_S")
              RST2.Fields("HES_K") = rst.Fields("HES_K")
              RST2.Fields("SHARH") = rst.Fields("SHARH")
              RST2.Fields("BED") = rst.Fields("BED")
              RST2.Fields("BES") = rst.Fields("BES")
              RST2.Fields("HES") = rst.Fields("HES")
              MAND = rst.Fields("MAND") + MAND
              RST2.Fields("MAND") = Round(Abs(MAND))
              RST2.Fields("TS") = IIf(MAND >= 0, 1, -1)
              RST2.update
              rst.MoveNext
        Wend
        rst.Close
        RST2.Close
        Set rst = New ADODB.Recordset
        Set RST2 = New ADODB.Recordset
        RST2.Open "AI_DAFTAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Open "AI_DAFTARMOIN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        DoCmd.RunSQL "DELETE FROM dbo.AI_DAFTAR"
        rst.MoveLast
        MAND = 0
        HESM = -1
        While Not rst.BOF
              If rst.Fields("HES") <> HESM Then
                MAND = 0
                HESM = rst.Fields("HES")
                CHK = False
              End If
              If (Round(rst.Fields("MAND") * rst.Fields("TS") * -1) >= Round(rst.Fields("BES"))) And (rst.Fields("MAND") <> 0) Then
                  RST2.AddNew
                  RST2.Fields("N_S") = rst.Fields("N_S")
                  RST2.Fields("HES_T") = rst.Fields("HES_T")
                  RST2.Fields("HES_M") = rst.Fields("HES_M")
                  RST2.Fields("DIFF") = DIFF(rst.Fields("DATE_S"), FARSIDATE(DATE))
                  RST2.Fields("DATE_S") = rst.Fields("DATE_S")
                  RST2.Fields("HES_K") = rst.Fields("HES_K")
                  RST2.Fields("SHARH") = rst.Fields("SHARH")
                  RST2.Fields("BED") = rst.Fields("BED")
                  RST2.Fields("BES") = rst.Fields("BES")
                  RST2.Fields("HES") = rst.Fields("HES")
                  RST2.Fields("MAND") = Round(rst.Fields("MAND"))
                  RST2.Fields("TS") = rst.Fields("TS")
                  RST2.update
                  CHK = True
              Else
                 If CHK Then
                      RST2.AddNew
                      RST2.Fields("N_S") = rst.Fields("N_S")
                      RST2.Fields("HES_T") = rst.Fields("HES_T")
                      RST2.Fields("DIFF") = DIFF(rst.Fields("DATE_S"), FARSIDATE(DATE))
                      RST2.Fields("HES") = rst.Fields("HES")
                      RST2.Fields("HES_M") = rst.Fields("HES_M")
                      RST2.Fields("DATE_S") = rst.Fields("DATE_S")
                      RST2.Fields("HES_K") = rst.Fields("HES_K")
                      RST2.Fields("SHARH") = "مانده تا اين تاريخ"
                      If rst.Fields("TS") = 1 Then
                          RST2.Fields("BED") = Round(rst.Fields("MAND"))
                          RST2.Fields("BES") = 0
                      Else
                          RST2.Fields("BEs") = Round(rst.Fields("MAND"))
                          RST2.Fields("BEd") = 0
                      End If
                      RST2.Fields("MAND") = Round(rst.Fields("MAND"))
                      RST2.Fields("TS") = rst.Fields("TS")
                      RST2.update
                      CHK = False
                  End If
                  While rst.Fields("HES") = HESM
                     rst.MovePrevious
                     If rst.BOF Then
                       GoTo A5
                     End If
                  Wend
                  rst.MoveNext
              End If
              rst.MovePrevious
A5:
        Wend
        RST2.Close
        rst.Close
        DoCmd.OpenReport "R_DAFTAR_AIS", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
  End If
End Sub

Private Sub Command6_Click()
                      On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub



Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.restore
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
