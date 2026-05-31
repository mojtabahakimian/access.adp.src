Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String, BEDEH, BESTA, max_ns, CON As Double
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, SHRST As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.DT2) Or IsNull(Me.DT1) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If IsNull(Me.KALA) Then
     MKALA = "%"
  Else
    MKALA = Me.KALA
  End If
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
SHRST.AddNew
SHRST.Fields("N_S") = max_ns
SHRST.Fields("DATE_S") = FARSIDATE(DATE)
 If Me.OpenArgs = "AM" Then
    SHRST.Fields("SHARH_S") = RIGHT(" سند عملكرد كالاها از تاريخ  " & Format(Me.DT1, "####/##/##") & "تا تاريخ " & Format(Me.DT2, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 0
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    rst3.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Close
    Set rst = New ADODB.Recordset
      rst.Open "SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MOGH FROM dbo.DEED_HED INNER JOIN  dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " And " & Me.DT2 & ") AND (dbo.DEED_DTL.HES_T <> 99999996 AND dbo.DEED_DTL.HES_T <> 99999998 AND dbo.DEED_DTL.HES_T <> 99999999 AND dbo.DEED_DTL.HES_T <> 99999997) AND (dbo.DEED_DTL.HES_K = " & [Forms]![BASEKNOW]![CONKAL] & " ) GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T HAVING (SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      rst.MoveLast
      rst.MoveFirst
      DoCmd.restore
      DoCmd.OpenForm "GUG"
      DoCmd.MoveSize 3200, 3000, 5500, 2000
      Forms![GUG]![Text0] = "n"
      Forms![GUG]![Text0].Visible = True
      Forms![GUG]![Box7].Visible = True
      Forms![GUG].Form.Refresh
      Forms![GUG].Form.Repaint
      Forms![GUG]![SNUM] = rst.RecordCount
      CON = rst.RecordCount / 50
      If CON < 1 Then
         CON = 1
      End If
      i = 1
     While Not rst.EOF() And rst.RecordCount > 0
           Forms![GUG]![num] = i
           rst3.AddNew
           rst3.Fields("N_S") = max_ns
           rst3.Fields("HES_K") = [Forms]![BASEKNOW]![CONKAL]
           rst3.Fields("HES_M") = rst.Fields("HES_M")
           rst3.Fields("HES_T") = rst.Fields("HES_T")
           rst3.Fields("HES") = [Forms]![BASEKNOW]![CONKAL] & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_T")
           rst3.Fields("SHARH") = "انتقال به حساب عملكرد "
           If rst.Fields("MOGH") > 0 Then
               rst3.Fields("BES") = rst.Fields("MOGH")
           Else
               rst3.Fields("BED") = rst.Fields("MOGH") * -1
           End If
           rst3.update
           Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
           Forms![GUG].Form.Refresh
           Forms![GUG]![Text2].Requery
           Forms![GUG].Form.Repaint
           If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst.Fields("HES_M")), CLng(1 & rst.Fields("HES_T"))) Then
             Err.Clear
             Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst.Fields("HES_M")), CLng(1 & rst.Fields("HES_T")), "مغايرت وزن " & GETKALANAME(rst.Fields("HES_T")))
           End If
           rst3.AddNew
           rst3.Fields("N_S") = max_ns
           rst3.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
           rst3.Fields("HES_M") = rst.Fields("HES_M")
           rst3.Fields("HES_T") = 1 & rst.Fields("HES_T")
           rst3.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-" & rst.Fields("HES_M") & "-" & 1 & rst.Fields("HES_T")
           rst3.Fields("SHARH") = "انتقال از حساب  كنترل كالاي درجريان ساخت"
           If rst.Fields("MOGH") > 0 Then
               rst3.Fields("BED") = rst.Fields("MOGH")
           Else
               rst3.Fields("BES") = rst.Fields("MOGH") * -1
           End If
           rst3.update
           Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
           Forms![GUG].Form.Refresh
           Forms![GUG]![Text2].Requery
           Forms![GUG].Form.Repaint
          rst.MoveNext
           If (i Mod CON) = 0 Then
             Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
             Forms![GUG].Form.Refresh
             Forms![GUG]![Text0].Requery
             Forms![GUG].Form.Repaint
           End If
           i = i + 1
           Forms![GUG]![Text2] = ""
    Wend
    rst.Close
    Set rst = New ADODB.Recordset
      rst.Open "SELECT DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T, Sum([BED]-[bes]) AS MOGH FROM DEED_HED INNER JOIN DEED_DTL ON (DEED_HED.N_S = DEED_DTL.N_S) AND (DEED_HED.N_S = DEED_DTL.N_S) WHERE (((DEED_HED.DATE_S) Between " & Me.DT1 & " And " & Me.DT2 & ")) GROUP BY DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T HAVING (((DEED_DTL.HES_K)=" & [Forms]![BASEKNOW]![CONKAL] & ") AND ((DEED_DTL.HES_T)=99999998 Or (DEED_DTL.HES_T)=99999999) AND ((Sum([BED]-[bes]))<>0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      rst.MoveLast
      rst.MoveFirst
      Forms![GUG]![Text0] = "n"
      Forms![GUG]![Text0].Visible = True
      Forms![GUG]![Box7].Visible = True
      Forms![GUG].Form.Refresh
      Forms![GUG].Form.Repaint
      Forms![GUG]![SNUM] = rst.RecordCount
      CON = rst.RecordCount / 50
      If CON < 1 Then
         CON = 1
      End If
      i = 1
      While Not rst.EOF() And rst.RecordCount > 0
           Forms![GUG]![num] = i
           rst3.AddNew
           rst3.Fields("N_S") = max_ns
           rst3.Fields("HES_K") = [Forms]![BASEKNOW]![CONKAL]
           rst3.Fields("HES_M") = rst.Fields("HES_M")
           rst3.Fields("HES_T") = rst.Fields("HES_T")
           rst3.Fields("SHARH") = "انتقال به حساب عملكرد "
           rst3.Fields("HES") = [Forms]![BASEKNOW]![CONKAL] & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_T")
           If rst.Fields("MOGH") > 0 Then
               rst3.Fields("BES") = rst.Fields("MOGH")
           Else
               rst3.Fields("BED") = rst.Fields("MOGH") * -1
           End If
           rst3.update
           Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
           Forms![GUG].Form.Refresh
           Forms![GUG]![Text2].Requery
           Forms![GUG].Form.Repaint
           If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst.Fields("HES_M")), CLng(rst.Fields("HES_T"))) Then
             Err.Clear
             If rst.Fields("HES_T") = 99999999 Then
               Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst.Fields("HES_M")), rst.Fields("HES_T"), "مغايرت دستمزد " & rst.Fields("HES_T"))
             Else
               Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst.Fields("HES_M")), rst.Fields("HES_T"), "مغايرت سربار " & GETKALANAME(rst.Fields("HES_T")))
             End If
           End If
           rst3.AddNew
           rst3.Fields("N_S") = max_ns
           rst3.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
           rst3.Fields("HES_M") = rst.Fields("HES_M")
           rst3.Fields("HES_T") = rst.Fields("HES_T")
           rst3.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_T")
           rst3.Fields("SHARH") = "انتقال از حساب  كنترل كالاي درجريان ساخت"
           If rst.Fields("MOGH") > 0 Then
               rst3.Fields("BED") = rst.Fields("MOGH")
           Else
               rst3.Fields("BES") = rst.Fields("MOGH") * -1
           End If
           rst3.update
           Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
           Forms![GUG].Form.Refresh
           Forms![GUG]![Text2].Requery
           Forms![GUG].Form.Repaint
           rst.MoveNext
           If (i Mod CON) = 0 Then
             Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
             Forms![GUG].Form.Refresh
             Forms![GUG]![Text0].Requery
             Forms![GUG].Form.Repaint
           End If
           i = i + 1
           Forms![GUG]![Text2] = ""
     Wend
     DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "سند عملكرد صادر گرديد..."
     DoCmd.Close acForm, "gug"
  Else
    SHRST.Fields("SHARH_S") = RIGHT(" سند تعديل قيمت تمام شده كالاها از تاريخ  " & Format(Me.DT1, "####/##/##") & "تا تاريخ " & Format(Me.DT2, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 0
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.update
     rst3.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "SELECT DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T, Sum([BED]-[bes]) AS MOGH FROM DEED_HED INNER JOIN DEED_DTL ON (DEED_HED.N_S = DEED_DTL.N_S) AND (DEED_HED.N_S = DEED_DTL.N_S) WHERE (((DEED_HED.DATE_S) Between " & Me.DT1 & " And " & Me.DT2 & " )) GROUP BY DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T HAVING (((DEED_DTL.HES_K)= " & [Forms]![BASEKNOW]![AMALKARD] & ") AND ((Sum([BED]-[bes]))<>0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      rst.MoveLast
      rst.MoveFirst
      DoCmd.restore
      DoCmd.OpenForm "GUG"
      DoCmd.MoveSize 3200, 3000, 5500, 2000
      Forms![GUG]![Text0] = "n"
      Forms![GUG]![Text0].Visible = True
      Forms![GUG]![Box7].Visible = True
      Forms![GUG].Form.Refresh
      Forms![GUG].Form.Repaint
      Forms![GUG]![SNUM] = rst.RecordCount
      CON = rst.RecordCount / 50
      If CON < 1 Then
         CON = 1
      End If
      i = 1
     While Not rst.EOF()
           Forms![GUG]![num] = i
           If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(rst.Fields("HES_M")), CLng(rst.Fields("HES_M"))) Then
             Err.Clear
             Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(rst.Fields("HES_M")), rst.Fields("HES_M"), "قيمت تمام شده  " & GETKALANAME(rst.Fields("HES_T")))
           End If
           rst3.AddNew
           rst3.Fields("N_S") = max_ns
           rst3.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
           rst3.Fields("HES_M") = rst.Fields("HES_M")
           rst3.Fields("HES_T") = rst.Fields("HES_M")
           rst3.Fields("HES") = [Forms]![BASEKNOW]![GHEYMAT] & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_M")
           rst3.Fields("SHARH") = "انتقال از حساب عملكرد "
           If rst.Fields("MOGH") > 0 Then
               rst3.Fields("BED") = rst.Fields("MOGH")
           Else
               rst3.Fields("BES") = rst.Fields("MOGH") * -1
           End If
           Err.Clear
           rst3.update
           If Err = 0 Then
                Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
                Forms![GUG].Form.Refresh
                Forms![GUG]![Text2].Requery
                Forms![GUG].Form.Repaint
                rst3.AddNew
                rst3.Fields("N_S") = max_ns
                rst3.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                rst3.Fields("HES_M") = rst.Fields("HES_M")
                rst3.Fields("HES_T") = rst.Fields("HES_T")
                rst3.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-" & rst.Fields("HES_M") & "-" & rst.Fields("HES_T")
                rst3.Fields("SHARH") = "انتقال به حساب  قيمت تمام  شده"
                If rst.Fields("MOGH") > 0 Then
                    rst3.Fields("BES") = rst.Fields("MOGH")
                Else
                    rst3.Fields("BED") = rst.Fields("MOGH") * -1
                End If
                rst3.update
            End If
            Err.Clear
            Forms![GUG]![Text2] = Forms![GUG]![Text2] + "nnnnn"
            Forms![GUG].Form.Refresh
            Forms![GUG]![Text2].Requery
            Forms![GUG].Form.Repaint
            rst.MoveNext
           If (i Mod CON) = 0 Then
             Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
             Forms![GUG].Form.Refresh
             Forms![GUG]![Text0].Requery
             Forms![GUG].Form.Repaint
           End If
           i = i + 1
           Forms![GUG]![Text2] = ""
    Wend
     DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "سند تعديل قيمت تمام شده صادر گرديد..."
     DoCmd.Close acForm, "gug"
  End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  If Me.OpenArgs = "RK" Then
    Me.DTT = 1
   Else
     Me.DTT = 2
   End If
End Sub

Private Sub KALA_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Me.KALA = NewData
  Response = acDataErrContinue
End Sub
