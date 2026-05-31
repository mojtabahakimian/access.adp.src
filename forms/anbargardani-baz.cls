Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub aben_AfterUpdate()
 Call CALCU
End Sub

Private Sub afbfr_AfterUpdate()
 Call CALCU
End Sub

Private Sub afbkh_AfterUpdate()
 Call CALCU
End Sub

Private Sub affr_AfterUpdate()
 Call CALCU

End Sub

Private Sub afkh_AfterUpdate()
 Call CALCU

End Sub

Private Sub ANBAR_AfterUpdate()
 Call CALCU
End Sub

Private Sub CODE_AfterUpdate()
 Call CALCU
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
   If Me.CODE.Text = "-" Or Me.CODE.Text = "+" And Not IsNull(Me.ANBAR) Then
'     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
 Else
     Me.CODE = NewData
 End If
  Response = acDataErrContinue
End Sub

Public Sub CALCU()
                 On Error Resume Next
 If Not IsNull(Me.ANBAR) And Not IsNull(Me.CODE) And Not IsNull(Me.moeb) And Not IsNull(Me.moen) Then
  Dim rst1 As New ADODB.Recordset, MFR As Double, MKH As Double, MBKH As Double, MBFR As Double, MAZA As Double, MBA As Double
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.affr & ") AND (TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MFR = rst1.Fields(0)
  Else
    MFR = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.afkh & ") AND (TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MKH = rst1.Fields(0)
  Else
    MKH = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS MBKH FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG - 2 = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.TAG = 3) AND (dbo.HEAD_LST.NUMBER <= " & Me.afbkh & ") AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBKH = rst1.Fields(0)
  Else
    MBKH = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS MBKH FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG - 2 = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.TAG = 4) AND (dbo.HEAD_LST.NUMBER <= " & Me.afbfr & ") AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBFR = rst1.Fields(0)
  Else
    MBFR = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.aben & ") AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ") AND (TAG = 5) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MAZA = rst1.Fields(0)
  Else
    MAZA = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(MEGHk) AS Expr1 FROM dbo.INVO_LST WHERE     (CODE = N'" & Me.CODE & "') AND (ANBARF = " & Me.ANBAR & ") AND (TAG = 5) AND (NUMBER <=" & Me.aben & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBA = rst1.Fields(0)
  Else
    MBA = 0
  End If
  Me.moeb = Me.moen - MKH + MFR - MBFR + MBKH + MAZA - MBA
  Else
     DoCmd.OpenForm "MESAG", , , , , acDialog, "پارامترها كافي نيست"
  End If

End Sub

Private Sub Command19_Click()
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.ANBAR) Or IsNull(Me.CODE) Or Me.CODE = "" Then
          DoCmd.OpenForm "MESAG", , , , , acDialog, "پارامترها كافي نيست"
    Else
        rst.Open "SELECT     * from stuf_fsk  WHERE     (CODE = N'" & Me.CODE & "') AND (ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            rst.Fields("MOGODI_A") = Me.moeb
            rst.Fields("MABL_A") = Me.moeb * rst.Fields("FI_A")
            rst.update
            DoCmd.OpenForm "MESAG", , , , , acDialog, "موجودي اصلاح شد."
        Else
                DoCmd.OpenForm "MESAG", , , , , acDialog, "عمليات ناموفق بود."
        End If
         rst.Close
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub moeb_AfterUpdate()
                 On Error Resume Next
 If Not IsNull(Me.ANBAR) And Not IsNull(Me.CODE) And Not IsNull(Me.moeb) And Not IsNull(Me.moen) Then
  Dim rst1 As New ADODB.Recordset, MFR As Double, MKH As Double, MBKH As Double, MBFR As Double, MAZA As Double, MBA As Double
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.affr & ") AND (TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MFR = rst1.Fields(0)
  Else
    MFR = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.afkh & ") AND (TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MKH = rst1.Fields(0)
  Else
    MKH = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS MBKH FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG - 2 = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.TAG = 3) AND (dbo.HEAD_LST.NUMBER <= " & Me.afbkh & ") AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBKH = rst1.Fields(0)
  Else
    MBKH = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS MBKH FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG - 2 = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.TAG = 4) AND (dbo.HEAD_LST.NUMBER <= " & Me.afbfr & ") AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBFR = rst1.Fields(0)
  Else
    MBFR = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT  SUM(MEGHk) AS mek FROM dbo.INVO_LST WHERE     (NUMBER <= " & Me.aben & ") AND (dbo.INVO_LST.ANBAR = " & Me.ANBAR & ") AND (TAG = 5) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MAZA = rst1.Fields(0)
  Else
    MAZA = 0
  End If
  rst1.Close
  Set rst = New ADODB.Recordset
  rst1.Open "SELECT     SUM(MEGHk) AS Expr1 FROM dbo.INVO_LST WHERE     (CODE = N'" & Me.CODE & "') AND (ANBARF = " & Me.ANBAR & ") AND (TAG = 5) AND (NUMBER <=" & Me.aben & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst1.RecordCount > 0 And Not IsNull(rst1.Fields(0)) Then
    MBA = rst1.Fields(0)
  Else
    MBA = 0
  End If
  Me.moen = Me.moeb + MKH - MFR + MBFR - MBKH + MAZA + MBA
  Else
     DoCmd.OpenForm "MESAG", , , , , acDialog, "پارامترها كافي نيست"
  End If

End Sub

Private Sub moen_AfterUpdate()
 Call CALCU

End Sub
