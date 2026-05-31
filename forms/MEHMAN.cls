Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
  If ID = 0 Or IsNull(ID) Then
    rst.Open "SELECT  MAX(ID) AS MaxOfID FROM  dbo.MEHMAN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.ID = 1
    Else
        Me.ID = rst.Fields(0) + 1
    End If
    rst.Close
    Set rst = New ADODB.Recordset
  End If
  If IsNull(Me.TAFZ) Then
    rst.Open "SELECT MAX(TNUMBER) AS MaxOfTNUMBER FROM dbo.TDETA_HES WHERE  (N_KOL = " & [Forms]![BASEKNOW]![BEDEHKAR] & " ) AND (NUMBER = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
       Call CREATHES(Forms![BASEKNOW]![BEDEHKAR], 1, 1, Me.MNAME & " " & Me.MFAMILY & " - " & Me.ROOMN)
       Me.TAFZ = 1
    Else
       Me.TAFZ = rst.Fields(0) + 1
       If Not ISHESAB(Forms![BASEKNOW]![BEDEHKAR], 1, Me.TAFZ) Then
            Err.Clear
            Call CREATHES(Forms![BASEKNOW]![BEDEHKAR], 1, Me.TAFZ, Me.MNAME & " " & Me.MFAMILY & " " & Me.SAZMANM & " - " & Me.ROOMN)
            If Err <> 0 Then
                RST2.Open "SELECT     COUNT(NAME) AS NCOUNT FROM dbo.TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![BEDEHKAR] & ") AND (NUMBER = 1) AND (NAME ='" & Me.MNAME & " " & Me.MFAMILY & " " & Me.SAZMANM & " - " & Me.ROOMN & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                Call CREATHES(Forms![BASEKNOW]![BEDEHKAR], 1, Me.TAFZ, Me.MNAME & " " & Me.MFAMILY & " " & RST2.Fields(0) + 1 & " " & Me.SAZMANM & " - " & Me.ROOMN)
                RST2.Close
            End If
       Else
           Me.TAFZ = Me.TAFZ + 1
            Call CREATHES(Forms![BASEKNOW]![BEDEHKAR], 1, Me.TAFZ, Me.MNAME & " " & Me.MFAMILY & " " & Me.SAZMANM & " - " & Me.ROOMN)
       End If
    End If
    rst.Close
    Set rst = New ADODB.Recordset
  End If
    rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "room_num = " & Me.ROOMN
    Set rst = rst
    If Not rst.RecordCount = 0 Then
      If IsNull(Me.KDATE) Then
        rst.Fields("rflag") = 1
        rst.update
      End If
    End If
End Sub

Private Sub Form_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "RACKROOM"
   DoCmd.OpenForm "RACKROOM"
End Sub

Private Sub Form_Current()
                      On Error Resume Next
  If IsNull(Me.ID) Then
     Me.MEHMAN_sub.Locked = True
  Else
      Me.MEHMAN_sub.Locked = False
  End If
   Me.ROOMN.Requery
End Sub

Private Sub Form_Load()
                    On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Me.MNAME.SetFocus
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub KDATE_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   If IsNull(KDATE) Then
    rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "room_num = " & Me.ROOMN
    Set rst = rst
    If Not rst.RecordCount = 0 Then
      If IsNull(Me.KDATE) Then
        rst.Fields("rflag") = 1
        rst.update
      End If
    End If
  End If
End Sub

Private Sub KDATE_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
                      On Error Resume Next: CANCEL = CHEKDATEM(KDATE, True)
End Sub

Private Sub KDATE_DblClick(CANCEL As Integer)
                    On Error Resume Next
  Dim rst As New ADODB.Recordset
    KDATE = FARSIDATE(DATE)
    KTIME = Hour(Now()) * 100 + Minute(Now())
    rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "room_num = " & Me.ROOMN
    Set rst = rst
    If Not rst.RecordCount = 0 Then
      If Not IsNull(Me.KDATE) Then
        rst.Fields("rflag") = 4
        rst.update
      End If
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub KTIME_DblClick(CANCEL As Integer)
                    On Error Resume Next
  Dim rst As New ADODB.Recordset
    KDATE = FARSIDATE(DATE)
    KTIME = Hour(Now()) * 100 + Minute(Now())
    rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "room_num = " & Me.ROOMN
    Set rst = rst
    If Not rst.RecordCount = 0 Then
      If Not IsNull(Me.KDATE) Then
        rst.Fields("rflag") = 4
        rst.update
      End If
    End If
End Sub

Private Sub MABLPISH_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub MEHMAN_sub_Enter()
                      On Error Resume Next
  If IsNull(Me.ID) Then
     Me.MEHMAN_sub.Locked = True
  Else
      Me.MEHMAN_sub.Locked = False
  End If
   Me.ROOMN.Requery
End Sub

Private Sub ROOMN_AfterUpdate()
  Dim rst As New ADODB.Recordset, i As Integer

  If IsNull(Me.TAFZ) Then
  Else
    rst.Open "SELECT     TNUMBER, NAME FROM dbo.TDETA_HES WHERE (N_KOL = " & [Forms]![BASEKNOW]![BEDEHKAR] & ") AND (NUMBER = 1) AND (TNUMBER = " & Me.TAFZ & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
       Call CREATHES(Forms![BASEKNOW]![BEDEHKAR], 1, Me.TAFZ, Me.MNAME & " " & Me.MFAMILY & " - " & Me.ROOMN)
    Else
       rst.Fields(1) = Me.MNAME & " " & Me.MFAMILY & " - " & Me.ROOMN
       rst.update
    End If
    rst.Close
    Set rst = New ADODB.Recordset
  End If
rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
rst.Filter = "room_num = " & Me.ROOMN
Set rst = rst
If Not rst.RecordCount = 0 Then
  If IsNull(Me.KDATE) Then
    rst.Fields("rflag") = 1
    rst.update
  End If
End If
End Sub

Private Sub VDATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(VDATE, True)
End Sub
