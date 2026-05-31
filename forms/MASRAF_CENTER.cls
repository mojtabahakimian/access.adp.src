Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_MANF.FNUMB) AS MaxOfFNUMB FROM HEAD_MANF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.FNUMB = 1
    Else
        Me.FNUMB = rst.Fields(0) + 1
    End If
       rst.Close
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                      On Error Resume Next
    Me.NUMBER.RowSource = "SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & Me.N_KOL & "));"
    Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME;"
    Me.NUMBER.Requery
    Me.TNUMBER.Requery
    Call SETSECURITY(Me.NAME, "MARM", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub N_KOL_AfterUpdate()
                      On Error Resume Next
    Me.NUMBER.RowSource = "SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & Me.N_KOL & "))"
    Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME"
    Me.NUMBER.Requery
    Me.TNUMBER.Requery
End Sub

Private Sub N_KOL_GotFocus()
                      On Error Resume Next
 If Not IsNull(Me.N_KOL) And Me.N_KOL <> "" Then
    Me.NUMBER.RowSource = "SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & Me.N_KOL & "))"
    Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME"
    Me.NUMBER.Requery
    Me.TNUMBER.Requery
End If
End Sub

Private Sub N_KOL_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If NewData > 0 Then
    Me.N_KOL = NewData
  End If
   Response = acDataErrContinue
End Sub

Private Sub NUMBER_AfterUpdate()
                      On Error Resume Next
    Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME"
    Me.TNUMBER.Requery
End Sub

Private Sub NUMBER_GotFocus()

                      On Error Resume Next
  If Not IsNull(Me.NUMBER) And Me.NUMBER <> "" Then
     If Not IsNull(Me.N_KOL) And Me.N_KOL <> "" Then
        Me.NUMBER.RowSource = "SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & Me.N_KOL & "))"
        Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME"
        Me.NUMBER.Requery
        Me.TNUMBER.Requery
    End If
  End If
End Sub

Private Sub NUMBER_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If NewData > 0 Then
    Me.NUMBER = NewData
  End If
   Response = acDataErrContinue
End Sub

Private Sub TNUMBER_GotFocus()
  If Not IsNull(Me.TNUMBER) And Me.TNUMBER <> "" Then
    If Not IsNull(Me.NUMBER) And Me.NUMBER <> "" Then
       If Not IsNull(Me.N_KOL) And Me.N_KOL <> "" Then
          Me.NUMBER.RowSource = "SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & Me.N_KOL & "))"
          Me.TNUMBER.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Me.N_KOL & ") AND ((TDETA_HES.NUMBER)=" & Me.NUMBER & ")) ORDER BY TDETA_HES.NAME"
          Me.NUMBER.Requery
          Me.TNUMBER.Requery
      End If
    End If
  End If
End Sub

Private Sub TNUMBER_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If NewData > 0 Then
    Me.TNUMBER = NewData
  End If
   Response = acDataErrContinue
End Sub
