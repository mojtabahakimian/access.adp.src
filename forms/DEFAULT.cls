Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command147_Click()
                        On Error Resume Next
                  On Error Resume Next
               If IsNull(Me.TFSAZMAN) Or IsNull(Me.SHIFT) Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "واحد و شيفت نميتواند خالي باشد"
                    Exit Sub
               End If
       DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Command148_Click()
                  On Error Resume Next
               If IsNull(Me.TFSAZMAN) Or IsNull(Me.SHIFT) Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "واحد و شيفت نميتواند خالي باشد"
                    Exit Sub
               End If
       DoCmd.RunCommand acCmdSaveRecord
       DoCmd.OpenForm Me.NAME, , , , , acHidden

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsLoaded("baseknow") Then
        Me.RecordSource = "SELECT * FROM [DEFAULTDEP] WHERE USERID = " & Forms![BASEKNOW]![USERCOD]
        Me.Refresh
        If Me.RecordsetClone.RecordCount = 0 Then
            CurrentProject.Connection.Execute ("insert into [DEFAULTDEP](USERID,TFSAZMAN,SHIFT)  values (" & Forms![BASEKNOW]![USERCOD] & ",1,1)")
            Me.Refresh
        End If
    End If
    If LETSGO("DEFA") Then
        Me.AllowEdits = True
    End If
    rst.Open "select * from SALA_DTL where idd = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
        Me.menup = rst.Fields("menup")
    End If
End Sub

Private Sub Form_Timer()
                  On Error Resume Next
       DoCmd.RunCommand acCmdSaveRecord
       DoCmd.OpenForm Me.NAME, , , , , acHidden
Me.TimerInterval = 0
End Sub

Private Sub menup_BeforeUpdate(CANCEL As Integer)
             On Error Resume Next
    Dim rst As New ADODB.Recordset, menup As Integer
    If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
        DoCmd.OpenForm "MESAG", , , , , acDialog, "اين گزينه در نرم افزار مستر کارکت در دسترس است"
        Me.menup = 0
    Else
        rst.Open "select * from SALA_DTL where idd = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount <> 0 Then
            rst.Fields("menup") = Me.menup
            rst.update
        End If
        Select Case Me.menup
          Case 0, -1: DoCmd.Close acForm, "BOARD1": DoCmd.Close acForm, "form2"
                      SELECTBAR (rst.Fields("GRSAL"))
          Case 1:
                  Call HIDBARS
                  DoCmd.OpenForm "BOARD1"
                  DoCmd.Close acForm, "form2"
          Case 2: DoCmd.Close acForm, "BOARD1"
                  DoCmd.OpenForm "form2"
        End Select
    End If
End Sub
Public Sub SELECTBAR(GRP As Integer)
     DoCmd.DeleteObject acMacro, "AUTOKEYS"
    Select Case GRP
      Case 1:    DoCmd.ShowToolbar "DIALOG"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYMAIN"
      Case 2:    DoCmd.ShowToolbar "HESAB"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYHES"
      Case 3:    DoCmd.ShowToolbar "FROOSH"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 4:    DoCmd.ShowToolbar "ANBAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYANB"
      Case 5:    DoCmd.ShowToolbar "KARKOZ"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYKAR"
      Case 6:    DoCmd.ShowToolbar "HOT"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 7:    DoCmd.ShowToolbar "TALAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 8:    DoCmd.ShowToolbar "ANBAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYANB"
      Case 9:     DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYOFF"
      Case 10
            DoCmd.ShowToolbar "DIALOG"
            DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYMAIN"
    End Select
End Sub

Public Sub HIDBARS()
                      On Error Resume Next
     DoCmd.ShowToolbar "DIALOG", acToolbarNo
     DoCmd.ShowToolbar "HESAB", acToolbarNo
     DoCmd.ShowToolbar "FROOSH", acToolbarNo
     DoCmd.ShowToolbar "ANBAR", acToolbarNo
     DoCmd.ShowToolbar "KARKOZ", acToolbarNo
     DoCmd.ShowToolbar "HOT", acToolbarNo
     DoCmd.ShowToolbar "TALAR", acToolbarNo
End Sub
