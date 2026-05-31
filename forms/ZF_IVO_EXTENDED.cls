Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
   Me.ServerFilter = ""
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                             On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
    Me.ID.DefaultValue = Me.OpenArgs
End Sub

Private Sub PARAMS_Click()
                    On Error Resume Next
   Forms![head_lst_khareed1]![MOLAH] = "چربي:" & Me.FLD1 & "- ماده خشک:" & Me.FLD2 & "- رطوبت:" & Me.FLD3 & "- پي اچ:" & Me.FLD4 & "- نمک:" & Me.FLD5 & "- دانسيته:" & Me.FLD6 & "- پروتئين:" & Me.FLD7 & "- انجماد:" & Me.FLD8 & "- اسيد:" & Me.FLD9 & "-الکل:" & Me.FLD10
   Forms![head_lst_rasid]![MOLAH] = "چربي:" & Me.FLD1 & "- ماده خشک:" & Me.FLD2 & "- رطوبت:" & Me.FLD3 & "- پي اچ:" & Me.FLD4 & "- نمک:" & Me.FLD5 & "- دانسيته:" & Me.FLD6 & "- پروتئين:" & Me.FLD7 & "- انجماد:" & Me.FLD8 & "- اسيد:" & Me.FLD9 & "-الکل:" & Me.FLD10
End Sub
