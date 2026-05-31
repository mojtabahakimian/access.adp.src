Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim BGIDD As Long

Private Sub Command19_Click()
    DoCmd.OpenForm "buget5", , , , , acDialog, BGIDD
End Sub

Private Sub Command3_Click()
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Command4_Click()
             On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    If Not IsLoaded("BUGET_MAIN_list") Then
        CurrentProject.Connection.Execute ("update BUGET_MAIN SET BGPCH1 =" & Me.BGPCH1 & ", BGPCH2 =" & Me.BGPCH2 & ", BGPCH3 =" & Me.BGPCH3 & ", BGPCH4 =" & Me.BGPCH4 & ", BGPCH5 =" & Me.BGPCH5 & ", BGPCH6 =" & Me.BGPCH6 & ", BGPCH7 =" & Me.BGPCH7 & ", BGPCH8 =" & Me.BGPCH8 & ", BGPCH9 =" & Me.BGPCH9 & ", BGPCH10 =" & Me.BGPCH10 & ", BGPCH11 =" & Me.BGPCH11 & ", BGPCH12 =" & Me.BGPCH12 & ", BGPCASH =" & Me.BGPCASH & ", BGPCH0 =" & Me.BGPCH0 & ", NTDZBMCHDABD = '" & Me.NTDZBMCHDABD & "' WHERE  BGID = " & BGIDD)
    End If
    DoCmd.OpenForm "buget2", , , , , acDialog, BGIDD
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"


End Sub

Private Sub Form_Open(CANCEL As Integer)
  '  DoCmd.Maximize:Forms![baseknow]![hhwin] = Me.hWnd
            On Error Resume Next
    If Me.RecordsetClone.RecordCount = 0 Then
        CurrentProject.Connection.Execute ("INSERT [BUGET_DEFAULT] VALUES(1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,''," & Forms![BASEKNOW]![USERCOD] & ",GETDATE()),''")
        Me.USERNAME = UCurrentUser()
        Me.Requery
    Else
        Me.USERNAME = GETUSERNAME(Me.USERID)
    End If
      BGIDD = IIf(IsNull(Me.OpenArgs), 1, Me.OpenArgs)
      Me.RecordSource = "select * from BUGET_MAIN WHERE BGID = " & BGIDD
    If IsLoaded("BUGET_MAIN_list") Then
        Me.AllowEdits = False
        Me.Command19.Enabled = False
    Else
        Me.AllowEdits = True
        Me.Command19.Enabled = True
    End If

End Sub
