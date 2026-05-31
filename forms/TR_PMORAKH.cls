Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub CODE_AfterUpdate()
                     On Error Resume Next
  Call MORAKHDAY_AfterUpdate

End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like '%" & Me.CODE.Text & "%'"
  Else
    Me.CODE = Me.CODE.Text
  End If
  Response = acDataErrContinue
End Sub

Private Sub Command54_Click()
                 On Error Resume Next
    DoCmd.OpenReport "morakhasi", acPreview, , "code = " & Me.CODE & " and MODATE = " & Me.MODATE
End Sub

Private Sub DA_AfterUpdate()
                     On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate
  Call MOenDATE_AfterUpdate
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
              '     On Error Resume Next
  If Not IsNull(Me.CODE) Then
        dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
     '   Call TR("PHOKM", "(HONUM = " & Me.HONUM & ")", dt, 1)
    End If
    Me.AllowDeletions = True
    Me.AllowEdits = True
  End If
  'Call SETSECURITY(Me.NAME, "HOKM", 3)
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
  Me.DA = Fix(Me.MORAKHDAY / 440)
  Me.HOU = Fix(((Me.MORAKHDAY / 440) - Me.DA) * 440 / 60)
  Me.minu = Fix(((((Me.MORAKHDAY / 440) - Me.DA) * 440 / 60) - Me.HOU) * 60)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
   End If

End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "MORAKH", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub HOU_AfterUpdate()
                     On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate
End Sub

Private Sub minu_AfterUpdate()
                     On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate

End Sub

Private Sub MODATE_AfterUpdate()
                     On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
End Sub

Private Sub MODATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(MODATE, False)
End Sub

Private Sub MORAKHDAY_AfterUpdate()
                   On Error Resume Next
  Dim ho As Integer, DD As Integer, mi As Integer
  DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  If Me.MAND = 110110110 Then
  '  Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
End Sub

Private Sub MOSTDATE_AfterUpdate()
Dim DD As Integer
  If MOENDATE > 0 Then
    DD = DIFF(Me.MOSTDATE, Me.MOENDATE) + 1
    If Me.DA <> DD And Me.KINDM <> 2 Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "تعداد روزهاي مرخصي با محدوده تاريخي داده شده تطبيق ندارد آن را اصلاح كردم!"
       Me.DA = DD
       Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
       Call MORAKHDAY_AfterUpdate
    End If
  End If
End Sub

Private Sub MOenDATE_AfterUpdate()
Dim DD As Integer
  If MOSTDATE > 0 Then
    DD = DIFF(Me.MOSTDATE, Me.MOENDATE) + 1
    If Me.DA <> DD And Me.KINDM <> 2 Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "تعداد روزهاي مرخصي با محدوده تاريخي داده شده تطبيق ندارد آن را اصلاح كردم!"
       Me.DA = DD
       Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
       Call MORAKHDAY_AfterUpdate
    End If
  End If
End Sub


Private Sub Form_AfterUpdate()
   If Forms![BASEKNOW]![LOCKFAP] Then
        If Not Me.OKF Then
          Me.OKF = True
        End If
    End If
End Sub
