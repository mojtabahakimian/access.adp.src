Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub CODE_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like N'%" & Me.CODE.Text & "%'"
  Else
    Me.CODE = Me.CODE.Text
  End If
  Response = acDataErrContinue

End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Me.MDA = Round(Me.MORAKH / 440)
  Me.MHOU = Round(((Me.MORAKH / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Round(((((Me.MORAKH / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "GHARA", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End Sub

Private Sub GHEND_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(GHEND, False)
End Sub

Private Sub GHSTART_AfterUpdate()
                  On Error Resume Next
  Me.MORAKH = MORAKH_MAND(Me.GHSTART, Me.CODE)
  Me.MDA = Round(Me.MORAKH / 440)
  Me.MHOU = Round(((Me.MORAKH / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Round(((((Me.MORAKH / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
  If Me.MORAKH = 110110110 Then
    Me.MORAKH = 0
    Me.MDA = 0
    Me.MHOU = 0
    Me.MMINU = 0
  End If
End Sub

Private Sub GHSTART_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(GHSTART, False)
End Sub

Private Sub MDA_AfterUpdate()
                     On Error Resume Next
  Me.MORAKH = Me.MDA * 440 + Me.MHOU * 60 + Me.MMINU
End Sub

Private Sub MHOU_AfterUpdate()
                     On Error Resume Next
  Me.MORAKH = Me.MDA * 440 + Me.MHOU * 60 + Me.MMINU
End Sub

Private Sub MMINU_AfterUpdate()
                     On Error Resume Next
  Me.MORAKH = Me.MDA * 440 + Me.MHOU * 60 + Me.MMINU
End Sub
