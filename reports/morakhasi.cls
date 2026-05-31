Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit



Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
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
  Call Showemza(Me, "SGN0137TX", "SGN0237TX", "SGN0337TX")

End Sub
