Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub DATE_N_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "HEAD_SERCH_MAIN"
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub



Private Sub Form_Open(CANCEL As Integer)
     On Error Resume Next
  '  Dim itemm
  '  For Each itemm In Me.Form
  '      Me(itemm.NAME).ColumnWidth = -2
  '  Next
    Me.MEGHk.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH_MAR.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
