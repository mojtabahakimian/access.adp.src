Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                       On Error Resume Next
      '1;"دلار آمريکا $";2;"يورو €";3;"درهم UAE";4;"پوند #";5;"ين"
      Select Case Forms![HEAD_LST_FROOSH2_SADER]![ARZKIND]
        Case 1: Me.Label139.CAPTION = "PRICE $"
        Case 2: Me.Label139.CAPTION = "PRICE €"
        Case 3: Me.Label139.CAPTION = "PRICE UAE"
        Case 4: Me.Label139.CAPTION = "PRICE JPY100"
        Case 5: Me.Label139.CAPTION = "PRICE #"
      End Select
     If [Forms]![HEAD_LST_FROOSH2_SADER]![MABL_HAZ] > 0 Then
        Me.Label208.Visible = True
        Me.MABL_HAZ1.Visible = True
     End If

   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Report_Page()
                      On Error Resume Next

    Me.Line (0, 2020)-(40, Me.ScaleHeight), , BF
    Me.Line (535, 2020)-(558, Me.ScaleHeight - 2681), , BF
    Me.Line (4318, 2020)-(4338, Me.ScaleHeight - 2681 + 434), , BF
    Me.Line (5354, 2020)-(5374, Me.ScaleHeight - 2681 + 434), , BF
    Me.Line (6704, 2020)-(6724, Me.ScaleHeight - 2681 + 434), , BF
    Me.Line (8549, 2020)-(8569, Me.ScaleHeight - 2681 + 434), , BF
    Me.Line (10090, 2020)-(10130, Me.ScaleHeight), , BF

End Sub
