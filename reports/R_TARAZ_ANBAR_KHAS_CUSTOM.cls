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
   Me.MEGHKHM.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHENVOM.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHTOM.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHSAYER.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHSAYES.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHENKRM.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MEGHFRM.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.MOG.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
