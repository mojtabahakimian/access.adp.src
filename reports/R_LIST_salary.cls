Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
   If Not IsNull(Forms![BASEKNOW]![LECOL1]) And Forms![BASEKNOW]![LECOL1] <> "" Then
     Me.SAYER1.CAPTION = Forms![BASEKNOW]![LECOL1]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL2]) And Forms![BASEKNOW]![LECOL2] <> "" Then
     Me.SAYER2.CAPTION = Forms![BASEKNOW]![LECOL2]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL3]) And Forms![BASEKNOW]![LECOL3] <> "" Then
     Me.SAYER3.CAPTION = Forms![BASEKNOW]![LECOL3]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL4]) And Forms![BASEKNOW]![LECOL4] <> "" Then
     Me.SAYER4.CAPTION = Forms![BASEKNOW]![LECOL4]
   End If
   If Not IsNull(Forms![BASEKNOW]![LKCOL1]) And Forms![BASEKNOW]![LKCOL1] <> "" Then
     Me.KOSORAT1.CAPTION = Forms![BASEKNOW]![LKCOL1]
   End If
End Sub

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_BIM"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
