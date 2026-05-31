Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
 '   Me.InputParameters = "@F_MENU_DATE_DT1=" & Forms![F_MENU_DATE]![DT1] & ",@F_MENU_DATE_DT2=" & Forms![F_MENU_DATE]![DT2] & ",@dep ='" & Forms![GOZARESH_FROOSH_MAHSUL]![DEPATMAN] & "' "
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
