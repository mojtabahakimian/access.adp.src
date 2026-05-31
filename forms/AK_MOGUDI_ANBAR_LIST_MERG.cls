Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Detail_MouseMove(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
   If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR_merg"
End Sub
Private Sub Form_Current()
   If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_MouseMove(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
   If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub
Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Me.MAND.DecimalPlaces = Forms![BASEKNOW]![DIG]
   If Forms![BASEKNOW]![UGRP] = 4 Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.RecordSource = "SELECT     CODE, GRCOD, GRNAME, KK, SUM(MABLK) AS MABLK, SUM(MAND) AS MAND, SUM(MANDF) AS MANDF, N_FANI, NAME, NAMES, NESBAT, TOZIH, VAZN, SUM(VAZNK) AS VAZNK, VCOD, AVG(FII) AS FII FROM AKMOGUDI_KOL_ANBAR(99999999, '%') AKMOGUDI_KOL_ANBAR WHERE  " & Me.OpenArgs & " GROUP BY CODE, GRCOD, GRNAME, KK, N_FANI, NAME, NAMES, NESBAT, TOZIH, VCOD, VAZN"
   Me.Refresh
End Sub
Private Sub MABLK_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "MOGHAYERAT_ANABR", acFormDS
End Sub
