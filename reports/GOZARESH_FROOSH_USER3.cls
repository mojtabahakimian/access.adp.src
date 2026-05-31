Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupFooter0_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     COUNT(M_NAGHD) AS SumOfM_NAGHD FROM dbo.HEAD_LST WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & " ) AND (TAG = 13)   AND  (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & " ) AND (M_NAGHD > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.CNAGHD.Visible = True
     Me.CNAGHD.CAPTION = "( " & rst.Fields(0) & " )"
   End If
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     COUNT(TAKHFIF) AS MM FROM dbo.HEAD_LST WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & " ) AND (TAG = 13)  AND  (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & " ) AND (TAKHFIF > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.CTAKH.Visible = True
     Me.CTAKH.CAPTION = "( " & rst.Fields(0) & " )"
     Me.Label137.Visible = True
   End If
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(M_NAGHD) AS naghd, COUNT(M_NAGHD) AS countn FROM dbo.HEAD_LST WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ") AND (TAG = 4) AND   (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (M_NAGHD > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.MNAGHD.Visible = True
     Me.MNAGHD.CAPTION = "( " & Format(rst.Fields(0), "#,###") & " )"
     Me.LMNAGHD.Visible = True
     Me.SANDOGH.Visible = True
     Me.LSANDOGH.Visible = True
     Me.SANDOGH.CAPTION = Format(SumOfM_NAGHD - rst.Fields(0), "#,###")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     COUNT(MABL_VAR) AS MM FROM dbo.HEAD_LST WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & " ) AND (TAG = 13)  AND  (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & " ) AND (MABL_VAR > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.CVAR.Visible = True
     Me.CVAR.CAPTION = "( " & rst.Fields(0) & " )"
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     COUNT(MABL_HAV) AS MM FROM dbo.HEAD_LST WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & " ) AND (TAG = 13) AND  (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & " ) AND (MABL_HAV > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.CHAV.Visible = True
     Me.CHAV.CAPTION = "( " & rst.Fields(0) & " )"
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     COUNT(dbo.PAY_GETD.MABL) AS SumOfMABL FROM         dbo.HEAD_LST INNER JOIN   dbo.PAY_GETD ON dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER AND  dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER AND dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND   dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER AND dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND  dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER WHERE     (DATE_N BETWEEN " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " AND " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & " ) AND  (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & " ) AND   (dbo.HEAD_LST.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And rst.Fields(0) > 0 Then
     Me.CASN.Visible = True
     Me.CASN.CAPTION = "( " & rst.Fields(0) & " )"
   End If
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     TOP 100 PERCENT SUM(SMAB + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD) AS MAND, COUNT(smab + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD) As TEDAD FROM         dbo.GUSER_FACT_PERDEPART_SUB3(" & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & ", " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ", " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ", '%', '%') GUSER_FACT_PERDEPART_SUB3 WHERE     (SMAB + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD > 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.BED.CAPTION = Format(rst.Fields("mand"), "#,###")
     Me.CBED.CAPTION = "( " & rst.Fields("tedad") & " )"
     Me.CBED.Visible = True
   End If
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     TOP 100 PERCENT SUM(SMAB + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD) AS MAND, COUNT(smab + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD) As TEDAD FROM         dbo.GUSER_FACT_PERDEPART_SUB3(" & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & ", " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ", " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ", '%', '%') GUSER_FACT_PERDEPART_SUB3 WHERE     (SMAB + MBAA - TAKHFIF - MABL_VAR - MABL_HAV - CHK + HAZ - M_NAGHD < 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.BES.CAPTION = Format(rst.Fields("mand") * -1, "#,###")
     Me.cbes.CAPTION = "( " & rst.Fields("tedad") & " )"
     Me.cbes.Visible = True
   End If
End Sub

Private Sub Report_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_GOZARESH_FROOSH"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
   If Forms![F_MENU_GOZARESH_FROOSH]![DTL] Then
        Me.Detail.Visible = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
