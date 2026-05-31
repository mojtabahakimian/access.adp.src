Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, NAGHDD As Double, NAGHDDP As Double, NAGHDBRP As Double, NAGHDP As Double, HKH As Double, MBAA As Double, TF As Double, GABEL As Double, NAGHD As Double, VARIZ As Double, HAVL As Double, CHK As Double, sql As String
   RST2.Open "SELECT     SUM(dbo.PGET_LST.MABL) AS MABLN FROM         dbo.PGET_HED INNER JOIN   dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID AND dbo.PGET_HED.DATE = dbo.PGET_LST.DATE WHERE     (dbo.PGET_LST.NAHVA = 1) AND (dbo.PGET_LST.NO_AM = 1) AND (dbo.PGET_HED.DATE Between " & Forms!F_MENU_GOZARESH_FROOSH!DT1 & " And " & Forms!F_MENU_GOZARESH_FROOSH!DT2 & ")  AND (dbo.PGET_HED.USER_NAME LIKE N'" & Forms!F_MENU_GOZARESH_FROOSH!USERR & "')  and  (dbo.PGET_HED.DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(dbo.PGET_HED.SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
      If Not IsNull(RST2.Fields("MABLN")) Then
            NAGHDD = RST2.Fields("MABLN")
      Else
            NAGHDD = 0
      End If
   End If
   Set RST2 = New ADODB.Recordset
   RST2.Open "SELECT     SUM(dbo.PGET_LST.MABL) AS MABLN FROM         dbo.PGET_HED INNER JOIN   dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID AND dbo.PGET_HED.DATE = dbo.PGET_LST.DATE WHERE     (dbo.PGET_LST.NAHVA = 1) AND (dbo.PGET_LST.NO_AM = 2) AND (dbo.PGET_HED.DATE Between " & Forms!F_MENU_GOZARESH_FROOSH!DT1 & " And " & Forms!F_MENU_GOZARESH_FROOSH!DT2 & ")  AND (dbo.PGET_HED.USER_NAME LIKE N'" & Forms!F_MENU_GOZARESH_FROOSH!USERR & "')  and  (dbo.PGET_HED.DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(dbo.PGET_HED.SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
      If Not IsNull(RST2.Fields("MABLN")) Then
            NAGHDP = RST2.Fields("MABLN")
      Else
            NAGHDP = 0
      End If
   End If
   Set RST2 = New ADODB.Recordset
   RST2.Open "SELECT COUNT(NUMBER) AS CountOfNUMBER, SUM(M_NAGHD) AS SumOfM_NAGHD, SUM(MBAA) AS SumOfMBAA, SUM(MABL_VAR) AS SumOfMABL_VAR, SUM(MABL_HAV) AS SumOfMABL_HAV, SUM(TAKHFIF) AS SumOfTAKHFIF, SUM(MABL_HAZ) AS MABLHAZ FROM dbo.HEAD_LST WHERE (((HEAD_LST.TAG)=4 or tag = 22 or tag = 25) AND ((HEAD_LST.DATE_N) Between " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " And " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ")   AND (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "') AND (USER_NAME LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![USERR] & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
      If Not IsNull(RST2.Fields("SumOfM_NAGHD")) Then
            NAGHDBRP = RST2.Fields("SumOfM_NAGHD")
      Else
            NAGHDBRP = 0
      End If
   End If
   Set RST2 = New ADODB.Recordset
   rst.Open "SELECT COUNT(NUMBER) AS CountOfNUMBER, SUM(M_NAGHD) AS SumOfM_NAGHD, SUM(MBAA) AS SumOfMBAA, SUM(MABL_VAR) AS SumOfMABL_VAR, SUM(MABL_HAV) AS SumOfMABL_HAV, SUM(TAKHFIF) AS SumOfTAKHFIF, SUM(MABL_HAZ) AS MABLHAZ FROM dbo.HEAD_LST WHERE (((HEAD_LST.TAG)=13 or tag = 14) AND ((HEAD_LST.DATE_N) Between " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " And " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ")   AND (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "') AND (USER_NAME LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![USERR] & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   RST2.Open "SELECT SUM(dbo.INVO_LST.MABL_K) AS SMABL_K, SUM(dbo.INVO_LST.MABL_K + dbo.INVO_LST.IMBAA - dbo.INVO_LST.N_MOIN) AS FROOSHKH FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG -11  = dbo.INVO_LST.TAG WHERE (((HEAD_LST.TAG)=13) AND ((HEAD_LST.DATE_N) Between " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " And " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ")   AND (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "') AND (USER_NAME LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![USERR] & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 1 Then
      HKH = rst.Fields("MABLHAZ")
      TF = rst.Fields("SumOfTAKHFIF")
      MBAA = rst.Fields("SumOfMBAA")
      GABEL = RST2.Fields("FROOSHKH") + rst.Fields("MABLHAZ")
      NAGHD = IIf(IsNull(rst.Fields("SumOfM_NAGHD")), 0 - NAGHDBRP, rst.Fields("SumOfM_NAGHD") - NAGHDBRP)
      NAGHDDP = IIf(IsNull(rst.Fields("SumOfM_NAGHD")), 0, rst.Fields("SumOfM_NAGHD")) + NAGHDD - NAGHDP - NAGHDBRP
      VARIZ = IIf(IsNull(rst.Fields("SumOfMABL_VAR")), 0, rst.Fields("SumOfMABL_VAR"))
      HAVL = IIf(IsNull(rst.Fields("SumOfMABL_HAV")), 0, rst.Fields("SumOfMABL_HAV"))

      Me.LNAGHDD.CAPTION = Format(NAGHDD, "#,##")
      Me.LNAGHDP.CAPTION = Format(NAGHDP, "#,##")
      Me.LNAGHDDP.CAPTION = Format(NAGHDDP, "#,##")
      Me.MBAA.CAPTION = Format(MBAA, "#,##")
      Me.HKH.CAPTION = Format(HKH, "#,##")
      Me.TF.CAPTION = Format(TF, "#,##")
      Me.GABEL.CAPTION = Format(GABEL, "#,##")
      Me.NAGHD.CAPTION = Format(NAGHD, "#,##")
      Me.VARIZ.CAPTION = Format(VARIZ, "#,##")
      Me.HAVAL.CAPTION = Format(HAVL, "#,##")
      rst.Close
      Set rst = New ADODB.Recordset
      rst.Open "SELECT Sum(IsNull(PAY_GETD.MABL,0)) AS mab, Count(PAY_GETD.MABL) AS CountOfMABL FROM HEAD_LST LEFT JOIN PAY_GETD ON (HEAD_LST.TAG = PAY_GETD.TAG) AND (HEAD_LST.NUMBER = PAY_GETD.NUMBER) WHERE (((HEAD_LST.TAG)=2) AND ((HEAD_LST.DATE_N) Between " & Forms![F_MENU_GOZARESH_FROOSH]![DT1] & " And " & Forms![F_MENU_GOZARESH_FROOSH]![DT2] & ")   AND (DEPATMAN = " & Forms![F_MENU_GOZARESH_FROOSH]![DEPART] & ") AND (CAST(SHIFT AS NVARCHAR) LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![SSHIFT] & "') AND (USER_NAME LIKE N'" & Forms![F_MENU_GOZARESH_FROOSH]![USERR] & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 1 Then
        Me.CHK.CAPTION = Format(IIf(IsNull(rst.Fields("MAB")), 0, rst.Fields("MAB")), "#,##")
        CHK = IIf(IsNull(rst.Fields("MAB")), 0, rst.Fields("MAB"))
      End If
      Me.JPAY.CAPTION = Format(NAGHD + VARIZ + HAVL + CHK, "#,##")
      Me.PAR.CAPTION = Format(NAGHD + VARIZ + HAVL + CHK, "#,##")
      Me.MAN.CAPTION = Format(GABEL - (NAGHD + VARIZ + HAVL + CHK), "#,##")
   Else
   End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_GOZARESH_FROOSH"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  If [Forms]![F_MENU_GOZARESH_FROOSH]![SSHIFT] = "%" Then
    Me.SHIFT.Visible = False
  End If
  If [Forms]![F_MENU_GOZARESH_FROOSH]![USERR] = "%" Then
    Me.USS.Visible = False
  End If
End Sub
