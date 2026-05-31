Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                  On Error Resume Next
  Dim incmd$, i

  Err.Clear
    i = 0
    Close #1
  '  Open "c:\EMZDPSP.txt" For Input As #1
  '  If Err = 0 Then
  '   Me.EMZ2 = ""
  '   Me.E1.Visible = False
  '   Me.E2.Visible = False
  '   Me.E3.Visible = False
  '   While Not EOF(1)
  '     Line Input #1, incmd$
  '     If I > 0 Then
  '          Me.EMZ2 = Me.EMZ2 & Chr(13) & incmd$
  '     Else
  '          Me.EMZ2 = Me.EMZ2 & incmd$
  '     End If
  '     I = I + 1
  '   Wend
  '  End If
  Call Showemza(Me, "SGN0134TX", "SGN0234TX", "SGN0334TX")
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                On Error Resume Next
   Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK As Double, JAMF As Double, HAZ, NAGHD, VAR, HAV, taf As Double
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   rst.Open "SELECT     SUM(dbo.PGET_LST.MABL) AS naghd, dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.PGET_LST.THES, dbo.CUST_HESAB.NAME , dbo.PGET_HED.MOLAH FROM dbo.PGET_LST INNER JOIN dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID AND dbo.PGET_LST.DATE = dbo.PGET_HED.DATE INNER JOIN dbo.CUST_HESAB ON dbo.PGET_LST.THES = dbo.CUST_HESAB.hes WHERE     (dbo.PGET_LST.NO_AM = 2) And (dbo.PGET_LST.ID = " & Forms![PGET_HED]![ID] & ") And (dbo.PGET_LST.N_SERI Is Null) And (dbo.PGET_LST.BANK Is Null) GROUP BY dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.PGET_LST.THES, dbo.CUST_HESAB.NAME, dbo.PGET_HED.MOLAH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

 If rst.RecordCount > 0 Then
       JAMF = rst.Fields("NAGHD")
   Else
       JAMF = 0
   End If
   JST.Open "SELECT     SUM(dbo.PGET_LST.MABL) AS CHK, dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, COUNT(dbo.PGET_LST.ID) AS TEDAD,  dbo.CUST_HESAB.NAME, dbo.PGET_HED.MOLAH FROM dbo.PGET_LST INNER JOIN dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID AND dbo.PGET_LST.DATE = dbo.PGET_HED.DATE INNER JOIN dbo.CUST_HESAB ON dbo.PGET_LST.THES = dbo.CUST_HESAB.hes WHERE     (dbo.PGET_LST.NO_AM = 2) AND ((dbo.PGET_LST.NAHVA = 2)OR (dbo.PGET_LST.NAHVA > 3)) AND (dbo.PGET_LST.ID = " & Forms![PGET_HED]![ID] & ")GROUP BY dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.CUST_HESAB.NAME, dbo.PGET_HED.MOLAH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 Then
        JCHK = JST.Fields("CHK")
   Else
        JCHK = 0
   End If
   Me.JF.CAPTION = Format(JCHK + JAMF, "#,##0;#,##0-")

    If JAMF <> 0 And JCHK <> 0 Then
         Me.HR.CAPTION = "مبلغ: " & Format(JAMF, "#,###") & "  " & ALPHANUM(JAMF) & " " & " ريال نـــقد        و    مبلغ: " & Format(JCHK, "#,###") & "  " & ALPHANUM(JCHK) & " ريال چـــك       طي " & JST.Fields("TEDAD") & " فقره  به شرح ذيل به " & JST.Fields("NAME") & IIf(JST.Fields("MOLAH") = "" Or IsNull(JST.Fields("MOLAH")), " ", " بابت " & JST.Fields("MOLAH")) & " پرداخت گرديد "
    Else
         If JAMF = 0 And JCHK <> 0 Then
              Me.HR.CAPTION = "مبلغ: " & Format(JCHK, "#,###") & "  " & ALPHANUM(JCHK) & " ريال چـــك      طي " & JST.Fields("TEDAD") & "  فقره به شرح ذيل به " & JST.Fields("NAME") & IIf(JST.Fields("MOLAH") = "" Or IsNull(JST.Fields("MOLAH") = ""), " ", " بابت " & JST.Fields("MOLAH")) & " پرداخت گرديد "
         Else
             If JAMF <> 0 And JCHK = 0 Then
                  Me.HR.CAPTION = "مبلغ: " & Format(JAMF, "#,###") & "  " & ALPHANUM(JAMF) & " " & " ريال نـــقد         به " & rst.Fields("NAME") & IIf(rst.Fields("MOLAH") = "" Or IsNull(rst.Fields("MOLAH")), " ", " بابت " & rst.Fields("MOLAH")) & " پرداخت گرديد "
                  Me.Detail.Visible = False
                  Me.Label139.Visible = False
                  Me.Label140.Visible = False
                  Me.Label141.Visible = False
                  Me.Label142.Visible = False
                  Me.Label143.Visible = False
                  Me.Label144.Visible = False
                  Me.Label146.Visible = False
                  Me.Label197.Visible = False
                  Me.hr2.Visible = False
                  Me.JF.Visible = False
                  Me.Line137.Visible = False
                  Me.Line136.Visible = False
                  Me.Line135.Visible = False
                  Me.Line134.Visible = False
                  Me.Line133.Visible = False
                  Me.Line131.Visible = False
                  Me.Line204.Visible = False
                  Me.Box201.Visible = False
                  Me.Box196.Height = 0
            End If
         End If
    End If
   Me.hr2.CAPTION = ALPHANUM(JCHK + JAMF) & " " & "ريال "
   rst.Close
   JST.Close
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
