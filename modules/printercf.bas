Attribute VB_Name = "printercf"
Option Compare Database
Option Explicit
 Type str_DEVMODE
      RGB As String * 94
  End Type
 Type str_PRTMIP
      strRGB As String * 28
  End Type
  Type type_PRTMIP
      intLeftMargin As Integer
      intTopMargin As Integer
      intRightMargin As Integer
      intBotMargin As Integer
      intDataOnly As Integer
      intWidth As Integer
      intHeight As Integer
      intDefaultSize As Integer
      intColumns As Integer
      intColumnSpacing As Integer
      intRowSpacing As Integer
      intItemLayout As Integer
      intFastPrint As Integer
      intDatasheet As Integer
  End Type
  Type type_DEVMODE
      strDeviceName As String * 16
      intSpecVersion As Integer
      intDriverVersion As Integer
      intSize As Integer
      intDriverExtra As Integer
      lngFields As Long
      intOrientation As Integer
      intPaperSize As Integer
      intPaperLength As Integer
      intPaperWidth As Integer
      intScale As Integer
      intCopies As Integer
      intDefaultSource As Integer
      intPrintQuality As Integer
      intColor As Integer

      intDuplex As Integer
      intResolution As Integer
      intTTOption As Integer
      intCollate As Integer
      strFormName As String * 16
      lngPad As Long
      lngBits As Long
      lngPW As Long
      lngPH As Long
      lngDFI As Long
      lngDFr As Long
  End Type

  Public Function CheckCustomPage(rptname As String, plen As Double)
      Dim DevString As str_DEVMODE
      Dim dm As type_DEVMODE
      Dim strDevModeExtra As String
      Dim PrtMipString As str_PRTMIP
      Dim PM As type_PRTMIP
      Dim rpt As Report
      Dim intResponse As Integer
      ' Opens report in Design view.
      DoCmd.OpenReport rptname, acDesign

      Set rpt = Reports(rptname)
      If Not IsNull(rpt.PrtDevMode) Then
          strDevModeExtra = rpt.PrtDevMode     ' Gets current DEVMODE structure.
          DevString.RGB = strDevModeExtra
          LSet dm = DevString
   '       If DM.intPaperSize = 256 Then
   '           intResponse = MsgBox("The current custom page size is " _
   '               & DM.intPaperWidth / 254 & " inches wide by " _
   '               & DM.intPaperLength / 254 & " inches long. Do you want " _
   '               & "to change the settings?", 4)
   '       Else
   '           ' Currently not user-defined.
   '           intResponse = MsgBox("The report does not have a custom page size. " _
   '               & "Do you want to define one?", 4)
   '       End If
   '       If intResponse = 6 Then
              ' User wants to change settings.
              ' Initialize fields.
              dm.lngFields = dm.lngFields Or dm.intPaperSize Or dm.intPaperLength _
                  Or dm.intPaperWidth
              dm.intPaperSize = 256           ' Set custom page.
              ' Prompt for length and width.

              dm.intPaperLength = plen * 100
              dm.intPaperWidth = 2200

              LSet DevString = dm             ' Update property.
              Mid(strDevModeExtra, 1, 94) = DevString.RGB
              rpt.PrtDevMode = strDevModeExtra
   '       End If set margins
            PrtMipString.strRGB = rpt.PrtMip
            LSet PM = PrtMipString
            PM.intLeftMargin = 0     ' Set margins.
            PM.intTopMargin = 0
            PM.intRightMargin = 0
            PM.intBotMargin = 0
            LSet PrtMipString = PM              ' Update property.
            rpt.PrtMip = PrtMipString.strRGB
      End If
      DoCmd.Close acReport, rptname, acSaveYes
  End Function


 Sub SwitchOrient(strName As String)
      Const DM_PORTRAIT = 1
      Const DM_LANDSCAPE = 2
      Dim DevString As str_DEVMODE
      Dim dm As type_DEVMODE
      Dim strDevModeExtra As String
      Dim rpt As Report
      DoCmd.OpenReport strName, acDesign              ' Opens report in Design view.
      Set rpt = Reports(strName)
      If Not IsNull(rpt.PrtDevMode) Then
          strDevModeExtra = rpt.PrtDevMode
          DevString.RGB = strDevModeExtra
          LSet dm = DevString
          dm.lngFields = dm.lngFields Or dm.intOrientation    ' Initialize fields.

          If dm.intOrientation = DM_PORTRAIT Then
              dm.intOrientation = DM_LANDSCAPE
          Else
              dm.intOrientation = DM_PORTRAIT
          End If
          LSet DevString = dm                         ' Update property.
          Mid(strDevModeExtra, 1, 94) = DevString.RGB
          rpt.PrtDevMode = strDevModeExtra
      End If
  End Sub

  Sub PrtMipCols(strName As String)

      Dim PrtMipString As str_PRTMIP
      Dim PM As type_PRTMIP
      Dim rpt As Report
      Const PM_HORIZONTALCOLS = 1953
      Const PM_VERTICALCOLS = 1954
      DoCmd.OpenReport strName, acDesign
      Set rpt = Reports(strName)
      PrtMipString.strRGB = rpt.PrtMip
      LSet PM = PrtMipString
      PM.intColumns = 2                       ' Create two columns.
      PM.intRowSpacing = 0.25 * 1440      ' Set 0.25 inch between rows.
      PM.intColumnSpacing = 0.5 * 1440    ' Set 0.5 inch between columns.
      PM.intItemLayout = PM_HORIZONTALCOLS


      LSet PrtMipString = PM              ' Update property.
      rpt.PrtMip = PrtMipString.strRGB
  End Sub

 Sub SetMarginsToDefault(strName As String)
      Dim PrtMipString As str_PRTMIP
      Dim PM As type_PRTMIP
      Dim rpt As Report
      DoCmd.OpenReport strName, acDesign
      Set rpt = Reports(strName)
      PrtMipString.strRGB = rpt.PrtMip
      LSet PM = PrtMipString
      PM.intLeftMargin = 1 * 1440     ' Set margins.
      PM.intTopMargin = 1 * 1440
      PM.intRightMargin = 1 * 1440
      PM.intBotMargin = 1 * 1440
      LSet PrtMipString = PM              ' Update property.
      rpt.PrtMip = PrtMipString.strRGB

 End Sub
 Public Function SETDEFPRN(USER As String)
                  On Error Resume Next
    Dim infileopen As Boolean
    Dim incmd$, batch$, Q$
    Open "c:\PRNCONF.CFG" For Input As #1
    If Err = 0 Then
        infileopen = True
        Do While Not EOF(1)
            Line Input #1, incmd$
            If incmd$ = USER Then
                Line Input #1, incmd$
                Application.Printer = Application.Printers(incmd$)
            End If
       Loop
    End If
    Err.Clear
        Close #1
End Function
Public Function printcusti1()
                   On Error Resume Next
  Dim rptname As String, STTIME, rst As New ADODB.Recordset, JAMF As Long
   Dim prt As Printer, RST2 As New ADODB.Recordset, NOTPR As Boolean
  If IsLoaded("HEAD_LST_FROOSH2") Then
    If Forms![HEAD_LST_FROOSH2].NUMBER > 0 Then
        If MID(Forms![BASEKNOW]![OPTIONSS], 4, 1) = "5" Then
              rptname = "I" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 5, 2))
        Else
               rptname = "I1"
        End If
        Err.Clear
        If rptname = "I6" Then
         '   DoCmd.OpenReport rptname, acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = 2"
            For Each prt In Application.Printers
                If Left(prt.DeviceName, 10) = "BTP-2002NP" Then
                   Application.Printer = prt
                End If
            Next
           rst.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Forms![HEAD_LST_FROOSH2].NUMBER & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                JAMF = rst.Fields(0)
           End If
           Set rst = New ADODB.Recordset
           rst.Open "select * from head_lst where tag = 2 and NUMBER =  " & Forms![HEAD_LST_FROOSH2].NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
               rst.Fields("m_naghd") = JAMF + rst.Fields("MABL_HAZ") + rst.Fields("MBAA") - rst.Fields("TAKHFIF")
               JAMF = JAMF + rst.Fields("MABL_HAZ") + rst.Fields("MBAA") - rst.Fields("TAKHFIF")
               rst.update
               rst.Close
           End If
           If Forms![HEAD_LST_FROOSH2]![M_NAGHD] <> JAMF Then
              Forms![HEAD_LST_FROOSH2]![M_NAGHD] = JAMF
           End If
           Forms![HEAD_LST_FROOSH2]![CUST_NO].SetFocus
           DoCmd.OpenReport rptname, acViewNormal, , "NUMBER =" & Forms![HEAD_LST_FROOSH2].NUMBER & " AND TAG =2"
             If Forms![BASEKNOW]![LOCKFAP] Then
               Forms![HEAD_LST_FROOSH2]![OKF] = True
             End If
             If Forms![HEAD_LST_FROOSH2]![OKF] Then
                  Forms![HEAD_LST_FROOSH2].AllowDeletions = False
                  Forms![HEAD_LST_FROOSH2].AllowEdits = False
                  Forms![HEAD_LST_FROOSH2]![INVO_LST_sub].Enabled = False
                  Forms![HEAD_LST_FROOSH2]![Page58].Enabled = False
                  Forms![HEAD_LST_FROOSH2]![ESLAH].Enabled = True
             End If
           If Err = 0 Then
             Call CTRLINS
           End If
        End If
    End If
Else
 If IsLoaded("head_lst_froosh22") Then
    NOTPR = False
    If Forms![BASEKNOW]![RMOG] Then
          RST2.Open "select * from invo_lst where number  = " & Forms![HEAD_LST_FROOSH22].NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          While Not RST2.EOF
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                   If rst.Fields("MAND") < 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & "از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست"
                      NOTPR = True
                   End If
                End If
              '  RST.Close
                Set rst = New ADODB.Recordset
                RST2.MoveNext
          Wend
    End If
    If NOTPR = False Then
        If Forms![HEAD_LST_FROOSH22].NUMBER > 0 Then
            If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
                 rptname = "I" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
            Else
                 rptname = "I1"
            End If
            Err.Clear
            If rptname = "I1" Then
             '   DoCmd.OpenReport rptname, acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = 2"
                For Each prt In Application.Printers
                    If Left(prt.DeviceName, 17) = "BTP-2002NP(U) 1 (" Then
                       Application.Printer = prt
                    End If
                Next
            End If
            DoCmd.OpenReport rptname, acViewNormal, , "NUMBER =" & Forms![HEAD_LST_FROOSH22].NUMBER & " AND HTAG =2", , 2
              If Forms![BASEKNOW]![LOCKFAP] Then
                Forms![HEAD_LST_FROOSH22]![OKF] = True
              End If
              If Forms![HEAD_LST_FROOSH22]![OKF] Then
                   Forms![HEAD_LST_FROOSH22].AllowDeletions = False
                   Forms![HEAD_LST_FROOSH22].AllowEdits = False
                   Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Enabled = False
                   Forms![HEAD_LST_FROOSH22]![Page58].Enabled = False
                   Forms![HEAD_LST_FROOSH22]![ESLAH].Enabled = True
              End If
            If Err = 0 Then
              Call CTRLINS
            End If
        End If
     End If
  End If
  End If
  Application.Printer = Nothing
End Function


Public Sub Printlog(num As Double, tg As Integer, no As Integer)
    Dim rst As New ADODB.Recordset
    rst.Open "chapnum", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = num
    rst.Fields("TAG") = tg
    rst.Fields("NUM") = no
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
    rst.Close
End Sub

Public Function Printed(num As Double, tg As Integer, Optional no As Integer) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "select * from chapnum where number = " & num & " AND  tag = " & tg, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Printed = True
    Else
        Printed = False
    End If
End Function
