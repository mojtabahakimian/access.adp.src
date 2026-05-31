Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim can As Boolean

Private Sub ARSESH_AfterUpdate()
     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تغيير اين آيتم تاثير زيادي روي آيتمهاي محاسبه ماليات ميگذارد حتما دقت كنيد كه عمليات درستي انجام مي دهيد."
End Sub
Private Sub Command147_Click()
                     On Error Resume Next
    can = True
    DoCmd.Close acForm, Me.NAME, acSaveNo
End Sub

Private Sub Command309_Click()
                   On Error Resume Next
  DoCmd.OpenForm "TCOD_ZARMALIAT", acFormDS, , , , acDialog
End Sub

Private Sub Command314_Click()
   DoCmd.OpenForm "SAZMAN_2", , , , , acDialog
End Sub

Private Sub Command322_Click()
    DoCmd.OpenForm "OPTIONS", , , , , acDialog
End Sub

Private Sub Command351_Click()
                 On Error Resume Next
    Dim dt As Date
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
       DoCmd.RunSQL ("INSERT INTO dbo.TR_SAZMAN (UNIVERSITY_CO, NAME, CITY, MANAGER, MOAVEN, ZIHESAB, AMINAMVAL, YEA, SANAD, GHAYM, KALA, PERSON, DIG, WAR, LST, TFTPAGE, TFSAZMAN,TFADDRESS, TFTEL, TFCODE_E, WIDTH_D, HIGH_D, CPI, SANDOGH, BANKHA, BESTANKAR, BEDEHKAR, KHARID, MKHARID, TKHARID, HKHARID, FROSH, MFROSH,TFROSH, HFROSH, MOGODIA, MOGODIP, DARAM, HDARAM, HKOL, ADA, APA, ADV, HAVALAH, CTRL_TS, F_ANBARF, GH_PK, L_NUMBER, SF_G, TAR_KM, BACKPATH,TKHF, HAZ_TOL, PJHAZ_TOL1, PHAZ_TOL, GHEYMAT, PPDAST, PPSAR, AMALKARD, PERSONEL, PERVAM, CONKAL, EMZA, HNAH, HEZA, HPAD, HOLA, HKHA, HJAZ,HRAN, HSAY, HCON, HSHI, HAZEDAR, EDABIM, HAZBIM, BESHO, BEDMOS, PARDAKH, HAZMALI, SAGHFH, MAND, MOJU, SA_HOGH, SA_40EZ, SA_EZAF, SA_PADA,SA_HOLA, SA_KHAR, SA_NAHA, SA_JAZB, SA_RAND, SA_COND, SA_SAYE, SA_23BI, HAZTOLID, HAZFROOSH, HAZKHADAMAT, PISHDAR, DEFANB, DEFTKH, ECONM,FRUP, UPDDATE, FINALS, PSANDHES, SANAVP, BON, ISO_FROOSH, ISO_KHAREED, ISO_MAVAD, ISO_TOLID, ISO_MAVADSAYER, SANAT, CODEVIEW, PKHARID," & _
                      " SIGN, BARCOD, SAGHF, SERVERNAM, TENDAR, LECOL1, LECOL2, LECOL3, LECOL4, LKCOL1, HESMBAA, ECODE, PCODE, IYALAT, MCODEM, HPOR, SAGHF2, OPTIONSS, CTL_DT, LOCKFAP, LOCKFSI, TRANSF, OKF, ARSESH, RMOG, APV, HOTCOD, STFR, STKH, STHFR, STHKH, STENT, STKHS, STKHH, STTOL, STFRB, STBKH, STMO, STKHA, SNDKH, UP_DATE, UP_TIME, UP_USER_NAME, PC_NAME, IPADD) SELECT     UNIVERSITY_CO, NAME, CITY, MANAGER, MOAVEN, ZIHESAB, AMINAMVAL, YEA, SANAD, GHAYM, KALA, PERSON, DIG, WAR, LST, TFTPAGE, TFSAZMAN,TFADDRESS, TFTEL, TFCODE_E, WIDTH_D, HIGH_D, CPI, SANDOGH, BANKHA, BESTANKAR, BEDEHKAR, KHARID, MKHARID, TKHARID, HKHARID, FROSH, MFROSH,TFROSH, HFROSH, MOGODIA, MOGODIP, DARAM, HDARAM, HKOL, ADA, APA, ADV, HAVALAH, CTRL_TS, F_ANBARF, GH_PK, L_NUMBER, SF_G, TAR_KM, BACKPATH,TKHF, HAZ_TOL, PJHAZ_TOL1, PHAZ_TOL, GHEYMAT, PPDAST, PPSAR, AMALKARD, PERSONEL, PERVAM, CONKAL, EMZA, HNAH, HEZA, HPAD, HOLA, HKHA, HJAZ," & _
                      " HRAN, HSAY, HCON, HSHI, HAZEDAR, EDABIM, HAZBIM, BESHO, BEDMOS, PARDAKH, HAZMALI, SAGHFH, MAND, MOJU, SA_HOGH, SA_40EZ, SA_EZAF, SA_PADA, SA_HOLA, SA_KHAR, SA_NAHA, SA_JAZB, SA_RAND, SA_COND, SA_SAYE, SA_23BI, HAZTOLID, HAZFROOSH, HAZKHADAMAT, PISHDAR, DEFANB, DEFTKH, ECONM, FRUP, UPDDATE, FINALS, PSANDHES, SANAVP, BON, ISO_FROOSH, ISO_KHAREED, ISO_MAVAD, ISO_TOLID, ISO_MAVADSAYER, SANAT, CODEVIEW, PKHARID, SIGN, BARCOD, SAGHF, SERVERNAM, TENDAR, LECOL1, LECOL2, LECOL3, LECOL4, LKCOL1, HESMBAA, ECODE, PCODE, IYALAT, MCODEM, HPOR, SAGHF2,OPTIONSS, CTL_DT, LOCKFAP, LOCKFSI, TRANSF, OKF, ARSESH, RMOG, APV, HOTCOD, STFR, STKH, STHFR, STHKH, STENT, STKHS, STKHH, STTOL, STFRB, STBKH , STMO, STKHA, SNDKH ," & FARSIDATE(Now()) & " AS Expr2 ," & CDbl(dt) & "   AS Expr1,'" & UCurrentUser() & "' AS EXPR3,'" & CurrentMachineName() & "' AS EXPR4 , '" & GETIPADD() & "' AS EXPR5 FROM         dbo.SAZMAN ")
    End If
    Me.AllowEdits = True
    Me.DAFT_ASN.Form.AllowAdditions = True
    Me.Command322.Enabled = True
    Me.Command309.Enabled = True
    Me.Command314.Enabled = True
    Me.Command367.Enabled = True
End Sub

Private Sub Command367_Click()
   DoCmd.OpenForm "SMS_FORMAT"
End Sub

Private Sub EDABIM_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.EDABIM) Then
        If ISTAF(Me.EDABIM) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
   If can Then
     CANCEL = True
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Command148_Click()
                     On Error Resume Next
    can = False
    DoCmd.RunCommand acCmdSaveRecord
    Forms![BASEKNOW].Requery
    DoCmd.Close acForm, Me.NAME, acSaveYes
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "select NUMBER,TAG FROM HEAD_LST WHERE TAG = 13 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    Me.TKHF.Enabled = False
  Else
    Me.TKHF.Enabled = True
  End If
  rst.Close
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
 If UCurrentUser() = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or UCurrentUser() = "َAdminister" Then
    Me.OPTIONSS.Visible = True
 End If
If IsNull(Me.OpenArgs) Then
      Call SETSECURITY(Me.NAME, "SYS", 3)
End If
If Me.OpenArgs = 1 Then
    Me.Page138.Visible = False
    Me.Page139.Visible = False
    Me.Page140.Visible = False
    Me.Page159.Visible = False
    Me.Page283.Visible = False
    Me.sanati.Visible = False
    Me.crm.Visible = False
    Me.sms.Visible = False
Else
    If Me.OpenArgs = 2 Then
        Me.Page138.Visible = False
        Me.Page139.Visible = False
        Me.Page140.Visible = False
        Me.Page159.Visible = False
        Me.Page283.Visible = False
        Me.sanati.Visible = False
        Me.crm.Visible = True
        Me.sms.Visible = False
        Me.SSM.Visible = False
    End If
End If

End Sub

Private Sub HAZMALI_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.HAZMALI) Then
        If ISTAF(Me.HAZMALI) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub OLEBound157_DblClick(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.RunCommand acCmdInsertObject
End Sub

Private Sub OPTIONSS_DblClick(CANCEL As Integer)
    If Me.Command322.Enabled Then
        DoCmd.OpenForm "OPTIONS", , , , , acDialog
    End If
End Sub

Private Sub PARDAKH_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.PARDAKH) Then
        If ISTAF(Me.PARDAKH) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub SA_EZAF_AfterUpdate()
                   On Error Resume Next
  If Me.SA_EZAF Then
      Me.SA_40EZ = True
  End If
End Sub

Private Sub SAGHFH_DblClick(CANCEL As Integer)
                 On Error Resume Next
  DoCmd.OpenForm "TCOD_ZARMALIAT", acFormDS, , , , acDialog
End Sub

Private Sub SAGHFH_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub Form_Close()
                 On Error Resume Next
      CurrentProject.Properties("AllowBypassKey") = False
      If Err = 2455 Then      ' Property not found.
         CurrentProject.Properties.ADD "AllowBypassKey", False
      End If
End Sub
