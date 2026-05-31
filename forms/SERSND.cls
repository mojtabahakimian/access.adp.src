Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim CTL As Control
Private Sub Form_Current()
                     On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
 If KeyAscii = 13 Then
 Select Case Left(Forms![sersndtaf].OpenArgs, 2)
    Case "01":
         Forms![DEED_HEAD]![Child14].Form![HES_K] = Me.N_KOL
         Forms![DEED_HEAD]![Child14].Form![HES_M] = Me.NUMBER
         Forms![DEED_HEAD]![Child14].Form![HES_T] = Me.TNUMBER
       '  Forms![DEED_HEAD]![Child14].Form![HES_T2] = Me.TNUMBER2
         Forms![DEED_HEAD]![Child14].Form![HES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![DEED_HEAD]![Child14].Form![Combo30].SetFocus
    Case "02":
         Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_K] = Me.N_KOL
         Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_M] = Me.NUMBER
         Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T] = Me.TNUMBER
         Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "03":
         Forms![PGET_HED]![PGET_LST_SUB].Form![THES_K] = Me.N_KOL
         Forms![PGET_HED]![PGET_LST_SUB].Form![THES_M] = Me.NUMBER
         Forms![PGET_HED]![PGET_LST_SUB].Form![THES_T] = Me.TNUMBER
         Forms![PGET_HED]![PGET_LST_SUB].Form![THES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "04":
        Forms![HEAD_LST_MOSTAGHIM]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "05":
         Forms![F_MENU_KOL_MOIN_TAFZIL]![Combo34] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "06":
        Forms![HEAD_LST_FROOSH1]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "07":
        Forms![head_lst_khareed1]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "08":
        Forms![HEAD_LST_KHADAMAT]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "09":
         Forms![F_MENU_KHFR]![HMOIN] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![F_MENU_KHFR]![HHMOIN] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "00":
         Forms![F_MENU_MOGHAYERAT]![KOL] = Me.N_KOL
         Forms![F_MENU_MOGHAYERAT]![MOIN] = Me.NUMBER
         Forms![F_MENU_MOGHAYERAT]![taf] = Me.TNUMBER
         Forms![F_MENU_MOGHAYERAT]![XTAF] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "10":
         Forms![HEAD_LST_MOSTAGHIM]![MOIN_HAZ] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HEAD_LST_MOSTAGHIM]![MOIN_HAZ1] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "11":
         Forms![HAVALAH_EXIT]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HAVALAH_EXIT]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "12":
         Forms![havalah_enter]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![havalah_enter]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "13":
         Forms![head_lst_havl]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![head_lst_havl]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "14":
         Forms![OHEAD_LST_KHADAMAT]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![OHEAD_LST_KHADAMAT]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "15":
         Forms![havalah_exit_sayer]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![havalah_exit_sayer]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "16":
         Forms![HEAD_LST_PISHFROOSH2]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HEAD_LST_PISHFROOSH2]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
           If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
          rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND] = rst.Fields("CUST_COD")
          End If
        End If
  Case "17":
         Forms![head_lst_pishfroosh]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![head_lst_pishfroosh]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "18":
         Forms![PTAMIR]![POWNER] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "19":
         Forms![PTAMIR]![TAMIRKAR1] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "20":
         Forms![PTAMIR]![TAMIRKAR2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "21":
         Forms![ORDR_HED]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "  -  " & Me.nam, Me.TNUMBER2 & "  -  " & Me.tname)
         Forms![ORDR_HED]![CUST_NO].SetFocus
    Case "22":
         Forms![HEAD_LST_FROOSH22]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HEAD_LST_FROOSH22]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
          rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              Forms![HEAD_LST_FROOSH22]![CUST_KIND] = rst.Fields("CUST_COD")
          End If
        End If
    Case "24":
         Forms![CREATE_CHEKDP]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![CREATE_CHEKDP]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "25":
         Forms![CREATE_CHEKPDP]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![CREATE_CHEKPDP]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "27":
         Forms![HEAD_LST_RASID_OTHER]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HEAD_LST_RASID_OTHER]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "28":
         Forms![AZAE]![HES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![AZAE]![HES2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "29":
         Forms![head_lst_rasid]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![head_lst_rasid]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "30":
         Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB]![N_RASID] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "31":
         Forms![f_menu_date_hes]![HES1] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![f_menu_date_hes]![HES11] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "32":
         Forms![f_menu_date_hes]![HES2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![f_menu_date_hes]![HES22] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "33":
         Forms![HEAD_LST_FROOSH2]![VISITOR_DTL_SUB]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "34":
         Forms![KAR_HEAD]![KAR_BARG_SUB]![DEPART] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "35":
         Forms![Visit_route_FORM]![HES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![Visit_route_FORM]![HES2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "36":
         Forms![visitory_day_head]![VISITORS_DAY_DTL_SUB]![COUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![visitory_day_head]![VISITORS_DAY_DTL_SUB]![COUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "37":
         Forms![Users]![HES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![Users]![HES2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "38":
         Forms![head_lst_brfr]![VISITOR_DTL_SUB]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "39":
         Forms![F_MENU_KOL_MOIN_DATE_AI]![Combo34] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![F_MENU_KOL_MOIN_DATE_AI]![Combo36] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "40":
         Forms![F_MENU_KOL_MOIN_DATE_AI]![Combo39] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![F_MENU_KOL_MOIN_DATE_AI]![Combo41] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "41":
         Forms![HEAD_LST_FROOSH22]![VISITOR_DTL_SUB]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "38":
         Forms![head_lst_brfr]![VISITOR_DTL_SUB]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "42":
         Forms![HEAD_LST_HAV_other]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![HEAD_LST_HAV_other]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "43":
         Forms![Visit_route_FORM]![Visit_route_dtl_FORM]![COUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![Visit_route_FORM]![Visit_route_dtl_FORM]![COUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "44":
         Forms![PAYMENTFORMORDER]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![PAYMENTFORMORDER]![CUST_NO2] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "45":
         Forms![PAYMENTFORMORDER]![ORDERER] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![PAYMENTFORMORDER]![ORDERERID] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "46":
         Forms![main]![TASKS]![COMP_COD] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
     Case "47":
         Forms![PARVAD]![COMP_COD] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "48":
         Forms![ENHESAR1]![EN_CUST_CO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
         Forms![ENHESAR1]![EN_CUST_CO22] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
       '  If Mid(Forms![baseknow]![OPTIONSS], 19, 1) = 5 Then
          rst.Open "SELECT     hes, SHAHRID FROM dbo.CUST_HESAB WHERE     (hes = N'" & IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 And Not IsNull(rst.Fields("SHAHRID")) Then
              Forms![ENHESAR1]![EN_CITY] = rst.Fields("SHAHRID")
          End If
       ' End If
    Case "49":
         Forms![ZASESABBEESAB]![Combo34] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "50":
         Forms![ZASESABBEESAB]![TOHES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "51":
         Forms![F_USER_PERMITION FORMS]![USER_ENDN_SUB].Form![HES] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
    Case "52":
         Forms![head_lst_froosh_back2]![VISITOR_DTL_SUB]![CUST_NO] = IIf(IsNull(Me.TNUMBER2), Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER, Me.TNUMBER2)
End Select
 DoCmd.Close acForm, "sersndtaf"
 End If
 If KeyAscii = 27 Then
        DoCmd.Close acForm, "sersndtaf"
 End If
End Sub
Private Sub Form_Load()
                    On Error Resume Next
 If Forms![BASEKNOW]![PERSON] = 2 Then
    If Len(Forms![sersndtaf].OpenArgs) > 2 Then
        Me.RecordSource = "SELECT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME, CODE_E FROM cust_hesab_dtl WHERE     (" & RIGHT(Forms![sersndtaf].OpenArgs, Len(Forms![sersndtaf].OpenArgs) - 2) & ") and ((CODE_E = N'0') OR  (CODE_E IS NULL) )  ORDER BY NAME, TNAME "
    Else
        Me.RecordSource = "SELECT     TOP 100 PERCENT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME FROM dbo.cust_hesab_dtl WHERE     (CODE_E = N'0') OR  (CODE_E IS NULL) ORDER BY NAME, TNAME"
    End If
 Else
    If Len(Forms![sersndtaf].OpenArgs) > 2 Then
        Me.RecordSource = "SELECT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME, CODE_E FROM cust_hesab_dtl WHERE     (" & RIGHT(Forms![sersndtaf].OpenArgs, Len(Forms![sersndtaf].OpenArgs) - 2) & ")  ORDER BY NAME, TNAME "
    Else
        Me.RecordSource = "SELECT     TOP 100 PERCENT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME FROM dbo.cust_hesab_dtl ORDER BY NAME, TNAME"
    End If
 End If
Me.Requery
End Sub
