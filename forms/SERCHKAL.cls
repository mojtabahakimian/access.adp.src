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
  Select Case Me.OpenArgs
   Case "anbargardani-baz":
       Forms![anbargardani-baz].Form![CODE] = Me.CODE
    Case "HEAD_SERCH_MAIN":
       Forms![HEAD_SERCH_MAIN].Form![CODE] = Me.CODE
    Case "HEAD_LST_REQUEST":
       Forms![HEAD_LST_REQUEST]![INVO_LST_REQUEST_SUB].Form![CODE] = Me.CODE
       Forms![HEAD_LST_REQUEST]![INVO_LST_REQUEST_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_REQUEST]![INVO_LST_REQUEST_SUB].Form![VAHED_K].Requery
       Forms![HEAD_LST_REQUEST]![INVO_LST_REQUEST_SUB].Form![VAHED_K] = Me.vahed
    Case "lastprice":
       Forms![LASTPRICE].Form![CODE] = Me.CODE
       Forms![LASTPRICE].Form![Combo13] = Me.CODE
    Case "F_MENU_ANBAR":
       Forms![F_MENU_ANBAR].Form![KALA] = Me.CODE
    Case "F_MENU_KART":
       Forms![F_MENU_KART].Form![KALA] = Me.CODE
    Case "F_MENU_ANBAR_FRKH":
       Forms![F_MENU_ANBAR_FRKH].Form![KALA] = Me.CODE
    Case "ORDR_HED":
       Forms![ORDR_HED]![ORDER_LST_SUB].Form![CODE] = Me.CODE
       Forms![ORDR_HED]![ORDER_LST_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![ORDR_HED]![ORDER_LST_SUB].Form![VAHED_K].Requery
       Forms![ORDR_HED]![ORDER_LST_SUB].Form![VAHED_K] = Me.vahed
       Forms![ORDR_HED]![ORDER_LST_SUB].Form![CUST_NO] = Me.TOZIH
        rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
        End If

    Case "HEAD_LST_FROOSH":
       Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![CODE] = Me.CODE
       Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![mabl] = Me.MABL_F
       Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![VAHED_K].Requery
       Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![VAHED_K] = Me.vahed
       Forms![HEAD_LST_FROOSH]![MOGU] = Me.MOG

    Case "HEAD_LST_PISHFROOSH":
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![CODE] = Me.CODE
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![mabl] = Me.MABL_F
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K].Requery
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K] = Me.vahed
       Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB]![MOGU] = Me.MOG

    Case "HEAD_LST_PISHFROOSH2":
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![CODE] = Me.CODE
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![mabl] = Me.MABL_F
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K].Requery
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![VAHED_K] = Me.vahed
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![MABL_K] = Me.MABL_F * Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![MEGHk]
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![N_MOIN] = Round(Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![N_KOL] * Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![MABL_K] / 100) + Round((Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![MABL_K] - Round(Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![N_KOL] * Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![MABL_K] / 100)) * Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![TKHN] / 100)
       Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB]![MOGU] = Me.MOG

    Case "HEAD_LST_KHAREED":
       Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![CODE] = Me.CODE
       Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![mabl] = Me.MABL_F
       Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![VAHED_K].Requery
       Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![VAHED_K] = Me.vahed
       Forms![head_lst_khareed]![MOGU] = Me.MOG

    Case "HEAD_LST_KHADAMAT":
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![CODE] = Me.CODE
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![mabl] = Me.MABL_F
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K].Requery
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K] = Me.vahed

    Case "HAVALAH_EXIT":
       Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![CODE] = Me.CODE
       Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![VAHED_K].Requery
       Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![VAHED_K] = Me.vahed
       Forms![HAVALAH_EXIT]![MOGU] = Me.MOG

    Case "HAVALAH_ENTER":
       Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![CODE] = Me.CODE
       Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![VAHED_K].Requery
       Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![VAHED_K] = Me.vahed
       Forms![havalah_enter]![MOGU] = Me.MOG

    Case "HEAD_LST_ENTEGHAL":
       Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![CODE] = Me.CODE
       Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![VAHED_K].Requery
       Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![VAHED_K] = Me.vahed
       Forms![HEAD_LST_ENTEGHAL]![MOGU] = Me.MOG

    Case "HEAD_LST_HAVL":
       Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![CODE] = Me.CODE
       Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![VAHED_K].Requery
       Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![VAHED_K] = Me.vahed
       Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![MEGHk] = Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![MEGH]

    Case "HEAD_MANF":
       Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![CODE] = Me.CODE
       Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![VAHED_K].Requery
       Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![VAHED_K] = Me.vahed

    Case "HEAD_LST_RASID":
       Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![CODE] = Me.CODE
       Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![VAHED_K].Requery
       Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![VAHED_K] = Me.vahed

    Case "HEAD_LST_MOSTAGHIM":
       Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![CODE] = Me.CODE
       Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![VAHED_K].Requery
       Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![VAHED_K] = Me.vahed

    Case "HAVALAH_EXIT_SAYER":
       Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![CODE] = Me.CODE
       Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![VAHED_K].Requery
       Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![VAHED_K] = Me.vahed
       Forms![havalah_exit_sayer]![MOGU] = Me.MOG
     Case "HEAD_LST_RASID_OTHER":
       Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![CODE] = Me.CODE
       Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![VAHED_K].Requery
       Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![VAHED_K] = Me.vahed
    Case "HEAD_LST_FROOSH22":
       Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Form![CODE] = Me.CODE
       Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Form![VAHED_K].Requery
       Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Form![VAHED_K] = Me.vahed
       Forms![HEAD_LST_FROOSH22]![MOGU] = Me.MOG
     Case "HEAD_LST_hav_OTHER":
       Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![CODE] = Me.CODE
       Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![VAHED_K].Requery
       Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![VAHED_K] = Me.vahed
      Case "ENHESAR_KALA_FORM":
       Forms![ENHESAR_KALA_FORM].Form![EN_KALA_COD] = Me.CODE
    Case "PM_ASSETS_FR":
       Forms![PM_ASSETS_FR]![PM_RUMAINTE_FR]![PM_MVT_FR].Form![MV_CODE] = Me.CODE
       Forms![PM_ASSETS_FR]![PM_RUMAINTE_FR]![PM_MVT_FR].Form![MV_vahed_k].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![PM_ASSETS_FR]![PM_RUMAINTE_FR]![PM_MVT_FR].Form![MV_vahed_k].Requery
       Forms![PM_ASSETS_FR]![PM_RUMAINTE_FR]![PM_MVT_FR].Form![MV_vahed_k] = Me.vahed
  End Select
    DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub

Private Function ANBARfunc() As Integer
    If Me.OpenArgs = "F_MENU_ANBAR_FRKH" Then
           ANBARfunc = Forms![F_MENU_ANBAR_FRKH].Form![ANBAR]
    End If
    If Me.OpenArgs = "F_MENU_KART" Then
           ANBARfunc = Forms![F_MENU_KART].Form![ANBAR]
    End If
    If Me.OpenArgs = "F_MENU_ANBAR" Then
           ANBARfunc = Forms![F_MENU_ANBAR].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_MOSTAGHIM" Then
           ANBARfunc = Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_RASID" Then
           ANBARfunc = Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_HAVL" Then
           ANBARfunc = Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "OHEAD_LST_FROOSH" Then
           ANBARfunc = Forms![OHEAD_LST_FROOSH]![OINVOLST_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_PISHFROOSH" Then
           ANBARfunc = Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_PISHFROOSH2" Then
           ANBARfunc = Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_FROOSH" Then
           ANBARfunc = Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_KHADAMAT" Then
           ANBARfunc = Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_KHAREED" Then
           ANBARfunc = Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HAVALAH_EXIT" Then
           ANBARfunc = Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HAVALAH_ENTER" Then
           ANBARfunc = Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_ENTEGHAL" Then
           ANBARfunc = Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_MANF" Then
           ANBARfunc = Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![ANBAR]
    End If
    If Me.OpenArgs = "HAVALAH_EXIT_SAYER" Then
           ANBARfunc = Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_RASID_OTHER" Then
           ANBARfunc = Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_FROOSH22" Then
           ANBARfunc = Forms![HEAD_LST_FROOSH22]![INVO_LST_sub].Form![ANBAR]
    End If
    If Me.OpenArgs = "HEAD_LST_hav_other" Then
           ANBARfunc = Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![ANBAR]
    End If
        Dim rst As New ADODB.Recordset
    If Me.OpenArgs = "ORDR_HED" Then
        rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           ANBARfunc = rst.Fields(0)
        Else
           ANBARfunc = 1
        End If
    End If

    If Me.OpenArgs = "PM_ASSETS_FR" Then
        rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           ANBARfunc = rst.Fields(0)
        Else
           ANBARfunc = 1
        End If
    End If



End Function

Private Sub Form_Load()
             On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 61, 1) = "5" Then
        Dim anb As Integer
        anb = ANBARfunc()
        Me.RecordSource = "SELECT     mogudi_1_1.CODE, ROUND(ISNULL(ISNULL(mogudi_1_1.SMEGH, 0) - ISNULL(mogudi_2_2.SMEGH, 0), 0), 2) AS mog, ISNULL(mogudi_2_2.SMEGH, 0) AS MEGF, " & _
                       " ISNULL(mogudi_1_1.SMEGH, 0) AS megk, STUF_DEF.NAME + N' ' + ISNULL(STUF_DEF.TOZIH, N' ') + N' ' + ISNULL(STUF_DEF.N_FANI, N' ') as name, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_VAHEDS.CODE AS VAHED, ' ' AS KK, mogudi_1_1.ANBAR, CAST(mogudi_1_1.CODE AS BIGINT) AS VCOD," & _
                       " dbo.STUF_DEF.N_FANI, dbo.STUF_DEF.TOZIH, ISNULL(dbo.FNESBAT.FNESBAT, 1) AS NESBAT, ROUND(ISNULL(ISNULL(mogudi_1_1.SMEGH, 0)" & _
                       " - ISNULL(mogudi_2_2.SMEGH, 0), 0) / ISNULL(dbo.FNESBAT.FNESBAT, 1), 0) AS MANDF, dbo.STUF_DEF.B_SEF, dbo.STUF_DEF.MABL_F, dbo.STUF_DEF.MAX_M," & _
                       " dbo.STUF_DEF.VAZN, ROUND(ISNULL(ISNULL(mogudi_1_1.SMEGH, 0) - ISNULL(mogudi_2_2.SMEGH, 0), 0), 2) * dbo.STUF_DEF.VAZN AS VAZNK" & _
                       " FROM         dbo.TCOD_VAHEDS INNER JOIN  dbo.STUF_DEF INNER JOIN    dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE ON dbo.TCOD_VAHEDS.CODE = dbo.STUF_DEF.VAHED LEFT OUTER JOIN  dbo.mogudi_2_2(" & anb & ") mogudi_2_2 ON dbo.STUF_FSK.ANBAR = mogudi_2_2.ANBAR AND   dbo.STUF_FSK.CODE = mogudi_2_2.CODE LEFT OUTER JOIN   dbo.mogudi_1_1(" & anb & ") mogudi_1_1 ON dbo.STUF_FSK.ANBAR = mogudi_1_1.ANBAR AND   dbo.STUF_FSK.CODE = mogudi_1_1.CODE LEFT OUTER JOIN" & _
                       " dbo.FNESBAT ON dbo.STUF_DEF.CODE = dbo.FNESBAT.CODE  WHERE     (NOT (mogudi_1_1.CODE IS NULL))"
    End If
    If Me.OpenArgs = "visit_gol_form" Then
        Me.RecordSource = "search_kalv"
    End If
    If Me.OpenArgs = "VISITORS_PORSANT_HEAD" Then
        Me.RecordSource = "search_kalM"
    End If

End Sub
