Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub EVENTs_DblClick(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenForm "zoom", acFormDS, , , , , "2" & Me.EVENTs
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      DoCmd.OpenForm "tasksp_EDIT", , , , , acDialog, Me.IDNUM
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Me.RecordSource = "SELECT     EVENTS.IDNUM, EVENTS.IDD, EVENTS.EVENTS, EVENTS.STDATE, EVENTS.STTIME, EVENTS.USERNAME, EVENTS.COMPANY, EVENTS.SUMTIME, EVENTS.pic, EVENTS.skid, EVENTS.num, EVENTS.tg, TASKS.GR, TASKS.PERSONEL, TASKS.TASK, TASKS.PERIORITY, TASKS.STATUS,  TASKS.STDATE AS STDATEm, TASKS.STTIME AS STTIMEm, TASKS.ENDATE, TASKS.ENTIME, TASKS.USERNAME AS USERNAMEm, TASKS.COMP_COD,TASKS.SUMTIME AS SUMTIMEm, TASKS.pic AS picm, TASKS.ss, TASKS.skid AS skidm, TASKS.num AS numm, TASKS.tg AS tgm, TASKS.CTIM, TASKS.USERCO , TASKS.SEE, TASKS.SEET FROM         TASKS INNER JOIN   EVENTS ON TASKS.IDNUM = EVENTS.IDNUM WHERE ( EVENTS.USERNAME = N'" & Me.OpenArgs & "')  ORDER BY TASKS.IDNUM OPTION (FORCE ORDER, LOOP JOIN, HASH JOIN, ORDER GROUP)"
    Me.PERSONEL.RowSource = GetUserList

End Sub

Private Sub skid_DblClick(CANCEL As Integer)
                On Error Resume Next
'0;"روي داد جاري";1;"رسيد خريد";2;"حواله فروش";3;"برگشت خريد";4;"برگشت فروش";20;"پيش فاکتور";100;"درخواست پرداخت";13;"فاکتور فروش";12;"فاکتور خريد";32;"فاکتور خريد صادرات";6;"انتقال کالا از انبار به انبار";90;"پذيرش تعميرگاه";25;"فاکتور برگشت فروش آزاد";24;"ساير رسيد انبارها";26;"حواله برگشت خريد";27;"برگشت خريد ازاد";33;"فروش صادراتي"
'اگر در فيلد مجري اصلاح کرده ونام تايپ کرده واينتر زده
    Select Case Me.skid
        Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "deed_head", , , "n_s = " & Me.num
        Case 1: If LETSGO("RASID") Then DoCmd.OpenForm "HEAD_LST_rasid", , , , , , "NUMBER = " & Me.num
        Case 2: If LETSGO("HAVL") Then DoCmd.OpenForm "HEAD_LST_HAVL", , , , , , "NUMBER = " & Me.num
        Case 3: If LETSGO("FACTBKH") Then DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.num
        Case 4: If LETSGO("FACTBFR") Then DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.num
        Case 6: If LETSGO("ESWAP") Then DoCmd.OpenForm "HEAD_LST_ENTEGHAL", , , , , , "NUMBER = " & Me.num
        Case 20: If LETSGO("PFACTFR") Then DoCmd.OpenForm "HEAD_LST_PISHFROOSH2", , , , , , "NUMBER = " & Me.num
        Case 24: If LETSGO("RBRFR") Then DoCmd.OpenForm "HEAD_LST_rasid_other", , , "NUMBER = " & Me.num
        Case 25: If LETSGO("BRFR") Then DoCmd.OpenForm "HEAD_LST_brfr", , , , , , "NUMBER = " & Me.num
        Case 26: If LETSGO("SAYERHAV") Then DoCmd.OpenForm "HEAD_LST_hav_other", , , "NUMBER = " & Me.num
        Case 27: If LETSGO("FACTBKH") Then DoCmd.OpenForm "HEAD_LST_kh_back_azad", , , , , , "NUMBER = " & Me.num
        Case 32: If LETSGO("FACTKHSA") Then DoCmd.OpenForm "HEAD_LST_KHAREED1_SADER", , , , , , "NUMBER = " & Me.num
        Case 33: If LETSGO("FACTFRSA") Then DoCmd.OpenForm "HEAD_LST_FROOSH2_sader", , , , , , "NUMBER = " & Me.num
        Case 34: If LETSGO("PGETD") Then DoCmd.OpenForm "pget_hed", , , "id = " & Me.num
        Case 35: If LETSGO("SEFARESH") Then DoCmd.OpenForm "ordr_hed", , , "id = " & Me.num
        Case 36: If LETSGO("DARKHR") Then DoCmd.OpenForm "HEAD_LST_request", , , , , , "NUMBER = " & Me.num
        Case 37: If LETSGO("MORAKH") Then DoCmd.OpenForm "pmorakh", , , "IDNUM = " & Me.num
        Case 38: If LETSGO("HSAYER") Then DoCmd.OpenForm "HAVALAH_EXIT_SAYER", , , , , , "NUMBER = " & Me.num
        Case 39: If LETSGO("HEXIT") Then DoCmd.OpenForm "HAVALAH_EXIT", , , , , , "NUMBER = " & Me.num
        Case 40: If LETSGO("PMASSETS") Then DoCmd.OpenForm "PM_ASSETS_FR", , , , , , "IDD = " & Me.num
        Case 41: If LETSGO("PMDARKH") Then DoCmd.OpenForm "PM_HEAD_LST_REQUEST_ANBAR", , , , , , "NUMBER = " & Me.num
        Case 42: If LETSGO("PMPMEM") Then DoCmd.OpenForm "PM_EM_FAILURE_EVENTS_FR", , , , , , "IDD = " & Me.num
        Case 100: If LETSGO("PAYORD") Then DoCmd.OpenForm "paymentformorder", , , , , , Me.num
        Case 101:  DoCmd.OpenForm "tasksp_EDIT", , , , , acDialog, Me.num
        Case 90: If LETSGO("PTM") Then DoCmd.OpenForm "PTAMIR", , , , , , Me.num
        Case 12: If LETSGO("FACTKH") Then DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.num
        Case 13:
                If LETSGO("FACTFRSA") Then
                    If MID(Forms![BASEKNOW]![OPTIONSS], 53, 1) = "5" Then
                        If LETSGO("FRMOST") Then
                            DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "NUMBER = " & Me.num
                        Else
                            DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal, , , , , "NUMBER = " & Me.num
                        End If
                    Else
                        If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
                            If Forms![BASEKNOW]![TKHF] < 2 Then
                                DoCmd.OpenForm "HEAD_LST_FROOSH1", acNormal, , , , , "NUMBER = " & Me.num
                            Else
                                DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal, , , , , "NUMBER = " & Me.num
                             End If
                         Else
                             DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "NUMBER = " & Me.num
                         End If

                    End If
                End If
    End Select
End Sub

Private Sub SUMTIME_DblClick(CANCEL As Integer)
                 On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If LETSGO("SMSFDTL") Then
       DoCmd.OpenForm "mesagepanel_CUSTOMER", , , , , acDialog, 1
    End If

End Sub
