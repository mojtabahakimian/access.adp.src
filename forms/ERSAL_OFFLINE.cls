Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command13_Click()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, RSTI As New ADODB.Recordset, rstP As New ADODB.Recordset, fs, a
   Dim DBPATH, DESPATH, COMM As String, zereshk As Long, ST$, ste$, pp As Integer
   Dim i As Integer, K As Integer
   K = 1
   If IsNull(Me.PATH) Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "مسير مشخص نشده است..."
   Else
        DoCmd.OpenForm "BUN"
        If Me.OpenArgs = 1 Then
             rst.Open "SELECT     dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE (TKHF = 1) AND (TAG = 13) ORDER BY NUMBER1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                If (RIGHT(PATH, 1) <> "\") Then
                   PATH = PATH & "\"
                End If
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set a = fs.CREATETextFile("c:\sendoff.txt", True)
                While Not rst.EOF()
                    Forms![BUN].[num].CAPTION = K
                    Forms![BUN].Form.Refresh
                    K = K + 1
                    RST2.Open "SELECT     dbo.invo_LST.* FROM dbo.INVO_LST WHERE (NUMBER = " & rst.Fields("NUMBER") & ") AND (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    RST2.MoveLast
                    RST2.MoveFirst
                    rst3.Open "SELECT     dbo.PAY_GETd.* FROM dbo.PAY_GETd WHERE (NUMBER = " & rst.Fields("NUMBER") & ") AND (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst3.MoveLast
                    rst3.MoveFirst
                    Forms![BUN].Form.Refresh
                    ST$ = Left(rst.Fields("NUMBER") & String(8, " "), 8) & Left(rst.Fields("tag") & String(8, " "), 8) & Left(rst.Fields("NUMBER1") & String(8, " "), 8) & Left(rst.Fields("DATE_N") & String(8, " "), 8) & Left(rst.Fields("TAH") & String(25, " "), 25) & Left(rst.Fields("MAS") & String(8, " "), 8) & Left(rst.Fields("VAS") & String(8, " "), 8) & Left(rst.Fields("CUST_NO") & String(40, " "), 40) & Left(rst.Fields("MOLAH") & String(50, " "), 50) & Left(rst.Fields("M_NAGHD") & String(12, " "), 12) & Left(rst.Fields("MABL_VAR") & String(12, " "), 12) & Left(rst.Fields("MOIN_VAR") & String(40, " "), 40) & Left(rst.Fields("MABL_HAV") & String(12, " "), 12) & Left(rst.Fields("MOIN_HAV") & String(40, " "), 40) & Left(rst.Fields("MABL_HAZ") & String(12, " "), 12) & Left(rst.Fields("MOIN_HAZ") & String(40, " "), 40) & Left(rst.Fields("TAKHFIF") & String(12, " "), 12) & Left(rst.Fields("FNUMCO") & String(8, " "), 8) & Left(rst.Fields("DEPATMAN") & String(4, " "), 4) & _
                                  Left(rst.Fields("SHIFT") & String(4, " "), 4) & Left(rst.Fields("CUST_KIND") & String(4, " "), 4) & Left(rst.Fields("USER_NAME") & String(40, " "), 40) & Left(rst.Fields("MBAA") & String(12, " "), 12) & Left(rst.Fields("HMBAA") & String(40, " "), 40) & Left(rst.Fields("TAMIR") & String(8, " "), 8) & Left(IIf(rst.Fields("TICMBAA"), 1, 0) & String(1, " "), 1) & Left(IIf(rst.Fields("OKF"), 1, 0) & String(1, " "), 1) & Left(rst.Fields("CDDATE") & String(8, " "), 8) & Left(rst.Fields("CDTIME") & String(4, " "), 4) & Left(rst.Fields("OKDATE") & String(8, " "), 8) & Left(rst.Fields("OKTIME") & String(4, " "), 4) & Left(RST2.RecordCount & String(4, " "), 4) & Left(rst3.RecordCount & String(4, " "), 4)
                    pp = InStr(1, ST$, Chr(13))
                    While pp > 0
                        ste$ = MID(ST$, 1, pp - 1) & " " & MID(ST$, pp + 1, Len(ST$) - pp)
                        ST$ = ste$
                        pp = InStr(1, ST$, Chr(13))
                    Wend
                    a.writeline (ST$)
                   ' ' DEBUG.PRINT st$
                    For i = 1 To RST2.RecordCount
                        a.writeline (Left(RST2.Fields("NUMBER") & String(8, " "), 8) & Left(RST2.Fields("tag") & String(8, " "), 8) & Left(RST2.Fields("ANBAR") & String(4, " "), 4) & Left(RST2.Fields("RADIF") & String(8, " "), 8) & Left(RST2.Fields("CODE") & String(15, " "), 15) & Left(RST2.Fields("MEGH") & String(8, " "), 8) & Left(RST2.Fields("MEGHk") & String(8, " "), 8) & Left(RST2.Fields("MEGH_MAR") & String(8, " "), 8) & Left(RST2.Fields("MANDAH") & String(50, " "), 50) & Left(RST2.Fields("MABL") & String(12, " "), 12) & Left(RST2.Fields("MABL_K") & String(12, " "), 12) & Left(RST2.Fields("N_KOL") & String(8, " "), 8) & Left(RST2.Fields("N_MOIN") & String(12, " "), 12) & Left(RST2.Fields("N_TAF") & String(8, " "), 8) & Left(RST2.Fields("IMBAA") & String(12, " "), 12) & Left(RST2.Fields("VAHED_K") & String(4, " "), 4))
                        RST2.MoveNext
                    Next i
                    InStr
                    Forms![BUN].Form.Refresh
                    For i = 1 To rst3.RecordCount
                        a.writeline (Left(rst3.Fields("N_SERI") & String(8, " "), 8) & Left(rst3.Fields("BANK") & String(4, " "), 4) & Left(rst3.Fields("DATE_S") & String(8, " "), 8) & Left(rst3.Fields("DATE") & String(8, " "), 8) & Left(rst3.Fields("SHOBEH") & String(20, " "), 20) & Left(rst3.Fields("MABL") & String(12, " "), 12) & Left(rst3.Fields("NAME_TAH") & String(60, " "), 60) & Left(rst3.Fields("N_HESAB") & String(30, " "), 30) & Left(rst3.Fields("N_KOL") & String(4, " "), 4) & Left(rst3.Fields("N_MOIN") & String(4, " "), 4) & Left(rst3.Fields("N_TAF") & String(4, " "), 4) & Left(rst3.Fields("NUMBER") & String(8, " "), 8) & Left(rst3.Fields("TAG") & String(8, " "), 8) & Left(rst3.Fields("RADIF") & String(8, " "), 8) & Left(rst3.Fields("CUST_NO") & String(20, " "), 20) & Left(rst3.Fields("KIND") & String(4, " "), 4))
                        rst3.MoveNext
                    Next i
                    rst.Fields("TKHF") = 0
                    rst.update
                    rst.MoveNext
                    Set RST2 = New ADODB.Recordset
                    Set rst3 = New ADODB.Recordset
                Wend
                 a.Close
                 DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال اطلاعات و تهيه دسيكت با موفقيت انجام شد با موفقيت به انجام رسيد ضروري است براي انتقال به سرور آنلاين بقيه كار را ادامه دهيد در غير اينصورت اطلاعات ازبين خواهد رفت"
            End If
        Else
            DoCmd.OpenForm "BUN"
            Dim incmd$, NUMBER, TAG, NUMBER1, DATE_N, TAH, MAS, VAS, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, MBAA, HMBAA, TAMIR, TICMBAA, OKF, CDDATE, CDTIME, OKDATE, OKTIME, invc, payc, factorn, havn, FFNUM, sttest
            Err.Clear
              i = 0
              Close #1
              Open "c:\sendoff.txt" For Input As #1
              If Err = 0 Then
                    rst.Open "head_lst", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    RSTI.Open "invo_lst", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rstP.Open "pay_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    While Not EOF(1)
                    incmd$ = ""
                            Line Input #1, incmd$
                           ' DEBUG.PRINT incmd$
                            Forms![BUN].[num].CAPTION = K
                            Forms![BUN].Form.Refresh
                            NUMBER = Trim(MID(incmd$, 1, 8))
                            TAG = Trim(MID(incmd$, 9, 8))
                            NUMBER1 = Trim(MID(incmd$, 17, 8))
                            DATE_N = Trim(MID(incmd$, 25, 8))
                            TAH = Trim(MID(incmd$, 33, 25))
                            MAS = Trim(MID(incmd$, 58, 8))
                            VAS = Trim(MID(incmd$, 66, 8))
                            CUST_NO = Trim(MID(incmd$, 74, 40))
                            MOLAH = Trim(MID(incmd$, 114, 50))
                            M_NAGHD = Trim(MID(incmd$, 164, 12))
                            MABL_VAR = Trim(MID(incmd$, 176, 12))
                            MOIN_VAR = Trim(MID(incmd$, 188, 40))
                            MABL_HAV = Trim(MID(incmd$, 228, 12))
                            MOIN_HAV = Trim(MID(incmd$, 240, 40))
                            MABL_HAZ = Trim(MID(incmd$, 280, 12))
                            MOIN_HAZ = Trim(MID(incmd$, 292, 40))
                            TAKHFIF = Trim(MID(incmd$, 332, 12))
                            FNUMCO = Trim(MID(incmd$, 344, 8))
                            DEPATMAN = Trim(MID(incmd$, 352, 4))
                            SHIFT = Trim(MID(incmd$, 356, 4))
                            CUST_KIND = Trim(MID(incmd$, 360, 4))
                            USER_NAME = Trim(MID(incmd$, 364, 40))
                            MBAA = Trim(MID(incmd$, 404, 12))
                            HMBAA = Trim(MID(incmd$, 416, 40))
                            TAMIR = Trim(MID(incmd$, 456, 8))
                            TICMBAA = Trim(MID(incmd$, 464, 1))
                            OKF = Trim(MID(incmd$, 465, 1))
                            CDDATE = Trim(MID(incmd$, 466, 8))
                            CDTIME = Trim(MID(incmd$, 474, 4))
                            OKDATE = Trim(MID(incmd$, 478, 8))
                            OKTIME = Trim(MID(incmd$, 486, 4))
                            invc = Trim(MID(incmd$, 490, 4))
                            payc = Trim(MID(incmd$, 494, 4))
                            If Not reapeated(NUMBER, TAG, CDDATE, CDTIME, DEPATMAN) Then
                                RST2.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                If rst.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                                     havn = Forms![BASEKNOW]![STHFR]
                                Else
                                    havn = RST2.Fields(0) + 1
                                End If
                                Forms![BUN].Form.Refresh
                               '----------------------------------------اضافه كردن هدر انبار------------
                                rst.AddNew
                                rst.Fields("NUMBER") = havn
                                rst.Fields("TAG") = 2
                              '  rst.Fields("NUMBER1") = Trim(Mid(incmd$, 17, 8))
                                rst.Fields("DATE_N") = DATE_N
                                rst.Fields("TAH") = TAH
                                rst.Fields("MAS") = MAS
                                rst.Fields("VAS") = VAS
                                rst.Fields("CUST_NO") = CUST_NO
                                rst.Fields("MOLAH") = MOLAH
                                rst.Fields("M_NAGHD") = M_NAGHD
                                rst.Fields("MABL_VAR") = MABL_VAR

                                rst.Fields("MOIN_VAR") = MOIN_VAR
                                rst.Fields("MABL_HAV") = MABL_HAV
                                rst.Fields("MOIN_HAV") = MOIN_HAV
                                rst.Fields("MABL_HAZ") = MABL_HAZ
                                rst.Fields("MOIN_HAZ") = MOIN_HAZ
                                rst.Fields("TAKHFIF") = TAKHFIF
                                rst.Fields("FNUMCO") = IIf(FNUMCO = "", Null, FNUMCO)
                                rst.Fields("DEPATMAN") = DEPATMAN
                                rst.Fields("SHIFT") = SHIFT
                                rst.Fields("CUST_KIND") = CUST_KIND
                                rst.Fields("USER_NAME") = USER_NAME
                                rst.Fields("MBAA") = MBAA
                                rst.Fields("HMBAA") = HMBAA
                                rst.Fields("TAMIR") = IIf(TAMIR = "", Null, TAMIR)
                                rst.Fields("TICMBAA") = TICMBAA
                                rst.Fields("OKF") = OKF
                                rst.Fields("CDDATE") = IIf(CDDATE = "", Null, CDDATE)
                                rst.Fields("CDTIME") = IIf(CDTIME = "", Null, CDTIME)
                                rst.Fields("OKDATE") = IIf(OKDATE = "", Null, OKDATE)
                                rst.Fields("OKTIME") = IIf(OKTIME = "", Null, OKTIME)
                                rst.Fields("anbar") = NUMBER
                                rst.update
                                Forms![BUN].Form.Refresh
                                '-------------------ايجاد شماره فاكتور
                                RST2.Close
                                Set RST2 = New ADODB.Recordset
                                RST2.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                If rst.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                                     factorn = Forms![BASEKNOW]![STFR]
                                Else
                                    factorn = RST2.Fields(0) + 1
                                End If
                                If K = 1 Then
                                    FFNUM = factorn
                                End If
                                K = K + 1
                                RST2.Close
                                Set RST2 = New ADODB.Recordset
                               '----------------------------------------------اضافه كردن هدر فاكتور------
                                rst.AddNew
                                rst.Fields("NUMBER") = havn
                                rst.Fields("TAG") = 13
                                rst.Fields("NUMBER1") = factorn
                                rst.Fields("DATE_N") = DATE_N
                                rst.Fields("TAH") = TAH
                                rst.Fields("MAS") = MAS
                                rst.Fields("VAS") = VAS
                                rst.Fields("CUST_NO") = CUST_NO
                                rst.Fields("MOLAH") = MOLAH
                                rst.Fields("M_NAGHD") = M_NAGHD
                                rst.Fields("MABL_VAR") = MABL_VAR
                                rst.Fields("MOIN_VAR") = MOIN_VAR
                                rst.Fields("MABL_HAV") = MABL_HAV
                                rst.Fields("MOIN_HAV") = MOIN_HAV
                                rst.Fields("MABL_HAZ") = MABL_HAZ
                                rst.Fields("MOIN_HAZ") = MOIN_HAZ
                                rst.Fields("TAKHFIF") = TAKHFIF
                                rst.Fields("FNUMCO") = IIf(FNUMCO = "", Null, FNUMCO)
                                rst.Fields("DEPATMAN") = DEPATMAN
                                rst.Fields("SHIFT") = SHIFT
                                rst.Fields("CUST_KIND") = CUST_KIND
                                rst.Fields("USER_NAME") = USER_NAME
                                rst.Fields("MBAA") = MBAA
                                rst.Fields("HMBAA") = HMBAA
                                rst.Fields("TAMIR") = IIf(TAMIR = "", Null, TAMIR)
                                rst.Fields("TICMBAA") = TICMBAA
                                rst.Fields("OKF") = OKF
                                rst.Fields("CDDATE") = IIf(CDDATE = "", Null, CDDATE)
                                rst.Fields("CDTIME") = IIf(CDTIME = "", Null, CDTIME)
                                rst.Fields("OKDATE") = IIf(OKDATE = "", Null, OKDATE)
                                rst.Fields("OKTIME") = IIf(OKTIME = "", Null, OKTIME)
                                rst.Fields("anbar") = NUMBER1
                                rst.update
                                '------------------------------------------آرتيكلهاي فاكتور
                                Forms![BUN].Form.Refresh
                                For i = 1 To invc
                                    Line Input #1, incmd$
                                    RSTI.AddNew
                                    RSTI.Fields("NUMBER") = havn
                                    RSTI.Fields("TAG") = 2
                                    RSTI.Fields("ANBAR") = Trim(MID(incmd$, 17, 4))
                                    RSTI.Fields("RADIF") = Trim(MID(incmd$, 21, 8))
                                    RSTI.Fields("CODE") = Trim(MID(incmd$, 29, 15))
                                    RSTI.Fields("MEGH") = Trim(MID(incmd$, 44, 8))
                                    RSTI.Fields("MEGHk") = Trim(MID(incmd$, 52, 8))
                                    RSTI.Fields("MEGH_MAR") = Trim(MID(incmd$, 60, 8))
                                    RSTI.Fields("MANDAH") = Trim(MID(incmd$, 68, 50))
                                    RSTI.Fields("MABL") = Trim(MID(incmd$, 118, 12))
                                    RSTI.Fields("MABL_K") = Trim(MID(incmd$, 130, 12))
                                    RSTI.Fields("N_KOL") = Trim(MID(incmd$, 142, 8))
                                    RSTI.Fields("N_MOIN") = Trim(MID(incmd$, 150, 8))
                                    RSTI.Fields("N_TAF") = IIf(Trim(MID(incmd$, 158, 8)) = "", Null, Trim(MID(incmd$, 158, 8)))
                                    RSTI.Fields("IMBAA") = Trim(MID(incmd$, 170, 12))
                                    RSTI.Fields("VAHED_K") = Trim(MID(incmd$, 182, 4))
                                    RSTI.update
                                Next i
                                '------------------------------------------آرتيكلهاي چكها
                                For i = 1 To payc
                                    Line Input #1, incmd$
                                    rstP.AddNew
                                    rstP.Fields("NUMBER") = havn
                                    rstP.Fields("TAG") = 2
                                    rstP.Fields("N_SERI") = Trim(MID(incmd$, 1, 8))
                                    rstP.Fields("BANK") = Trim(MID(incmd$, 9, 4))
                                    rstP.Fields("DATE_S") = Trim(MID(incmd$, 13, 8))
                                    rstP.Fields("DATE") = Trim(MID(incmd$, 21, 8))
                                    rstP.Fields("SHOBEH") = Trim(MID(incmd$, 29, 20))
                                    rstP.Fields("MABL") = Trim(MID(incmd$, 49, 12))
                                    rstP.Fields("NAME_TAH") = Trim(MID(incmd$, 61, 60))
                                    rstP.Fields("N_HESAB") = Trim(MID(incmd$, 121, 30))
                                    rstP.Fields("N_KOL") = Trim(MID(incmd$, 151, 4))
                                    rstP.Fields("N_MOIN") = Trim(MID(incmd$, 155, 4))
                                    rstP.Fields("N_TAF") = Trim(MID(incmd$, 159, 4))
                                    rstP.Fields("RADIF") = Trim(MID(incmd$, 179, 8))
                                    rstP.Fields("CUST_NO") = Trim(MID(incmd$, 187, 20))
                                    rstP.Fields("KIND") = Trim(MID(incmd$, 207, 4))
                                    rstP.update
                                Next i
                                Forms![BUN].Form.Refresh
                            Else
                            End If
                    Wend
                    Call GENSANADFROOSH(FFNUM, Int(factorn))
              End If
        End If
        DoCmd.Close acForm, "bun"
  End If
End Sub

Private Sub Command2_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Public Function reapeated(NUMBER, TAG, CDDATE, CDTIME, DEPATMAN) As Boolean
    Dim RRST As New ADODB.Recordset
    If CDDATE = "" Then
        CDDATE = 0
    End If
    If CDTIME = "" Then
        CDTIME = 0
    End If
'    rrst.Open "select count(NUMBER) as num from head_lst where NUMBER =" & NUMBER & " and tag = " & TAG & " and cddate = " & CDDATE & " and cdtime = " & CDTIME & " and DEPATMAN = " & DEPATMAN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'    If rrst.Fields("num") = 0 Then
 '       reapeated = False
  '  Else
  '      reapeated = True
   ' End If
End Function
