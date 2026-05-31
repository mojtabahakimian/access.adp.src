Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command0_Click()
    Dim cn As New ADODB.Connection, cm As New ADODB.Command
    Dim rst As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim infileopen As Boolean, BANKK As Long
    Dim BASES As Long, PGETID As Long, FID As Long, IID As Long, DTT As Long, dtt2 As Long
    Dim incmd$, batch$, Q$
    rst.Open "SELECT MAX(BASE) AS BS FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    BASES = rst.Fields(0)
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(ID) AS BS FROM PGET_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    PGETID = rst.Fields(0)
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(DATE_N) AS BS FROM HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    FID = rst.Fields(0)
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(ID) AS BS FROM INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    IID = rst.Fields(0)
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(date) AS BS FROM CHKREC_H", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    DTT = rst.Fields(0)
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT MAX(date) AS BS FROM CHREC_HP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    dtt2 = rst.Fields(0)
    Q$ = """"  'quoted id char
                     On Error Resume Next
    'Open up a new connection direct to SQL Server

    cn.connectionString = CurrentProject.BaseConnectionString
    cn.Open

    cm.ActiveConnection = cn
    cm.CommandType = adCmdText
    Err.Clear
          Set cn = CurrentProject.Connection
          Set rs = cn.Execute("INSERT INTO dbo.tcod_hesvaz (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.tcod_hesvaz WHERE (NOT (CODE IN (SELECT code FROM tcod_hesvaz)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_HESKIND (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.TCOD_HESKIND WHERE (NOT (CODE IN (SELECT code FROM TCOD_HESKIND)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_HESGROUP (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.TCOD_HESGROUP WHERE (NOT (CODE IN (SELECT code FROM TCOD_HESGROUP)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_STUFGROUP (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.TCOD_STUFGROUP WHERE (NOT (CODE IN (SELECT code FROM TCOD_STUFGROUP)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_VAHEDS (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.TCOD_VAHEDS WHERE (NOT (CODE IN (SELECT code FROM TCOD_VAHEDS)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_BANKS (  CODE, NAMES  ) SELECT  CODE, NAMES FROM " & PATH & ".dbo.TCOD_BANKS WHERE (NOT (CODE IN (SELECT code FROM TCOD_BANKS)))")
          Set rs = cn.Execute("INSERT INTO dbo.TCOD_ANBAR (  CODE, NAMES,KIND  ) SELECT  CODE, NAMES,KIND FROM " & PATH & ".dbo.TCOD_ANBAR WHERE (NOT (CODE IN (SELECT code FROM TCOD_ANBAR)))")
          Set rs = cn.Execute("INSERT INTO dbo.TOTA_HES (  NUMBER, NAME, NO_HES, M_D, [GROUP] ) SELECT     NUMBER, NAME, NO_HES, M_D, [GROUP] FROM  " & PATH & ".dbo.TOTA_HES WHERE (NOT (NUMBER IN (SELECT NUMBER FROM tota_hes)))")
          Set rs = cn.Execute("INSERT INTO dbo.DETA_HES (N_KOL, NUMBER, NAME, TOZIH, BED_BES, ADDRESS, TEL, CODE_E) SELECT N_KOL, NUMBER, NAME, TOZIH, BED_BES, ADDRESS, TEL, CODE_E FROM " & PATH & ".dbo.deta_hes WHERE (NOT ((CAST(N_KOL AS float) * 100000000 + NUMBER) IN (SELECT     CAST(N_KOL AS float) * 100000000 + NUMBER  FROM deta_hes)))")
          Set rs = cn.Execute("INSERT INTO dbo.TDETA_HES   (N_KOL, NUMBER, TNUMBER, NAME, TOZIH) SELECT     N_KOL, NUMBER, TNUMBER, NAME, TOZIH FROM " & PATH & ".dbo.TDETA_HES WHERE (NOT (((CAST(N_KOL AS float) * 10000 + NUMBER) * 100000000 + TNUMBER) IN (SELECT (CAST(N_KOL AS float) * 10000 + NUMBER) * 100000000 + TNUMBER FROM tdeta_hes)))")
          Set rs = cn.Execute("INSERT INTO dbo.STUF_DEF (CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART) SELECT     CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART FROM " & PATH & ".dbo.stuf_def WHERE     (NOT (CODE IN (SELECT code  FROM stuf_def)))")
          Set rs = cn.Execute("INSERT INTO dbo.STUF_FSK (CODE, ANBAR, MOGODI_A, FI_A, MABL_A, MANDAH_A, VAZ) SELECT  CODE, ANBAR, MOGODI_A, FI_A, MABL_A, MANDAH_A, VAZ FROM " & PATH & ".dbo.STUF_FSK WHERE (NOT ((RIGHT('               ' + CODE, 15) + CAST(ANBAR AS nvarchar)) IN  (SELECT     RIGHT('               ' + CODE, 15) + CAST(ANBAR AS nvarchar)  FROM stuf_fsk)))")
          Set rs = cn.Execute("INSERT INTO dbo.STUF_STK (CODE, ANBAR, MOGODI_A, MOGODI, MABL_M) SELECT     CODE, ANBAR, MOGODI_A, MOGODI, MABL_M FROM   " & PATH & ".dbo.STUF_STK WHERE     (NOT ((RIGHT('               ' + CODE, 15) + CAST(ANBAR AS nvarchar)) IN (SELECT     RIGHT('               ' + CODE, 15) + CAST(ANBAR AS nvarchar)   FROM   stuf_stk)))")

          Set rs = cn.Execute("INSERT INTO dbo.PAY_GETD  (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ) SELECT  N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3 , N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ FROM " & PATH & ".dbo.pay_getd WHERE     (NOT ((CAST(N_SERI AS nchar(15)) + CAST(BANK AS nchar(2)) + CAST(DATE_S AS nvarchar)) IN (SELECT  CAST(N_SERI AS nchar(15)) + CAST(BANK AS nchar(2)) + CAST(DATE_S AS nvarchar) FROM  pay_getd)))")
          Set rs = cn.Execute("INSERT INTO dbo.PAY_GETP (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, N_MOIN3 , N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO FROM " & PATH & ".dbo.pay_getp WHERE     (NOT ((CAST(N_SERI AS nchar(15)) + CAST(BANK AS nchar(2)) + CAST(DATE_S AS nvarchar)) IN  (SELECT     CAST(N_SERI AS nchar(15)) + CAST(BANK AS nchar(2)) + CAST(DATE_S AS nvarchar)  FROM pay_getp)))")
          Set rs = cn.Execute("INSERT INTO dbo.DEED_HED (  N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME ) SELECT   N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME  from " & PATH & ".DBO.DEED_HED where base > " & BASES)
          Set rs = cn.Execute("INSERT INTO dbo.CHKREC_H (DATE, MOLAH, N_S) SELECT     DATE, MOLAH, N_S FROM " & PATH & ".dbo.chkrec_h WHERE     (NOT (DATE IN (SELECT date FROM chkrec_h)))")
          Set rs = cn.Execute("INSERT INTO dbo.CHRE_LST (N_SERI, BANK, DATE_S, DATE, RADIF, N_MOIN, N_TAF) SELECT     N_SERI, BANK, DATE_S, DATE, RADIF, N_MOIN, N_TAF FROM " & PATH & ".dbo.CHRE_LST WHERE (DATE > " & DTT & ")")
          Set rs = cn.Execute("INSERT INTO dbo.CHREC_HP (DATE, MOLAH, N_S) SELECT DATE, MOLAH, N_S FROM " & PATH & ".dbo.CHREC_HP WHERE  (DATE > " & dtt2 & ")")
          Set rs = cn.Execute("INSERT INTO dbo.CHRE_LSP (N_SERI, BANK, DATE_S, DATE, RADIF, N_MOIN, N_TAF) SELECT     N_SERI, BANK, DATE_S, DATE, RADIF, N_MOIN, N_TAF FROM " & PATH & ".dbo.CHRE_LSP WHERE (DATE > " & dtt2 & ")")
          Set rs = cn.Execute("INSERT INTO dbo.PGET_HED (ID, DATE, MOLAH, N_S, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, KIND, IDK ) SELECT  " & PATH & ".dbo.PGET_HED.*  from " & PATH & ".DBO.PGET_HED where ID > " & PGETID)
          Set rs = cn.Execute("INSERT INTO dbo.PGET_LST  (ID , DATE, RADIF, NO_AM, NAHVA, FHES_K, FHES_M, FHES_T, THES_K, THES_M, THES_T, SHARH, mabl, N_SERI, BANK, FHES, THES) SELECT  ID , DATE, RADIF, NO_AM, NAHVA, FHES_K, FHES_M, FHES_T, THES_K, THES_M, THES_T, SHARH, mabl, N_SERI, BANK, FHES, THES  from " & PATH & ".DBO.PGET_LST where ID > " & PGETID)
          Set rs = cn.Execute("INSERT INTO dbo.HEAD_LST (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV,MABL_HAZ , MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME)  SELECT  NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV,MABL_HAZ , MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME  from " & PATH & ".DBO.HEAD_LST where DATE_N > " & FID)
          Set rs = cn.Execute("INSERT INTO dbo.INVO_LST (NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO,CUST_NO , ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE)  SELECT  " & PATH & ".DBO.INVO_LST.NUMBER, " & PATH & ".DBO.INVO_LST.TAG, " & PATH & ".DBO.INVO_LST.ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO," & PATH & ".DBO.INVO_LST.CUST_NO , " & PATH & ".dbo.INVO_LST.ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE  from " & PATH & ".dbo.HEAD_LST INNER JOIN   " & PATH & ".dbo.INVO_LST ON " & PATH & ".dbo.HEAD_LST.NUMBER = " & PATH & ".dbo.INVO_LST.NUMBER AND  " & PATH & ".dbo.HEAD_LST.TAG = " & PATH & ".dbo.INVO_LST.TAG  where " & PATH & ".dbo.INVO_LST.ID > " & IID & " AND DATE_N > " & FID)
          Set rs = cn.Execute("INSERT INTO dbo.DEED_DTL (N_S, RADIF, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES) SELECT " & PATH & ".dbo.DEED_DTL.N_S, " & PATH & ".dbo.DEED_DTL.RADIF, " & PATH & ".dbo.DEED_DTL.HES_K, " & PATH & ".dbo.DEED_DTL.HES_M," & PATH & ".dbo.DEED_DTL.HES_T, " & PATH & ".dbo.DEED_DTL.SHARH, " & PATH & ".dbo.DEED_DTL.BED, " & PATH & ".dbo.DEED_DTL.BES," & PATH & ".dbo.DEED_DTL.N_SERI, " & PATH & ".dbo.DEED_DTL.BANK, " & PATH & ".dbo.DEED_DTL.NUMBER, " & PATH & ".dbo.DEED_DTL.TAG," & PATH & ".dbo.DEED_DTL.HES FROM " & PATH & ".dbo.DEED_DTL INNER JOIN " & PATH & ".dbo.DEED_HED ON " & PATH & ".dbo.DEED_DTL.N_S = " & PATH & ".dbo.DEED_HED.N_S WHERE (" & PATH & ".dbo.DEED_HED.base >" & BASES & ")")
          DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ايجاد سال با موفقيت انجام شد"

    cn.Close
End Sub

Private Sub Command4_Click()
                   On Error Resume Next
 DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
