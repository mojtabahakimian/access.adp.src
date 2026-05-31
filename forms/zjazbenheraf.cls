Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command1_Click()
       On Error Resume Next
    Dim rst As New ADODB.Recordset, DT1 As Long, DT2 As Long, FNB As Double
    'آمار توليد به تفکيک اينماه
    DoCmd.OpenForm "BUN"
    DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 1)
    DT2 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 31)
    rst.Open "SELECT     dbo.INVO_LST.N_KOL, dbo.HEAD_MANF.CODE FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") GROUP BY dbo.INVO_LST.N_KOL, dbo.HEAD_MANF.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If copf Then
        While Not rst.EOF
            'کپي کردن فرمول
            Forms![BUN].Form.Refresh
           ' Debug.Print rst.Fields("n_kol") & " - " & rst.Fields("code")
          '  Debug.Print rst.RecordCount
                FNB = Copyformul(rst.Fields("n_kol"), Me.FMON)

            'آبديت توليدات با شماره فرمول جديد
            If uptolidf Then
                 DoCmd.RunSQL ("UPDATE dbo.INVO_LST set N_KOL = " & FNB & "  FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE  (DBO.INVO_LST.tag = 9 and N_KOL = " & rst.Fields("N_KOL") & " AND DATE_N BETWEEN " & DT1 & " AND " & DT2 & " )")
            End If
            rst.MoveNext
        Wend
    End If
       'محاسبهانحراف مقداري انبار گرداني
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'zanbekht" & Me.FMON & "')   DROP VIEW zanbekht" & Me.FMON)
    DoCmd.RunSQL ("create view dbo.zanbekht" & Me.FMON & " as SELECT     SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS ekh, dbo.ANBGRD_LST.CODE FROM         dbo.ANBGRD_HEAD INNER JOIN                       dbo.ANBGRD_LST ON dbo.ANBGRD_HEAD.GRD_NUM = dbo.ANBGRD_LST.GRD_NUM WHERE     (dbo.ANBGRD_HEAD.GRD_DATE BETWEEN " & DT1 & " AND " & DT2 & ") AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & Me.ANBARMAS & ") GROUP BY dbo.ANBGRD_LST.CODE HAVING      (SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) <> 0)")
    DoCmd.RunSQL ("alter  view dbo.zanbekht" & Me.FMON & " as SELECT     SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS ekh, dbo.ANBGRD_LST.CODE FROM         dbo.ANBGRD_HEAD INNER JOIN                       dbo.ANBGRD_LST ON dbo.ANBGRD_HEAD.GRD_NUM = dbo.ANBGRD_LST.GRD_NUM WHERE     (dbo.ANBGRD_HEAD.GRD_DATE BETWEEN " & DT1 & " AND " & DT2 & ") AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & Me.ANBARMAS & ") GROUP BY dbo.ANBGRD_LST.CODE HAVING      (SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) <> 0)")

    'محاسبه جمع مقادير مواد مصرفي
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'Zmavadmasrafe" & Me.FMON & "')   DROP VIEW Zmavadmasrafe" & Me.FMON)
    DoCmd.RunSQL ("create view dbo.Zmavadmasrafe" & Me.FMON & " as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = " & Me.FMON & ") GROUP BY code")
    DoCmd.RunSQL ("alter  view  dbo.Zmavadmasrafe" & Me.FMON & " as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = " & Me.FMON & ") GROUP BY code")
    'محاسبه درصد انحراف
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'zdarsadenheraf" & Me.FMON & "')   DROP VIEW zdarsadenheraf" & Me.FMON)
    DoCmd.RunSQL ("create view dbo.zdarsadenheraf" & Me.FMON & " as  SELECT     dbo.zanbekht" & Me.FMON & ".ekh / dbo.Zmavadmasrafe" & Me.FMON & ".MEGHk * 100 AS darsad, dbo.Zmavadmasrafe" & Me.FMON & ".code FROM         dbo.zanbekht" & Me.FMON & " INNER JOIN  dbo.Zmavadmasrafe" & Me.FMON & " ON dbo.zanbekht" & Me.FMON & ".CODE = dbo.Zmavadmasrafe" & Me.FMON & ".code where dbo.Zmavadmasrafe" & Me.FMON & ".MEGHk <>0 ")
    DoCmd.RunSQL ("alter  view dbo.zdarsadenheraf" & Me.FMON & " as  SELECT     dbo.zanbekht" & Me.FMON & ".ekh / dbo.Zmavadmasrafe" & Me.FMON & ".MEGHk * 100 AS darsad, dbo.Zmavadmasrafe" & Me.FMON & ".code FROM         dbo.zanbekht" & Me.FMON & " INNER JOIN  dbo.Zmavadmasrafe" & Me.FMON & " ON dbo.zanbekht" & Me.FMON & ".CODE = dbo.Zmavadmasrafe" & Me.FMON & ".code where dbo.Zmavadmasrafe" & Me.FMON & ".MEGHk <>0 ")
    'اعمال درصد انحراف مقادير روي فرمولها
    If UPMEGH Then
        DoCmd.RunSQL ("update DTL_MANF set dbo.DTL_MANF.MEGH = DTL_MANF.MEGH + DTL_MANF.MEGH * dbo.zdarsadenheraf" & Me.FMON & ".darsad /100 ,dbo.DTL_MANF.MEGHk = DTL_MANF.MEGHk + DTL_MANF.MEGHk * dbo.zdarsadenheraf" & Me.FMON & ".DARSAD / 100 from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zdarsadenheraf" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zdarsadenheraf" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zdarsadenheraf" & Me.FMON & ".DARSAD Is Null)) ")
    End If
    'قيمت ميانگين مواد
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'zgheymatmiangin" & Me.FMON & "')   DROP VIEW zgheymatmiangin" & Me.FMON)
    DoCmd.RunSQL ("create view dbo.zgheymatmiangin" & Me.FMON & " as SELECT    SUM(MABL_K)/ SUM(MEGHk)  AS fi, code FROM dbo.KALAS WHERE (TAG = 10) And (MM = " & Me.FMON & ") AND (MEGHk <> 0)  GROUP BY code ")
    DoCmd.RunSQL ("alter  view dbo.zgheymatmiangin" & Me.FMON & " as SELECT    SUM(MABL_K)/ SUM(MEGHk)  AS fi, code FROM dbo.KALAS WHERE (TAG = 10) And (MM = " & Me.FMON & ") AND (MEGHk <> 0)  GROUP BY code ")
    'اعمال مبلغ واقعي روي فرمولها
    If UPGHEYMAT Then
        DoCmd.RunSQL ("update DTL_MANF  set dbo.DTL_MANF.SMABL = dbo.zgheymatmiangin" & Me.FMON & ".fi,dbo.DTL_MANF.MABLK = ROUND(dbo.zgheymatmiangin" & Me.FMON & ".fi * dbo.DTL_MANF.MEGHk, 0) from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zgheymatmiangin" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zgheymatmiangin" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zgheymatmiangin" & Me.FMON & ".fi Is Null))")
    End If
    'حذف خروج موادبر اساس توليد
    If UPKHORUG Then
        DoCmd.RunSQL ("delete  from  INVO_LST WHERE TAG = 10  AND NUMBER  IN (select NUMBER from HEAD_LST  WHERE  (tag  = 10  AND DATE_N BETWEEN " & DT1 & " AND " & DT2 & " ))")
        'صادرکردن خروج مواد بر اساس توليد
     '   Call EXITMAVAD(Me.FMON)
    End If
    'حذف انبار گرداني
     If anb Then
        DoCmd.RunSQL ("delete  from  ANBGRD_LST WHERE GRD_NUM  IN (SELECT     dbo.ANBGRD_LST.GRD_NUM FROM         dbo.ANBGRD_HEAD INNER JOIN   dbo.ANBGRD_LST ON dbo.ANBGRD_HEAD.GRD_NUM = dbo.ANBGRD_LST.GRD_NUM WHERE     (dbo.ANBGRD_HEAD.GRD_DATE BETWEEN " & DT1 & " AND " & DT2 & " ) AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & Me.ANBARMAS & "))")
        Call AnbargrdUpdate(Me.FMON, Me.ANBARMAS)
    End If
    'بازسازي نرخ ميانگين
    If Me.BAZSA Then
       Call BAZSAZI(Me.FMON, 1)
        'اعمال مبلغ واقعي روي فرمولها
        If UPGHEYMAT Then
            DoCmd.RunSQL ("update DTL_MANF  set dbo.DTL_MANF.SMABL = dbo.zgheymatmiangin" & Me.FMON & ".fi,dbo.DTL_MANF.MABLK = ROUND(dbo.zgheymatmiangin" & Me.FMON & ".fi * dbo.DTL_MANF.MEGHk, 0) from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zgheymatmiangin" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zgheymatmiangin" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zgheymatmiangin" & Me.FMON & ".fi Is Null))")
        End If
       Call BAZSAZI(Me.FMON, 2)
        'اعمال مبلغ واقعي روي فرمولها
        If UPGHEYMAT Then
            DoCmd.RunSQL ("update DTL_MANF  set dbo.DTL_MANF.SMABL = dbo.zgheymatmiangin" & Me.FMON & ".fi,dbo.DTL_MANF.MABLK = ROUND(dbo.zgheymatmiangin" & Me.FMON & ".fi * dbo.DTL_MANF.MEGHk, 0) from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zgheymatmiangin" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zgheymatmiangin" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zgheymatmiangin" & Me.FMON & ".fi Is Null))")
        End If
       Call BAZSAZI(Me.FMON, 2)
        'اعمال مبلغ واقعي روي فرمولها
        If UPGHEYMAT Then
            DoCmd.RunSQL ("update DTL_MANF  set dbo.DTL_MANF.SMABL = dbo.zgheymatmiangin" & Me.FMON & ".fi,dbo.DTL_MANF.MABLK = ROUND(dbo.zgheymatmiangin" & Me.FMON & ".fi * dbo.DTL_MANF.MEGHk, 0) from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zgheymatmiangin" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zgheymatmiangin" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zgheymatmiangin" & Me.FMON & ".fi Is Null))")
        End If
       Call BAZSAZI(Me.FMON, 2)
        'اعمال مبلغ واقعي روي فرمولها
        If UPGHEYMAT Then
            DoCmd.RunSQL ("update DTL_MANF  set dbo.DTL_MANF.SMABL = dbo.zgheymatmiangin" & Me.FMON & ".fi,dbo.DTL_MANF.MABLK = ROUND(dbo.zgheymatmiangin" & Me.FMON & ".fi * dbo.DTL_MANF.MEGHk, 0) from  dbo.INVO_LST INNER JOIN    dbo.HEAD_MANF ON dbo.INVO_LST.N_KOL = dbo.HEAD_MANF.FNUMB INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB LEFT OUTER JOIN  dbo.zgheymatmiangin" & Me.FMON & " ON dbo.DTL_MANF.CODE = dbo.zgheymatmiangin" & Me.FMON & ".code WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") And (Not (dbo.zgheymatmiangin" & Me.FMON & ".fi Is Null))")
        End If
       Call BAZSAZI(Me.FMON, 2)
    End If

    DoCmd.Close acForm, "BUN"
End Sub

Public Function Copyformul(fnum As Long, mm As Long) As Double
       On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, FNB As Double, DT1 As Long, i As Integer
    RST2.Open "HEAD_MANF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SELECT * FROM HEAD_MANF WHERE FNUMB =  " & fnum, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + mm * 100 + 1)
    If rst.RecordCount = 1 Then
        FNB = mm & rst.Fields("CODE")
        i = 0
        While Isformul(FNB) And i < 100
            FNB = FNB & i
            i = i + 1
        Wend
        RST2.AddNew
        RST2.Fields("FNUMB") = FNB
        RST2.Fields("CODE") = rst.Fields("CODE")
        RST2.Fields("DATE_ACTIV") = DT1
        RST2.Fields("SA_HOUR") = rst.Fields("SA_HOUR")
        RST2.Fields("SA_NHOU") = rst.Fields("SA_NHOU")
        RST2.Fields("IMBIBE_MANF") = rst.Fields("IMBIBE_MANF")
        RST2.Fields("IMBIBE_SAR") = rst.Fields("IMBIBE_SAR")
        RST2.Fields("TOZIH") = MAHNAME(mm)
        RST2.update
        DoCmd.RunSQL ("INSERT INTO DTL_MANF ( FNUMB, CODE, ANBAR, VAHED_K, MEGH, MEGHk, PERT, SMABL, MABLK, TOZIH) SELECT  " & FNB & ", CODE, ANBAR, VAHED_K, MEGH, MEGHk, PERT, SMABL, MABLK, TOZIH FROM DTL_MANF WHERE  FNUMB =  " & fnum)
        Copyformul = FNB
    End If
End Function

Public Sub EXITMAVAD(mm As Long)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, HRST As New ADODB.Recordset, rstf As New ADODB.Recordset, RSTK As New ADODB.Recordset, RSTM As New ADODB.Recordset, MABLTMP As Long
   Dim num As Long, DT1 As Long, DT2 As Long
   DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 1)
   DT2 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 31)

  DoCmd.OpenForm "BUN"
  HRST.Open "SELECT * FROM HEAD_LST WHERE     (TAG = 9) AND (DATE_N BETWEEN " & DT1 & " AND " & DT2 & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not HRST.EOF
     If Forms![BASEKNOW]![ECONM] Then
       If MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) <> "5" Then
            If Not IsNull(HRST.Fields("NUMBER1")) Then
                rst.Open "SELECT * FROM HEAD_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 1 Then
                   DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10"
                   RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & HRST.Fields("NUMBER") & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not RSTK.EOF()
                       rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.CODE = '" & RSTK.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rstf.EOF()
                           RSTM.AddNew
                           RSTM.Fields("NUMBER") = HRST.Fields("NUMBER1")
                           RSTM.Fields("TAG") = 10
                           RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                           RSTM.Fields("CODE") = rstf.Fields("CODB")
                           RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                           RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                           MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), HRST.Fields("DATE_N"))
                           RSTM.Fields("MABL") = MABLTMP
                           RSTM.Fields("AVRAGE") = MABLTMP
                           RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.update
                           rstf.MoveNext
                       Wend
                       rstf.Close
                       Set rst = New ADODB.Recordset
                       RSTK.MoveNext
                       Forms![BUN].Form.Refresh
                   Wend
                   Call SANADKHORUGMAVAD(HRST.Fields("NUMBER1"), HRST.Fields("NUMBER1"))
                Else
                End If
            Else
                rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                    num = 1
                Else
                   num = rst.Fields(0) + 1
                End If
                Set rst = New ADODB.Recordset
                rst.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("NUMBER") = num
                rst.Fields("tag") = 10
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.Fields("DATE_N") = HRST.Fields("DATE_N")
                rst.Fields("FNUMCO") = HRST.Fields("NUMBER")
                rst.Fields("CUST_NO") = HRST.Fields("CUST_NO")
                rst.Fields("MOLAH") = "بر اساس توليد"
                rst.update
                RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & HRST.Fields("NUMBER") & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RSTK.EOF()
                   rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.CODE = '" & RSTK.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rstf.EOF()
                       RSTM.AddNew
                       RSTM.Fields("NUMBER") = num
                       RSTM.Fields("TAG") = 10
                       RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                       RSTM.Fields("CODE") = rstf.Fields("CODB")
                       RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                       RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                       MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), HRST.Fields("DATE_N"))
                       RSTM.Fields("MABL") = MABLTMP
                       RSTM.Fields("AVRAGE") = MABLTMP
                       RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.update
                       rstf.MoveNext
                      Forms![BUN].Form.Refresh
                   Wend
                   Set rst = New ADODB.Recordset
                   RSTK.MoveNext
               Wend
               HRST.Fields("NUMBER1") = num
               HRST.update
               Call SANADKHORUGMAVAD(num, num)
            End If
        Else
            If Not IsNull(HRST.Fields("NUMBER1")) Then
                rst.Open "SELECT * FROM HEAD_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 1 Then
                   DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10"
                   Set RSTK = New ADODB.Recordset
                   RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & HRST.Fields("NUMBER") & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not RSTK.EOF()
                       rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.FNUMB = '" & IIf(IsNull(RSTK.Fields("n_kol")), 0, RSTK.Fields("n_kol")) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rstf.EOF()
                           RSTM.AddNew
                           RSTM.Fields("NUMBER") = HRST.Fields("NUMBER1")
                           RSTM.Fields("TAG") = 10
                           RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                           RSTM.Fields("CODE") = rstf.Fields("CODB")
                           RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                           RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                           MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), HRST.Fields("DATE_N"))
                           RSTM.Fields("MABL") = MABLTMP
                           RSTM.Fields("AVRAGE") = MABLTMP
                           RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.update
                           rstf.MoveNext
                          Forms![BUN].Form.Refresh
                       Wend
                       rstf.Close
                       Set rst = New ADODB.Recordset
                       RSTK.MoveNext
                   Wend
                   Call SANADKHORUGMAVAD(HRST.Fields("NUMBER1"), HRST.Fields("NUMBER1"))
                Else
                End If
            Else
                rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                    num = 1
                Else
                   num = rst.Fields(0) + 1
                End If
                Set rst = New ADODB.Recordset
                rst.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("NUMBER") = num
                rst.Fields("tag") = 10
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.Fields("DATE_N") = HRST.Fields("DATE_N")
                rst.Fields("FNUMCO") = HRST.Fields("NUMBER")
                rst.Fields("CUST_NO") = HRST.Fields("CUST_NO")
                rst.Fields("MOLAH") = "بر اساس توليد"
                rst.update
                RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & HRST.Fields("NUMBER") & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RSTK.EOF()
                   rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.FNUMB = '" & RSTK.Fields("n_kol") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rstf.EOF()
                       RSTM.AddNew
                       RSTM.Fields("NUMBER") = num
                       RSTM.Fields("TAG") = 10
                       RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                       RSTM.Fields("CODE") = rstf.Fields("CODB")
                       RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                       RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                       MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), HRST.Fields("DATE_N"))
                       RSTM.Fields("MABL") = MABLTMP
                       RSTM.Fields("AVRAGE") = MABLTMP
                       RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.update
                       rstf.MoveNext
                      Forms![BUN].Form.Refresh
                   Wend
                   rstf.Close
                   Set rst = New ADODB.Recordset
                   RSTK.MoveNext
               Wend
               HRST.Fields("NUMBER1") = num
               HRST.update
               Call SANADKHORUGMAVAD(num, num)
            End If
        End If
     End If
     HRST.MoveNext
 Wend


End Sub

Public Function Isformul(FNUMB As Double) As Boolean
        On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "select fnumb from head_manf where fnumb = " & FNUMB, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 1 Then
       Isformul = True
    Else
        Isformul = False
    End If
End Function

Public Sub AnbargrdUpdate(mm As Long, anb As Long)
         On Error Resume Next
  Dim rst As New ADODB.Recordset, DT1 As Long, DT2 As Long
    DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 1)
    DT2 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 31)
    rst.Open "select * from ANBGRD_HEAD where GRD_ANBAR = " & anb & "  AND GRD_DATE BETWEEN  " & DT1 & " AND " & DT2 & " ORDER BY GRD_DATE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        DoCmd.RunSQL "INSERT INTO dbo.ANBGRD_LST  (CODE, MOG, GRD_NUM) SELECT     CODE, MAND, " & rst.Fields("GRD_NUM") & " AS GN FROM         dbo.MOGUDI(" & rst.Fields("GRD_DATE") & "," & rst.Fields("GRD_ANBAR") & ") MOGUDI"
        rst.MoveNext
    Wend
End Sub


Private Sub BAZSAZI(mm As Long, GRP As Integer)
        On Error Resume Next
  Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST6 As New ADODB.Recordset
   Dim CON, i As Long, MIAN As Double, MBKM As Double, MOGUDI As Double, DT1 As Long, DT2 As Long
   Dim cn As New ADODB.Connection, cm As New ADODB.Command
   DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + mm * 100)
   DT2 = CLng(Forms![BASEKNOW]![YEA] * 10000 + mm * 100 + 32)
    DoCmd.OpenForm "BUN"
   rst3.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If GRP = 1 Then
        rst.Open "SELECT TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH <> 2 AND  dbo.STUF_DEF.RADAH <> 3 ) ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Else
        rst.Open "SELECT TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH = 2 OR  dbo.STUF_DEF.RADAH = 3 ) ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   End If
   RST6.Open "select count(id) from invo_lst where tag <> 20 and tag <> 23", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.TEDADP = RST6.Fields(0)
   Set RST6 = New ADODB.Recordset
   RST6.Open "ANBGRD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   cm.ActiveConnection = CurrentProject.Connection
   cm.CommandTimeout = 3600
   cm.CommandText = "SELECT TOP 100 PERCENT DATE_N, TAG, NUMBER, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID,  MEGH_R , RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, ID, BARGAH FROM dbo.TMPAV101  where date_n between  " & DT1 & " and " & DT2 & " ORDER BY DATE_N, BARGAH"
   If MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) = "5" Then
        While Not rst.EOF
            Me.co = rst.Fields("CODE")
            MIAN = LASTAVRAGE(rst.Fields("CODE"), rst.Fields("ANBAR"), DT1)
            If MIAN = 0 Then
                MIAN = GETSTANDARDPRICE(rst.Fields("CODE"))
                If MIAN = 0 Then
                    MIAN = GETFIRSTPRICE(rst.Fields("CODE"))
                End If
            End If
            MOGUDI = GetMandMog(rst.Fields("CODE"), rst.Fields("ANBAR"), DT1)
            MBKM = MOGUDI * MIAN
            CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
            CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & DT1 & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1  = dbo.TAGCOD.CODE " & _
              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & DT1 & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & DT1 & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
              " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & DT1 & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER  union " & _
              " SELECT     TOP 100 PERCENT dbo.ANBGRD_HEAD.GRD_DATE, dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG,dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS radif, dbo.ANBGRD_LST.CODE,(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEGK, 0 AS megh_mar,' ' AS mol, dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 0 AS froma, '' AS nrasid, 0 AS MEGH_R, dbo.STUF_DEF.RADAH, 0 AS sanadno, '  ' AS cust_no, 0 AS anbarf, dbo.STUF_DEF.VAHED, 0 AS n_kol, 0 AS n_moin, 0 AS n_taf, dbo.ANBGRD_LST.MABL AS avrage, 0 AS id, '17' AS Expr1 FROM   dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN  dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE " & _
              " WHERE      (dbo.ANBGRD_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & rst.Fields("ANBAR") & ") and ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0) AND (dbo.ANBGRD_HEAD.GRD_DATE > " & DT1 & ") AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) ORDER BY dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM ")
            Set RST2 = New ADODB.Recordset
            Set RST2 = cm.Execute
            While Not RST2.EOF
                 TEDADP = TEDADP + 1
                 Me.Repaint
                 rst3.Filter = "ID = " & RST2.Fields("ID")
                 Set rst3 = rst3
                 If RST2.Fields("CODE") = "2534" Then
                 DoCmd.Beep
                 End If
                ' If MIAN < 0 Then
                '          MIAN = 0
               '  End If
                 Select Case RST2.Fields("TAG")
                     Case 1: 'خريد
                          MBKM = MBKM + RST2.Fields("MABL_K")
                          MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                          If MBKM = 0 Then
                       '     MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                        '         MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.update
                     Case 22: 'برگشت فروش سال قبل
                          MBKM = MBKM + MIAN * RST2.Fields("MEGH_MAR")
                          MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                          If MBKM = 0 Then
                         '    MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                          '       MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.update
                     Case 24: 'برگشت فروش سال قبل
                          MBKM = MBKM + RST2.Fields("MEGHK") * MIAN
                          MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                          If MBKM = 0 Then
                           '  MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                            '     MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.update
                     Case 2: 'فروش
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.update
                     Case 3: 'برگشت خريد
                          MBKM = MBKM - (RST2.Fields("MEGH_MAR") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGH_MAR")
                          If MBKM = 0 Then
                             'MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                              '   MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE2") = MIAN
                          rst3.update
                     Case 4: 'برگشت فروش
                          MBKM = MBKM + (RST2.Fields("MEGH_MAR") * rst3.Fields("AVRAGE"))
                          MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                          If MBKM = 0 Then
                             'MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                              '   MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE2") = MIAN
                          rst3.update
                     Case 5: 'انتقالي خروج
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.Fields("MABL") = MIAN
                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                          rst3.update
                     Case 6: 'انتقالي ورود
                          MBKM = MBKM + RST2.Fields("MABL_K")
                          MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                          If MBKM = 0 Then
                             'MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                              '   MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                          End If
                          rst3.Fields("AVRAGE2") = MIAN
                          rst3.update
                     Case 10: 'مواد خروج
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.Fields("MABL") = MIAN
                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                          rst3.update
                     Case 11:    'موادساير خروج
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.Fields("MABL") = MIAN
                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                          rst3.update
                    Case 9:    'توليد
                          If RST2.Fields("n_kol") <> 0 And Not IsNull(RST2.Fields("n_kol")) And MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) = "5" Then
                              Set rst4 = New ADODB.Recordset
                              rst4.Open "SELECT  dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, SUM(dbo.DTL_MANF.MABLK) AS MABLKs FROM         dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE (dbo.HEAD_MANF.FNUMB = " & RST2.Fields("n_kol") & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                              If rst4.RecordCount > 0 Then
                                 rst3.Fields("MABL") = rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)
                                 rst3.Fields("MABL_k") = Round((rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)) * RST2.Fields("MEGHK"))
                                 rst3.update
                              End If
                           Else
                                rst3.Fields("MABL") = GETSTANDARDPRICE_KOL(rst.Fields("CODE"), RST2.Fields("date_n"))
                                If rst3.Fields("MABL") = 0 Then
                                    rst3.Fields("MABL") = GETFIRSTPRICE(rst.Fields("CODE"))
                                End If
                                rst3.Fields("MABL_k") = Round(rst3.Fields("MABL") * RST2.Fields("MEGHK"))
                                rst3.update
                           End If
                           MBKM = MBKM + rst3.Fields("MABL_k")
                           MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                           If MBKM = 0 Then
                             'MIAN = 0
                           Else
                             If MOGUDI = 0 Then
                                ' MIAN = 0
                                 MBKM = 0
                             Else
                                 MIAN = MBKM / MOGUDI
                             End If
                           End If
                           rst3.Fields("AVRAGE") = MIAN
                           rst3.update
                     Case 17: 'كسري انبار
                          MBKM = MBKM + MIAN * RST2.Fields("MEGHK")
                          MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                          If MBKM = 0 Then
                           '  MIAN = 0
                          Else
                             If MOGUDI = 0 Then
                           '      MIAN = 0
                                 MBKM = 0
                            Else
                                MIAN = MBKM / MOGUDI
                             End If
                          End If
                           ' If MIAN < 0 Then
                            '   MIAN = 0
                            'End If
                            RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                            Set RST6 = RST6
                            RST6.Fields("MABL") = MIAN
                            RST6.update
                     Case 18: 'فروش
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                          Set RST6 = RST6
                          RST6.Fields("MABL") = MIAN
                          RST6.update
                     Case 26: 'برگشت خريد
                          MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                          MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                          rst3.Fields("AVRAGE") = MIAN
                          rst3.update
                 End Select
                 RST2.MoveNext
              Wend
'              RST2.MovePrevious
            If GRP = 2 Or GRP = 3 Then
              DoCmd.RunSQL "UPDATE    dbo.DTL_MANF  SET  MABLK = (MEGHk + PERT) * " & MIAN & ", SMABL = " & MIAN & " WHERE     CODE = N'" & rst.Fields("CODE") & "' and dbo.DTL_MANF.FNUMB in (SELECT     dbo.INVO_LST.N_KOL FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") GROUP BY dbo.INVO_LST.N_KOL)"
            End If
              RST2.Close
              rst.MoveNext
              DoCmd.OpenForm "BUN"
              Forms![BUN].Form.Refresh
              DoEvents
        Wend
    End If
End Sub

Private Sub Detail_DblClick(CANCEL As Integer)
    Me.UPKHORUG = Not Me.UPKHORUG
    Me.copf = Not Me.copf
    Me.uptolidf = Not Me.uptolidf
    Me.UPGHEYMAT = Not Me.UPGHEYMAT
    Me.UPMEGH = Not Me.UPMEGH
    Me.BAZSA = Not Me.BAZSA
    Me.anb = Not Me.anb
End Sub

Private Sub FMON_AfterUpdate()
    Dim rst As New ADODB.Recordset, DT1 As Long, DT2 As Long
    DT1 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100)
    DT2 = CLng(Forms![BASEKNOW]![YEA] * 10000 + Me.FMON * 100 + 32)
    rst.Open "SELECT   SUM(dbo.DEED_DTL.BES) AS dast FROM   dbo.DEED_DTL INNER JOIN  dbo.HEAD_LST ON dbo.DEED_DTL.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.DEED_DTL.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.DEED_DTL.HES_T = 99999999) AND (dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![CONKAL] & ") AND dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.DAST = rst.Fields("dast")
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT   SUM(dbo.DEED_DTL.BES) AS SARB FROM   dbo.DEED_DTL INNER JOIN  dbo.HEAD_LST ON dbo.DEED_DTL.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.DEED_DTL.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.DEED_DTL.HES_T = 99999998) AND (dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![CONKAL] & ") AND dbo.HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.SARB = rst.Fields("SARB")
    End If
End Sub
