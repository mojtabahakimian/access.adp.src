Attribute VB_Name = "HESABDARI"
Option Compare Database


Option Explicit

Public Sub HAVALAH_ENTER_mavad_genereat(DT1, DT2 As Long)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, rstf As New ADODB.Recordset, RSTK As New ADODB.Recordset, RSTM As New ADODB.Recordset, MABLTMP As Long
   Dim num As Long, HRST As New ADODB.Recordset
   DoCmd.OpenForm "BUN"
  RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HRST.Open "SELECT * FROM HEAD_LST WHERE (DATE_N >=  " & DT1 & " AND  DATE_N <=" & DT2 & ") AND (TAG = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not HRST.EOF()
  DoEvents
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
                'RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
               Call SANADKHORUGMAVAD(num, num)
            End If
        Else
            If Not IsNull(HRST.Fields("NUMBER1")) Then
                rst.Open "SELECT * FROM HEAD_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 1 Then
                   DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & HRST.Fields("NUMBER1") & " AND TAG = 10"
                  ' RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
             '   RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
               Call SANADKHORUGMAVAD(num, num)
            End If
        End If
     End If

   Set rstf = New ADODB.Recordset
   Set RSTK = New ADODB.Recordset
   Set rst = New ADODB.Recordset
   HRST.MoveNext
 Wend
   DoCmd.Close acForm, "bun"
End Sub

Public Function FIRSTM(KOL As Double) As Double
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL) = " & KOL & " )) HAVING ((Not (Min(DETA_HES.NUMBER)) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        FIRSTM = rst.Fields(0)
    Else
        FIRSTM = Null
    End If
    rst.Close
End Function
Public Function FIRSTT(KOL As Double, MOIN As Double) As Double
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Min(TDETA_HES.TNUMBER) AS MinOfTNUMBER FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) = " & MOIN & ") And ((TDETA_HES.N_KOL) = " & KOL & ")) HAVING ((Not (Min(TDETA_HES.TNUMBER)) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        FIRSTT = rst.Fields(0)
    Else
        FIRSTT = Null
    End If
    rst.Close
End Function

Function GETMOINNAME(KOL, MOIN As Double) As String
        On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT DETA_HES.N_KOL, DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.TOZIH, DETA_HES.BED_BES, DETA_HES.ADDRESS, DETA_HES.TEL, DETA_HES.CODE_E FROM DETA_HES WHERE (((DETA_HES.N_KOL)=" & KOL & ") AND ((DETA_HES.NUMBER)=" & MOIN & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(2)) Then
        GETMOINNAME = ""
    Else
        GETMOINNAME = RRST.Fields(2)
    End If
  Else
    GETMOINNAME = ""
  End If
  RRST.Close
End Function
Function GETKOLNAME(KOL As Double) As String
        On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT TOTA_HES.NAME FROM TOTA_HES WHERE (TOTA_HES.NUMBER=" & KOL & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(0)) Then
        GETKOLNAME = ""
    Else
        GETKOLNAME = RRST.Fields(0)
    End If
  Else
    GETKOLNAME = ""
  End If
  RRST.Close
End Function
Function GETHESNAME(HES As Variant) As String
         On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT CUST_HESAB.NAME FROM CUST_HESAB WHERE (CUST_HESAB.HES='" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(0)) Then
        GETHESNAME = ""
    Else
        GETHESNAME = RRST.Fields(0)
    End If
  Else
    GETHESNAME = ""
  End If
  RRST.Close
End Function
Function GETHESMOBIL(HES As Variant) As String
         On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT CUST_HESAB.MOBILE FROM CUST_HESAB WHERE (CUST_HESAB.HES='" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(0)) Then
        GETHESMOBIL = ""
    Else
        GETHESMOBIL = RRST.Fields(0)
    End If
  Else
    GETHESMOBIL = ""
  End If
  RRST.Close
End Function
Function GETTAFNAME(HES As String) As String
        On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT     NAME FROM dbo.CUST_HESAB WHERE     (hes = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(0)) Then
        GETTAFNAME = " "
    Else
        GETTAFNAME = RRST.Fields(0)
    End If
  Else
    GETTAFNAME = " "
  End If
  RRST.Close
End Function
Function GETTAFNAME2(KOL, MOIN, taf, TAF2 As Double) As String
        On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT TDETA_HES2.N_KOL, TDETA_HES2.NUMBER, TDETA_HES2.TNUMBER, TDETA_HES2.TNUMBER2, TDETA_HES2.NAME, TDETA_HES2.TOZIH, TDETA_HES2.BED_BES, TDETA_HES2.ADDRESS, TDETA_HES2.TEL, TDETA_HES2.CODE_E FROM TDETA_HES2 WHERE (((TDETA_HES2.N_KOL)=" & KOL & ") AND ((TDETA_HES2.NUMBER)=" & MOIN & ") AND ((TDETA_HES2.TNUMBER)=" & taf & ") AND ((TDETA_HES2.TNUMBER2)=" & TAF2 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields(4)) Then
        GETTAFNAME2 = " "
    Else
        GETTAFNAME2 = RRST.Fields(4)
    End If
  Else
    GETTAFNAME2 = " "
  End If
  RRST.Close
End Function
Function GETBANK(BANK As Double) As String
        On Error Resume Next
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS WHERE (((TCOD_BANKS.CODE)= " & BANK & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    GETBANK = RRST.Fields(1)
  Else
    GETBANK = ""
  End If
  RRST.Close
End Function

Public Function GETKOL(SHES As String) As Long
                 On Error Resume Next
    Dim i As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
        While MID(SHES, i, 1) <> "-"
            i = i + 1
            If i > 200 Then
                Exit Function
            End If
        Wend
        GETKOL = Left(SHES, i - 1)
    End If
End Function
Public Function GETMOIN(SHES As String) As Long
    Dim i, j As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
    While MID(SHES, i, 1) <> "-"
        i = i + 1
            If i > 200 Then
                Exit Function
            End If
    Wend
    j = i + 1
    While MID(SHES, j, 1) <> "-" And j <= Len(SHES)
        j = j + 1
    Wend
    i = i + 1
    GETMOIN = MID(SHES, i, j - i)
    End If
End Function
Public Function GETTAF(SHES As String) As Long
    Dim i, j As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
        While MID(SHES, i, 1) <> "-"
            i = i + 1
            If i > 200 Then
                Exit Function
            End If
        Wend
        j = i + 1
        While MID(SHES, j, 1) <> "-"
            j = j + 1
        Wend
        i = j
        While MID(SHES, j + 1, 1) <> "-" And j < Len(SHES)
            j = j + 1
        Wend
        GETTAF = MID(SHES, i + 1, j - i)
    End If
End Function
Public Function GETTAF2(SHES As String) As Long
    Dim i, j As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
        While MID(SHES, i, 1) <> "-"
            i = i + 1
        Wend
        j = i + 1
        While MID(SHES, j, 1) <> "-"
            j = j + 1
        Wend
        i = j
        While MID(SHES, j + 1, 1) <> "-" And j < Len(SHES)
            j = j + 1
        Wend
        If j = Len(SHES) Then
             GETTAF2 = -1
        Else
            GETTAF2 = MID(SHES, j + 2, Len(SHES) - (j + 1))
        End If
    End If
End Function
Public Function GETTAF3(SHES As String, ByRef KOL, ByRef MOIN, ByRef taf, ByRef TAF2, ByRef TAF3, ByRef TAF4)
                     On Error Resume Next
    Dim i1, I2, I3, I4, I5, I6, I7, I8, j, K As Byte
    i1 = 0: I2 = 0: I3 = 0: I4 = 0: I5 = 0: I6 = 0: I7 = 0: I8 = 0
    j = 1
    K = 1
    If (Len(SHES) > 4) Then
        For j = 1 To Len(SHES)
            If MID(SHES, j, 1) = "-" Then
                Select Case K
                    Case 1: i1 = j
                    Case 2: I2 = j
                    Case 3: I3 = j + 1
                    Case 4: I4 = j + 1
                    Case 5: I5 = j + 1
                    Case 6: I6 = j + 1
                    Case 7: I7 = j + 1
                    Case 8: I8 = j + 1
                End Select
                K = K + 1
            End If
        Next j
    End If
    KOL = Null
    MOIN = Null
    taf = Null
    TAF2 = Null
    TAF3 = Null
    TAF4 = Null
    If i1 = 0 Then
        KOL = SHES
        Exit Function
    Else
        KOL = MID(SHES, 1, i1 - 1)
    End If
    If I2 = 0 Then
        MOIN = MID(SHES, i1 + 1, Len(SHES) - i1)
        Exit Function
    Else
        MOIN = MID(SHES, i1 + 1, I2 - i1 - 1)
    End If
    If I3 = 0 Then
        taf = MID(SHES, I2 + 1, Len(SHES) - I2)
        Exit Function
    Else
        taf = MID(SHES, I2 + 1, I3 - I2 - 2)
    End If
    If I4 = 0 Then
        TAF2 = MID(SHES, I3, Len(SHES) - I3 + 1)
    Else
        TAF2 = MID(SHES, I3, I4 - I3 - 1)
        If I5 = 0 Then
            TAF3 = MID(SHES, I4, Len(SHES) - I4 + 1)
        Else
            TAF3 = MID(SHES, I4, I5 - I4 - 1)
            If I6 = 0 Then
                TAF4 = MID(SHES, I5, Len(SHES) - I5 + 1)
            Else
                TAF4 = MID(SHES, I5, I6 - I5 - 1)
            End If
        End If
    End If
End Function
Public Function ISHESAB(KOL, MOIN, taf As Long) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT N_KOL,NUMBER,TNUMBER FROM TDETA_HES WHERE N_KOL = " & KOL & " AND NUMBER = " & MOIN & " AND TNUMBER = " & taf, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      ISHESAB = False
    Else
      ISHESAB = True
    End If

End Function
Public Function ISHESAB3(HES As String) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT HES FROM CUST_HESAB WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      ISHESAB3 = False
    Else
      ISHESAB3 = True
    End If
End Function


Public Sub CREATHES(KOL, MOIN, taf As Long, nam As String)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & KOL & " AND NUMBER = " & MOIN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_KOL") = KOL
        rst.Fields("NUMBER") = MOIN
        rst.Fields("NAME") = nam
        rst.update
    End If
    Err.Clear
    Set rst = New ADODB.Recordset
    rst.Open "TDETA_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("N_KOL") = KOL
    rst.Fields("NUMBER") = MOIN
    rst.Fields("TNUMBER") = taf
    rst.Fields("NAME") = nam
    rst.update
    If Err <> 0 Then
        rst.Fields("NAME") = nam & " " & taf
        rst.update
    End If

End Sub

Public Sub CREATHESCO(HES As String, nam As String)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
       On Error Resume Next
    Call GETTAF3(HES, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
    If Not ISHESAB3(HES) Then
        If IsNull(CTAF2) Then
            rst.Open "TDETA_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
            rst.Fields("N_KOL") = CKOL
            rst.Fields("NUMBER") = CMOIN
            rst.Fields("TNUMBER") = CTAF
            rst.Fields("NAME") = nam
            rst.update
         Else
             If IsNull(CTAF3) Then
                rst.Open "TDETA_HES2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("N_KOL") = CKOL
                rst.Fields("NUMBER") = CMOIN
                rst.Fields("TNUMBER") = CTAF
                rst.Fields("TNUMBER2") = CTAF2
                rst.Fields("NAME") = nam
                rst.update
             Else
                 If IsNull(CTAF4) Then
                    rst.Open "TDETA_HES3", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("N_KOL") = CKOL
                    rst.Fields("NUMBER") = CMOIN
                    rst.Fields("TNUMBER") = CTAF
                    rst.Fields("TNUMBER2") = CTAF2
                    rst.Fields("TNUMBER3") = CTAF3
                    rst.Fields("NAME") = nam
                    rst.update
                 Else
                    rst.Open "TDETA_HES4", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("N_KOL") = CKOL
                    rst.Fields("NUMBER") = CMOIN
                    rst.Fields("TNUMBER") = CTAF
                    rst.Fields("TNUMBER2") = CTAF2
                    rst.Fields("TNUMBER3") = CTAF3
                    rst.Fields("TNUMBER4") = CTAF4
                    rst.Fields("NAME") = nam
                    rst.update
                 End If
             End If
        End If
    End If

End Sub


Public Sub CREATHES4(KOL, MOIN, taf, TAF2 As Long, nam As String)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & KOL & " AND NUMBER = " & MOIN & " AND TNUMBER = " & TAF2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_KOL") = KOL
        rst.Fields("NUMBER") = MOIN
        rst.Fields("TNUMBER") = taf
        rst.Fields("NAME") = nam
        rst.update
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "TDETA_HES2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("N_KOL") = KOL
    rst.Fields("NUMBER") = MOIN
    rst.Fields("TNUMBER") = taf
    rst.Fields("TNUMBER2") = TAF2
    rst.Fields("NAME") = nam
    rst.update
End Sub


Public Sub SANADKHAD(NUMBER, NUMBER2 As Long)
                     On Error Resume Next
   Dim CON, i As Long
   DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
   DoCmd.OpenForm "GUG"
   Forms![GUG]![Text0] = "n"
   Forms![GUG]![Text0].Visible = True
   Forms![GUG]![Box7].Visible = True
   Forms![GUG].Form.Refresh
   Forms![GUG].Form.Repaint
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, HEDRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, takh As Double, SHART As String
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.Open "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG=14) AND (NUMBER >=" & NUMBER & ") AND (NUMBER <=" & NUMBER2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  Forms![GUG]![SNUM] = HEDRST.RecordCount
  CON = HEDRST.RecordCount / 50
  If CON < 1 Then
    CON = 1
  End If
  i = 1
  While Not HEDRST.EOF
      Forms![GUG]![num] = i
      If Not IsNull(HEDRST.Fields("CUST_NO")) Then
           Call GETTAF3(HEDRST.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
'         rst.Close
         Set rst = New ADODB.Recordset
         rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If IsNull(rst.Fields(0)) Then
           max_ns = 1
         Else
           max_ns = rst.Fields(0) + 1
         End If
         Err.Clear
         SHART = "NO_S = 14 AND N_S = " & HEDRST.Fields("N_S")
         SHRST.Filter = SHART: Set SHRST = SHRST
         If SHRST.RecordCount = 0 Or Err <> 0 Then
           Err.Clear
           SHRST.AddNew
           SHRST.Fields("N_S") = max_ns
           SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
           SHRST.Fields("SHARH_S") = RIGHT(" فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " خريدار: " & GETTAFNAME(HEDRST.Fields("CUST_NO")), 255)
           SHRST.Fields("GHATEI") = 0
           SHRST.Fields("NO_S") = 14
           SHRST.Fields("OKF") = True
           SHRST.Fields("CRT") = Now()
           SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
           SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
           SHRST.update
         Else
              max_ns = SHRST.Fields("N_S")
         End If
         If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
            HEDRST.Fields("N_S") = max_ns
         End If
    '     JST.Close
         Set JST = New ADODB.Recordset
          JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & HEDRST.Fields("NUMBER") & " ) AND ((INVO_LST.TAG)=14))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
               JAMF = JST.Fields(0)
          Else
             JAMF = 0
          End If
     '     JST.Close
          Set JST = New ADODB.Recordset
          JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=14) AND ((PAY_GETD.NUMBER)= " & HEDRST.Fields("NUMBER") & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
               JAMCH = JST.Fields(0)
          Else
             JAMCH = 0
          End If
      '    JST.Close
         DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 14))")
       Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         If JAMF <> 0 Then
           SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##") & HEDRST.Fields("MOLAH"), 255)
           SDRST.Fields("BED") = JAMF + HEDRST.Fields("MABL_HAZ") + HEDRST.Fields("MBAA")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.Fields("RADIF") = HEDRST.Fields("NUMBER")
          SDRST.update
          End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      Set JST = New ADODB.Recordset
      JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & HEDRST.Fields("NUMBER") & ") AND ((INVO_LST.TAG)=14)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          While Not JST.EOF()
                If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(HEDRST.Fields("DEPATMAN")), CLng(JST.Fields("CODE"))) Then
                  Err.Clear
                  Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(HEDRST.Fields("DEPATMAN")), CLng(JST.Fields("CODE")), JST.Fields(4))
                  If Err <> 0 Then
                     DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر خدمات در درآمد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                     End If
               End If
              SDRST.AddNew 'در آمد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
               SDRST.Fields("HES_M") = HEDRST.Fields("DEPATMAN")
               SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
               SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & HEDRST.Fields("DEPATMAN") & "-" & CLng(JST.Fields("code"))
               SDRST.Fields("SHARH") = Left("فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خدمات " & Trim(JST.Fields("NAME")), 255)
               SDRST.Fields("BES") = JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 14
               SDRST.update
               JST.MoveNext
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         Wend
         If HEDRST.Fields("MABL_HAZ") <> 0 Then
           SDRST.AddNew 'كرايه حمل ياخدمات
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = GETKOL(HEDRST.Fields("MOIN_HAZ"))
           SDRST.Fields("HES_M") = GETMOIN(HEDRST.Fields("MOIN_HAZ"))
           SDRST.Fields("HES_T") = GETTAF(HEDRST.Fields("MOIN_HAZ"))
           SDRST.Fields("hes") = GETKOL(HEDRST.Fields("MOIN_HAZ")) & "-" & GETMOIN(HEDRST.Fields("MOIN_HAZ")) & "-" & GETTAF(HEDRST.Fields("MOIN_HAZ"))
           SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " - " & GETTAFNAME(HEDRST.Fields("MOIN_HAZ")), 255)
           SDRST.Fields("BES") = HEDRST.Fields("MABL_HAZ")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         End If
         If JAMCH <> 0 Then 'چكهاي دريافتي
           Set CHRST = New ADODB.Recordset
           CHRST.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           CHRST.MoveLast
           CHRST.MoveFirst
           CHRST.Filter = "NUMBER = " & HEDRST.Fields("NUMBER") & " AND TAG = " & 14
           If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
           Do Until CHRST.EOF
               MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
               SDRST.AddNew 'اسناد دريافتني
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![ADA])
               SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![ADA])
               SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![ADA])
               SDRST.Fields("hes") = Forms![BASEKNOW]![ADA]
               SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
               SDRST.Fields("BED") = CHRST.Fields("MABL")
               SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
               SDRST.Fields("BANK") = CHRST.Fields("BANK")
               SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 14
               SDRST.update
               SDRST.AddNew 'چكهاي دريافتي
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = CKOL
               SDRST.Fields("HES_M") = CMOIN
               SDRST.Fields("HES_T") = CTAF
               SDRST.Fields("HES_T2") = CTAF2
               SDRST.Fields("HES_T3") = CTAF3
               SDRST.Fields("HES_T4") = CTAF4
               SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
               SDRST.Fields("SHARH") = RIGHT("ف.خ." & HEDRST.Fields("NUMBER") & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
               SDRST.Fields("BES") = CHRST.Fields("MABL")
               SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 14
               SDRST.update
               CHRST.MoveNext
           Loop
           Else
           End If
         End If
       Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
        If HEDRST.Fields("M_NAGHD") <> 0 Then
           SDRST.AddNew 'مبلغ نقدشخص
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BES") = HEDRST.Fields("M_NAGHD")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         If HEDRST.Fields("M_NAGHD") <> 0 Then
           SDRST.AddNew 'مبلغ نقدصندوق
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
           SDRST.Fields("HES_M") = HEDRST.Fields("DEPATMAN")
           SDRST.Fields("HES_T") = HEDRST.Fields("SHIFT")
           SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & HEDRST.Fields("DEPATMAN") & "-" & HEDRST.Fields("SHIFT")
           SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BED") = HEDRST.Fields("M_NAGHD")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
            takh = 0
     If Forms![BASEKNOW]![TKHF] = 1 Then
           SDRST.AddNew 'تخفيف خدمات
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
           SDRST.Fields("HES_M") = 1
           SDRST.Fields("HES_T") = 1
           SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-1-1"
           SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 60)
           SDRST.Fields("BED") = HEDRST.Fields("TAKHFIF")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
     Else
        If Forms![BASEKNOW]![TKHF] = 3 Then
          Set rst = New ADODB.Recordset
          rst.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.HEAD_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") AND (dbo.HEAD_LST.TAG = 14)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
            takh = 0
            While Not rst.EOF()
               If Round(rst.Fields("N_MOIN")) <> 0 Then
               If Not ISHESAB(CLng(Forms![BASEKNOW]![HDARAM]), CLng(HEDRST.Fields("CUST_KIND")), CLng(rst.Fields("CODE"))) Then
                 Err.Clear
                 Call CREATHES(CLng(Forms![BASEKNOW]![HDARAM]), CLng(HEDRST.Fields("CUST_KIND")), CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                 If Err <> 0 Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                 End If
                End If
                SDRST.AddNew 'تخفيف فروش
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
                SDRST.Fields("HES_M") = HEDRST.Fields("DEPATMAN")
                SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
                SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-" & HEDRST.Fields("DEPATMAN") & "-" & CLng(rst.Fields("CODE"))
                SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 14
                takh = takh + Round(rst.Fields("N_MOIN"))
                SDRST.update
              End If
                rst.MoveNext
             Wend
          End If
          If HEDRST.Fields("TAKHFIF") <> takh Then
             HEDRST.Fields("TAKHFIF") = takh
          End If
        Else
           DoCmd.RunSQL "UPDATE    dbo.INVO_LST SET  N_KOL = 0, N_MOIN = 0 WHERE (NUMBER = " & HEDRST.Fields("NUMBER") & " ) AND (TAG = 14)"
           Set rst = New ADODB.Recordset
           rst.Open "SELECT dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.TAKHPERS.CUST_CO, dbo.TAKHPERS.TAKH_COD, dbo.TAKHPERS.TAFPER, dbo.INVO_LST.MABL_K, dbo.STUF_DEF.NAME,  dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN   FROM  dbo.INVO_LST INNER JOIN dbo.TAKHPERS ON dbo.INVO_LST.CODE = dbo.TAKHPERS.TAKH_COD INNER JOIN  dbo.STUF_DEF ON dbo.TAKHPERS.TAKH_COD = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") And (dbo.INVO_LST.TAG = 14) And (dbo.TAKHPERS.CUST_CO = " & HEDRST.Fields("CUST_KIND") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
             takh = 0
             While Not rst.EOF()
                 rst.Fields("N_KOL") = rst.Fields("TAFPER")
                 rst.Fields("N_MOIN") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                 rst.update
                 If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![HDARAM]), CLng(HEDRST.Fields("DEPATMAN")), rst.Fields(3)) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![HDARAM]), CLng(HEDRST.Fields("DEPATMAN")), rst.Fields(3), "تخفيف " & rst.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات درآمد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                          End If
                       End If
                    SDRST.AddNew 'تخفيف خدمات
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
                    SDRST.Fields("HES_M") = HEDRST.Fields("DEPATMAN")
                    SDRST.Fields("HES_T") = rst.Fields(3)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-" & HEDRST.Fields("DEPATMAN") & "-" & rst.Fields(3)
                    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 60)
                    SDRST.Fields("BED") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                    SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 14
                    takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                    SDRST.update
                End If
                rst.MoveNext
              Wend
           End If
           If HEDRST.Fields("TAKHFIF") <> takh Then
             HEDRST.Fields("TAKHFIF") = takh
           End If
       End If
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         If HEDRST.Fields("TAKHFIF") <> 0 Then
           SDRST.AddNew 'مبلغ تخفيف شخص
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BES") = HEDRST.Fields("TAKHFIF")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If

      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         If HEDRST.Fields("MABL_HAV") <> 0 Then
           SDRST.AddNew 'مبلغ حواله شخص
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BES") = HEDRST.Fields("MABL_HAV")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
       Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
        If HEDRST.Fields("MABL_HAV") <> 0 Then
           SDRST.AddNew 'مبلغ حواله
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = GETKOL(HEDRST.Fields("MOIN_HAV"))
           SDRST.Fields("HES_M") = GETMOIN(HEDRST.Fields("MOIN_HAV"))
           SDRST.Fields("HES_T") = GETTAF(HEDRST.Fields("MOIN_HAV"))
           SDRST.Fields("hes") = GETKOL(HEDRST.Fields("MOIN_HAV")) & "-" & GETMOIN(HEDRST.Fields("MOIN_HAV")) & "-" & GETTAF(HEDRST.Fields("MOIN_HAV"))
           SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BED") = HEDRST.Fields("MABL_HAV")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
         If HEDRST.Fields("MABL_VAR") <> 0 Then
           SDRST.AddNew 'مبلغ واريزي شخص
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HEDRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BES") = HEDRST.Fields("MABL_VAR")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
       Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
        If HEDRST.Fields("MABL_VAR") <> 0 Then
           SDRST.AddNew 'مبلغ حواله
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = GETKOL(HEDRST.Fields("MOIN_VAR"))
           SDRST.Fields("HES_M") = GETMOIN(HEDRST.Fields("MOIN_VAR"))
           SDRST.Fields("HES_T") = GETTAF(HEDRST.Fields("MOIN_VAR"))
           SDRST.Fields("hes") = HEDRST.Fields("MOIN_VAR")
           SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خدمات شماره " & HEDRST.Fields("NUMBER") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BED") = HEDRST.Fields("MABL_VAR")
           SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 14
           SDRST.update
         End If
      If HEDRST.Fields("MBAA") <> 0 Then
        SDRST.AddNew 'ماليات بر ارزش افزوده
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HEDRST.Fields("HMBAA")) Then
             Call GETTAF3(HEDRST.Fields("HMBAA"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HEDRST.Fields("HMBAA")
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور خدمات شماره " & HEDRST.Fields("NUMBER1") & " مورخ" & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BES") = HEDRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 14
        SDRST.Fields("ARZD") = HEDRST.Fields("ARZD")
        SDRST.update
      End If
       HEDRST.MoveNext
      rst.Close
      Set rst = New ADODB.Recordset
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Refresh
        Forms![GUG]![Text0].Requery
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"

End Sub
Public Sub SANADDP(IDNUM As Long)
                     On Error Resume Next
  Dim rst, SHRST As New ADODB.Recordset, HEDRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK As Double, SHART, SHRH As String
  HEDRST.Open "SELECT PGET_HED.* FROM PGET_HED WHERE (((PGET_HED.ID)=" & IDNUM & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If HEDRST.RecordCount > 0 Then
        rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Then
          max_ns = 1
        Else
          max_ns = rst.Fields(0) + 1
        End If
        SHRH = "خزانه داري شماره " & HEDRST.Fields("ID") & "مورخ " & Format(HEDRST.Fields("DATE"), "####/##/##")
        SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(HEDRST.Fields("N_S")) Then
              SHRST.AddNew
              SHRST.Fields("N_S") = max_ns
              SHRST.Fields("DATE_S") = HEDRST.Fields("DATE")
              SHRST.Fields("SHARH_S") = SHRH
              SHRST.Fields("GHATEI") = 0
              SHRST.Fields("NO_S") = 5
              SHRST.Fields("USER_NAME") = UCurrentUser()
              SHRST.Fields("CRT") = Now()
              SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
              SHRST.update
              HEDRST.Fields("N_S") = max_ns
              HEDRST.update
        Else
            SHART = "N_S = " & HEDRST.Fields("N_S")
            SHRST.Filter = SHART
            Set SHRST = SHRST
            If SHRST.RecordCount = 0 Then
              SHRST.AddNew
              SHRST.Fields("N_S") = max_ns
              SHRST.Fields("DATE_S") = HEDRST.Fields("DATE")
              SHRST.Fields("SHARH_S") = SHRH
              SHRST.Fields("GHATEI") = 0
              SHRST.Fields("NO_S") = 5
              SHRST.Fields("USER_NAME") = UCurrentUser()
              SHRST.Fields("CRT") = Now()
              SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
              SHRST.update
            Else

              SHRST.Fields("DATE_S") = HEDRST.Fields("DATE")
              SHRST.Fields("SHARH_S") = SHRH
              SHRST.Fields("USER_NAME") = UCurrentUser()
              SHRST.update
            End If
        End If
        If Not IsNull(HEDRST.Fields("N_S")) Then
          DoCmd.RunSQL ("DELETE DEED_DTL.N_S FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & HEDRST.Fields("N_S") & " ))")
        End If
        Forms![ENTEGHAL_DARYAFT_PARDAKHT]![SND] = HEDRST.Fields("N_S")
        DoCmd.OpenQuery "OPGET_TSANAD"
        DoCmd.OpenQuery "OPGET_FSANAD"
        rst.Close
        SHRST.Close
  End If
End Sub
Public Sub GENSANADFROOSH(fnum, TNUM As Long)
                  On Error Resume Next
    Dim CON, i As Long
    DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
    DoCmd.OpenForm "GUG"
    Forms![GUG]![Text0] = "n"
    Forms![GUG]![Text0].Visible = True
    Forms![GUG]![Box7].Visible = True
    Forms![GUG].Form.Refresh
    Forms![GUG].Form.Repaint
    Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SARST  As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, PRST As New ADODB.Recordset
    Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, takh As Double, SHART As String, MAVAD As Double, DAST As Double, SAR As Double
    Dim ii As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long, SHSH As String
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Set SDRST = Forms![recording].Form.RecordsetClone
     HFRST.Open "SELECT     * FROM dbo.HEAD_LST WHERE     (NUMBER BETWEEN " & fnum & " AND " & TNUM & ") AND (TAG = 13) ORDER BY NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'    HFRST.MoveLast
'    HFRST.MoveFirst
    Forms![GUG]![SNUM] = HFRST.RecordCount
    CON = HFRST.RecordCount / 50
    If CON < 1 Then
      CON = 1
    End If
    i = 1
    While Not HFRST.EOF
      Forms![GUG]![num] = i
      If Not IsNull(HFRST.Fields("CUST_NO")) Then
           Call GETTAF3(HFRST.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
      rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If IsNull(rst.Fields(0)) Then
        max_ns = 1
      Else
        max_ns = rst.Fields(0) + 1
      End If
      SHSH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي  فروش  " & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255), Left(" فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " خريدار: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255))
      If Forms![BASEKNOW]![SNDKH] Then 'سند روزانه است
        If Not IsNull(HFRST.Fields("N_S")) Then 'فاکتور سند دارد
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT   BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and n_s = " & HFRST.Fields("N_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then  'اگرسند  فاکتورهست
                If SARST.Fields("DATE_S") = HFRST.Fields("DATE_N") Then 'تاريخ سند و فاکتوريکي است
                    max_ns = HFRST.Fields("N_S")
                Else
SEJ:
                    Set SARST = New ADODB.Recordset: SARST.Open "SELECT    BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and DATE_s = " & HFRST.Fields("DATE_N"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If SARST.RecordCount > 0 Then   'اگرسند به تاريخ فاکتورهست
                        max_ns = SARST.Fields("N_S")
                    Else
                        max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 2, -1, HFRST.Fields("USER_NAME"))
                        HFRST.Fields("N_S") = max_ns
                    End If
                End If
            Else
                GoTo SEJ ' چک کن اگه نيست صادر کن
            End If
        Else
           GoTo SEJ ' چک کن اگه نيست صادر کن
        End If
      Else 'تک سندي
        If Not IsNull(HFRST.Fields("N_S")) Then 'فاکتور سند دارد
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT    n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and N_s = " & HFRST.Fields("N_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then   'اگرسند فاکتورهست
                If SARST.Fields("DATE_S") <> HFRST.Fields("DATE_N") Then 'تاريخ سند و فاکتوريکي است
                    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & HFRST.Fields("DATE_N") & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 2,OKF=-1,USER_NAME ='" & HFRST.Fields("USER_NAME") & "' WHERE N_S =" & HFRST.Fields("N_S"))
                End If
                max_ns = HFRST.Fields("N_S")
            Else
                max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 2, -1, HFRST.Fields("USER_NAME"))
                HFRST.Fields("N_S") = max_ns
            End If
        Else
            max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 2, -1, HFRST.Fields("USER_NAME"))
            HFRST.Fields("N_S") = max_ns
        End If
      End If

      If IsNull(HFRST.Fields("N_S")) Or HFRST.Fields("N_S") <> max_ns Then
         HFRST.Fields("N_S") = max_ns
         HFRST.update
      End If
       JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & HFRST.Fields("NUMBER") & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          JAMF = JST.Fields(0)
       Else
          JAMF = 0
       End If
       Set JST = New ADODB.Recordset
       JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=2) AND ((PAY_GETD.NUMBER)= " & HFRST.Fields("NUMBER") & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          JAMCH = JST.Fields(0)
       Else
          JAMCH = 0
       End If
       Set JST = New ADODB.Recordset
       DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HFRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 13))")
       If JAMF + HFRST.Fields("MABL_HAZ") + HFRST.Fields("MBAA") - HFRST.Fields("TAKHFIF") <> 0 Then
        SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("ف ف ش " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##") & HFRST.Fields("MOLAH") & IIf(MID(Forms![BASEKNOW]![OPTIONSS], 55, 1) = "5", " - " & GETDEPART(HFRST.Fields("DEPATMAN")), " "), 255)
        SDRST.Fields("BED") = Round(JAMF + HFRST.Fields("MABL_HAZ") + HFRST.Fields("MBAA") - HFRST.Fields("TAKHFIF"))
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.Fields("RADIF") = HFRST.Fields("NUMBER")
        SDRST.update
       End If
       Set JST = New ADODB.Recordset
       JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE     (dbo.INVO_LST.NUMBER = " & HFRST.Fields("NUMBER") & ") AND (dbo.INVO_LST.TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       While Not JST.EOF()
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                   If HFRST.Fields("sader") = 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), 1, CLng(JST.Fields("code"))) Then
                             Err.Clear
                             Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), 1, CLng(JST.Fields("code")), JST.Fields(4))
                             If Err <> 0 Then
                                DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                             End If
                         End If
                         If JST.Fields("MABL_K") > 0 Then
                             SDRST.AddNew 'فروش
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
                             SDRST.Fields("HES_M") = 1
                             SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-1-" & CLng(JST.Fields("code"))
                             SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                             SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                             SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                             SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                             SDRST.Fields("TAG") = 13
                             SDRST.update
                          End If
                    Else
                        If JST.Fields("anbar") <> 0 Then
                            If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), 2, CLng(JST.Fields("code"))) Then
                                 Err.Clear
                                 Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), 2, CLng(JST.Fields("code")), JST.Fields(4))
                                 If Err <> 0 Then
                                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                                 End If
                             End If
                             If JST.Fields("MABL_K") > 0 Then
                                 SDRST.AddNew 'فروش
                                 SDRST.Fields("N_S") = max_ns
                                 SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
                                 SDRST.Fields("HES_M") = 2
                                 SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                                 SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-2-" & CLng(JST.Fields("code"))
                                 SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                                 SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                                 SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                                 SDRST.Fields("TAG") = 13
                                 SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                                 SDRST.update
                              End If
                        Else
                            If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code"))) Then
                               Err.Clear
                               Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code")), JST.Fields(4))
                               If Err <> 0 Then
                                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                               End If
                            End If
                            SDRST.AddNew 'در آمد
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
                            SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
                            SDRST.Fields("HES_T") = JST.Fields("CODE")
                            SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & HFRST.Fields("DEPATMAN") & "-" & JST.Fields("CODE")
                            SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                            SDRST.Fields("TAG") = 13
                            SDRST.update
                        End If
                    End If
                 Else
                    If JST.Fields("anbar") <> 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                           Err.Clear
                           Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                           If Err <> 0 Then
                              DoCmd.OpenForm "MESAG", , , , , acDialog, Forms![BASEKNOW]![FROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code")) & "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                           End If
                        End If
                        If JST.Fields("MABL_K") > 0 Then
                            SDRST.AddNew 'فروش
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                            SDRST.Fields("TAG") = 13
                            SDRST.update
                         End If
                    Else
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code"))) Then
                           Err.Clear
                           Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code")), JST.Fields(4))
                           If Err <> 0 Then
                              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                           End If
                        End If
                        SDRST.AddNew 'در آمد
                        SDRST.Fields("N_S") = max_ns
                        SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
                        SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
                        SDRST.Fields("HES_T") = JST.Fields("CODE")
                        SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & HFRST.Fields("DEPATMAN") & "-" & JST.Fields("CODE")
                        SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                        SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                        SDRST.Fields("TAG") = 13
                        SDRST.update
                    End If
                End If
            JST.MoveNext
      Wend
      Set JST = New ADODB.Recordset
     If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
            JST.Open "SELECT     dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.STUF_DEF.NAME,  dbo.INVO_LST.AVRAGE FROM  dbo.INVO_LST INNER JOIN  dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & HFRST.Fields("NUMBER") & ") And (dbo.INVO_LST.TAG = 2) And (dbo.INVO_LST.ANBAR <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not JST.EOF()
                MAVAD = Round(GETSTANDARDPRICE_MAVAD(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
                DAST = Round(GETSTANDARDPRICE_DAST(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
                SAR = Round(GETSTANDARDPRICE_SAR(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
              If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                 Err.Clear
                 Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(4))
                 If Err <> 0 Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "كد كالا : " & JST.Fields("code")
                 End If
              End If
              If (MAVAD + DAST + SAR) <> 0 And MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) <> "5" Then
                  SDRST.AddNew 'انبار محصول
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                  SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                  SDRST.Fields("BES") = (MAVAD + DAST + SAR)
                  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                  SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                  SDRST.Fields("TAG") = 13
                  SDRST.update
                If MID(Forms![BASEKNOW]![tindata], 9, 1) <> 1 Then
                  If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                     End If
                  End If
                End If
                  If MAVAD > 0 Then
                      SDRST.AddNew 'قيمت تمام شده
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                    If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                      SDRST.Fields("HES_M") = 1
                      SDRST.Fields("HES_T") = 1
                      SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                    Else
                      SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                      SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                      SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                    End If
                      SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("BED") = MAVAD
                      SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                      SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                      SDRST.Fields("TAG") = 13
                      SDRST.update
                  End If
                If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999, "دستمزد " & JST.Fields(4))
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                   End If
                End If
                If DAST <> 0 Then
                      SDRST.AddNew 'قيمت تمام شده
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                      SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                      SDRST.Fields("HES_T") = 9999999
                      SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                      SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                          SDRST.Fields("BED") = DAST
                      SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                      SDRST.Fields("TAG") = 13
                      SDRST.update

                      If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998) Then
                         Err.Clear
                         Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998, "سربار " & JST.Fields(4))
                         If Err <> 0 Then
                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & JST.Fields(4) & " - " & JST.Fields("code")
                         End If
                      End If
                End If
                If SAR <> 0 Then
                       SDRST.AddNew 'قيمت تمام شده
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                       SDRST.Fields("HES_T") = 9999998
                       SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                           SDRST.Fields("BED") = SAR
                       SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 13
                       SDRST.update
                End If
            Else
                  If JST.Fields("avrage") > 0 Then
                    SDRST.AddNew 'انبار محصول
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                    SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                    SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                    SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                    SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                    SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 13
                    SDRST.update

                    SDRST.AddNew 'قيمت تمام شده
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                    If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                      SDRST.Fields("HES_M") = 1
                      SDRST.Fields("HES_T") = 1
                      SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                    Else
                      SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                      SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                      SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                    End If
                      SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                      SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                      SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                      SDRST.Fields("TAG") = 13
                      SDRST.update
                  End If
            End If
            JST.MoveNext
         Wend
      End If
      Set JST = New ADODB.Recordset
      If HFRST.Fields("MABL_HAZ") <> 0 Then
        If IsNull(HFRST.Fields("MOIN_HAZ")) Then
               DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد." & " فاكتور فروش شماره " & HFRST.Fields("NUMBER")
        Else
            SDRST.AddNew 'كرايه حمل ياخدمات
            SDRST.Fields("N_S") = max_ns
            If Not IsNull(HFRST.Fields("MOIN_HAZ")) Then
                 Call GETTAF3(HFRST.Fields("MOIN_HAZ"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            End If
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = HFRST.Fields("MOIN_HAZ")
            SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " - " & GETTAFNAME(HFRST.Fields("MOIN_HAZ")), 255)
            SDRST.Fields("BES") = Round(HFRST.Fields("MABL_HAZ"))
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.Fields("TAG") = 13
            SDRST.update
        End If
      End If
      If JAMCH <> 0 Then 'چكهاي دريافتي
        CHRST.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        CHRST.MoveLast
        CHRST.MoveFirst
        CHRST.Filter = "NUMBER = " & HFRST.Fields("NUMBER") & " AND TAG = 2"
        Set CHRST = CHRST
        If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
        Do Until CHRST.EOF
            MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
            SDRST.AddNew 'اسناد دريافتني
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![ADA])
            SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![ADA])
            SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![ADA])
            SDRST.Fields("hes") = Forms![BASEKNOW]![ADA]
            SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
            SDRST.Fields("BED") = CHRST.Fields("MABL")
            SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
            SDRST.Fields("BANK") = CHRST.Fields("BANK")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.Fields("TAG") = 13
            SDRST.update
            SDRST.AddNew 'چكهاي دريافتي
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = CKOL
            SDRST.Fields("HES_M") = CMOIN
            SDRST.Fields("HES_T") = CTAF
            SDRST.Fields("HES_T") = CTAF
            SDRST.Fields("HES_T2") = CTAF2
            SDRST.Fields("HES_T3") = CTAF3
            SDRST.Fields("HES_T4") = CTAF4
            SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
            SDRST.Fields("SHARH") = RIGHT("ف.ف." & HFRST.Fields("NUMBER1") & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
            SDRST.Fields("BES") = CHRST.Fields("MABL")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.Fields("TAG") = 13
            SDRST.update
            CHRST.MoveNext
        Loop
        Else
        End If
        CHRST.Close
        Set CHRST = New ADODB.Recordset
      End If
      If HFRST.Fields("M_NAGHD") <> 0 Then
        SDRST.AddNew 'مبلغ نقدشخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        If HFRST.Fields("M_NAGHD") > 0 Then
            SDRST.Fields("BES") = HFRST.Fields("M_NAGHD")
        Else
            SDRST.Fields("BEd") = Abs(HFRST.Fields("M_NAGHD"))
        End If
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If HFRST.Fields("M_NAGHD") <> 0 Then
        SDRST.AddNew 'مبلغ نقدصندوق
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
        SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
        SDRST.Fields("HES_T") = HFRST.Fields("SHIFT")
        SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & HFRST.Fields("DEPATMAN") & "-" & HFRST.Fields("SHIFT")
        If HFRST.Fields("M_NAGHD") > 0 Then
            SDRST.Fields("BEd") = HFRST.Fields("M_NAGHD")
        Else
            SDRST.Fields("BEs") = Abs(HFRST.Fields("M_NAGHD"))
        End If
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If Forms![BASEKNOW]![TKHF] = 1 Then
         If HFRST.Fields("TAKHFIF") <> 0 Then
            SDRST.AddNew 'تخفيف فروش
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
            SDRST.Fields("HES_M") = 1
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
            SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-1-1"
            SDRST.Fields("BED") = HFRST.Fields("TAKHFIF")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.Fields("TAG") = 13
            SDRST.update
         End If
      Else
            takh = 0
        If Forms![BASEKNOW]![TKHF] = 2 Then
            takh = 0
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, TAKHPERS.CUST_CO, TAKHPERS.TAKH_COD, TAKHPERS.TAFPER ,INVO_LST.MABL_K FROM INVO_LST INNER JOIN TAKHPERS ON INVO_LST.CODE = TAKHPERS.TAKH_COD WHERE (((INVO_LST.NUMBER)=" & HFRST.Fields("NUMBER") & ") AND ((INVO_LST.TAG)=2) AND ((TAKHPERS.CUST_CO)= " & HFRST.Fields("CUST_KIND") & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
              While Not rst.EOF()
                 If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
                    If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields(3))) Then
                            Err.Clear
                            Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields(3)), "تخفيف " & GETKALANAME(rst.Fields(3)))
                            If Err <> 0 Then
                               DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & GETKALANAME(rst.Fields(3)) & " - " & rst.Fields("code")
                            End If
                        End If
                        SDRST.AddNew 'تخفيف فروش
                        SDRST.Fields("N_S") = max_ns
                        SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                        SDRST.Fields("HES_M") = 3
                        SDRST.Fields("HES_T") = rst.Fields(3)
                        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & rst.Fields(3)
                        SDRST.Fields("BED") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                        SDRST.Fields("TAG") = 13
                        takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.update
                    Else
                        SDRST.AddNew 'تخفيف فروش
                        SDRST.Fields("N_S") = max_ns
                        SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                        SDRST.Fields("HES_M") = HFRST.Fields("CUST_KIND")
                        SDRST.Fields("HES_T") = rst.Fields(3)
                        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & HFRST.Fields("CUST_KIND") & "-" & rst.Fields(3)
                        SDRST.Fields("BED") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                        SDRST.Fields("TAG") = 13
                        takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.update
                    End If
                End If
                  rst.MoveNext
               Wend
            End If
            If HFRST.Fields("TAKHFIF") <> takh Then
               HFRST.Fields("TAKHFIF") = takh
               HFRST.update
            End If
      Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.HEAD_LST.NUMBER = " & HFRST.Fields("NUMBER") & ") AND (dbo.HEAD_LST.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          takh = 0
          While Not rst.EOF()
             If Round(rst.Fields("N_MOIN")) <> 0 Then
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & GETKALANAME(rst.Fields("CODE")) & " - " & rst.Fields("code")
                        End If
                    End If
                    SDRST.AddNew 'تخفيف فروش
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                    SDRST.Fields("HES_M") = 3
                    SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                    SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                    SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                    SDRST.Fields("TAG") = 13
                    takh = takh + Round(rst.Fields("N_MOIN"))
                    SDRST.update
                Else
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), CLng(HFRST.Fields("CUST_KIND")), CLng(rst.Fields("CODE"))) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), CLng(HFRST.Fields("CUST_KIND")), CLng(rst.Fields("CODE")), "تخفيف " & rst.Fields("CODE"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & "فاكتور شماره : " & HFRST.Fields("NUMBER") & "نام كالا : " & GETKALANAME(rst.Fields("CODE")) & " - " & rst.Fields("code")
                          End If
                    End If
                    SDRST.AddNew 'تخفيف فروش
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                    SDRST.Fields("HES_M") = HFRST.Fields("CUST_KIND")
                    SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & HFRST.Fields("CUST_KIND") & "-" & CLng(rst.Fields("CODE"))
                    SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                    SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 13
                    takh = takh + Round(rst.Fields("N_MOIN"))
                    SDRST.update
                End If
            End If
            rst.MoveNext
        Wend
      End If
      If HFRST.Fields("TAKHFIF") <> takh Then
         HFRST.Fields("TAKHFIF") = takh
         HFRST.update
      End If
    End If
    End If
      If HFRST.Fields("MABL_HAV") <> 0 Then
        SDRST.AddNew 'مبلغ حواله شخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BES") = HFRST.Fields("MABL_HAV")
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If HFRST.Fields("MABL_HAV") <> 0 Then
        SDRST.AddNew 'مبلغ حواله
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("MOIN_HAV")) Then
             Call GETTAF3(HFRST.Fields("MOIN_HAV"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("MOIN_HAV")
        SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("BED") = HFRST.Fields("MABL_HAV")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If HFRST.Fields("MABL_VAR") <> 0 Then
        SDRST.AddNew 'مبلغ واريزي شخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BES") = HFRST.Fields("MABL_VAR")
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If HFRST.Fields("MABL_VAR") <> 0 Then
        SDRST.AddNew 'مبلغ واريزي
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("MOIN_VAR")) Then
             Call GETTAF3(HFRST.Fields("MOIN_VAR"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("MOIN_VAR")
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BED") = HFRST.Fields("MABL_VAR")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.update
      End If
      If HFRST.Fields("MBAA") <> 0 Then
        SDRST.AddNew 'ماليات بر ارزش افزوده
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("HMBAA")) Then
             Call GETTAF3(HFRST.Fields("HMBAA"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("HMBAA")
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BES") = HFRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 13
        SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.update
      End If
      JAMP = 0
      If JAMF > 0 Then
        Set PRST = New ADODB.Recordset
        PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & HFRST.Fields("NUMBER") & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not PRST.EOF
            SDRST.AddNew 'پورسانت
            SDRST.Fields("N_S") = max_ns
            If Not IsNull(PRST.Fields("CUST_NO")) Then
                 Call GETTAF3(PRST.Fields("CUST_NO"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            End If
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = PRST.Fields("CUST_NO")
            If IsNull(TAMIR) Then
                TAMIR = PRST.Fields("CUST_NO")
            End If
            SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & HFRST.Fields("NUMBER") & " : " & HFRST.Fields("NUMBER1") & " بابت " & PRST.Fields("DARSAD") & "% مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")) & "مبلغ :  " & Format(Round(JAMF + HFRST.Fields("MABL_HAZ") + HFRST.Fields("MBAA") - HFRST.Fields("TAKHFIF")), "#,###") & " " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
            If IsNull(PRST.Fields("porid")) Then
                If Not PRST.Fields("STAT") Then
                    If Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100) <> PRST.Fields("PURSANT") Then
                        PRST.Fields("PURSANT") = Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100)
                        PRST.update
                    End If
                Else
                    If PRST.Fields("DARSAD") <> PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100 Then
                        PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100
                        PRST.update
                    End If
                End If
            Else
                Dim rst1 As New ADODB.Recordset, RST2 As New ADODB.Recordset, prs As Long, MBK As Long
                prs = 0
                MBK = 0
                rst1.Open "select code ,mabl_k  - n_moin as mablk from invo_lst where tag = 2 and number = " & HFRST.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst1.EOF
                    RST2.Open "SELECT  PORSANT FROM dbo.VISITORS_PORSANT_KALA WHERE (PORID = " & PRST.Fields("porid") & ") and (code = '" & rst1.Fields("code") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount = 1 Then
                        prs = prs + Round(rst1.Fields("mablk") * RST2.Fields("PORSANT") / 100)
                        MBK = MBK + rst1.Fields("mablk")
                    Else
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "تذكر مهم :اين كالا فاقد الگو براي اين ويزيتور است و پورسانت محاسبه نشد.درصورت لزوم براي آن تعريف كنيد و همينجا مجددا الگو را انتخاب كنيد  : " & GETKALANAME(rst1.Fields("code")) & " فاكتور شماره : " & HFRST.Fields("NUMBER")
                    End If
                    rst1.MoveNext
                    RST2.Close
                Wend
                rst1.Close
                PRST.Fields("PURSANT") = prs
                If MBK > 0 Then
                    PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / MBK * 100
                End If
                PRST.update
            End If
            SDRST.Fields("BES") = PRST.Fields("PURSANT")
            JAMP = JAMP + PRST.Fields("PURSANT")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("TAG") = 13
            If PRST.Fields("PURSANT") <> 0 Then
                SDRST.update
            End If
            PRST.MoveNext
        Wend
        If JAMP <> 0 Then
            If IsNull(Forms![BASEKNOW]![HPOR]) Then
                 DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب پورسانت در حسابهاي خودگردان مشخص نشده سند تراز نمي شود لطفا حساب پورسانت را مشخص كنيد و سند اين فاكتور را با دابل كليك مجددا صادر كنيد...!"
            Else
                SDRST.AddNew 'پورسانت
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![HPOR])
                SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![HPOR])
                SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![HPOR])
                SDRST.Fields("hes") = Forms![BASEKNOW]![HPOR]
                SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور فروش شماره " & HFRST.Fields("NUMBER") & " : " & HFRST.Fields("NUMBER1") & GETTAFNAME(TAMIR), 255)
                SDRST.Fields("BED") = JAMP
                SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 13
                SDRST.update
            End If
        End If
    End If
   Set rst = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   HFRST.MoveNext
   If (i Mod CON) = 0 Then
     Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
     Forms![GUG].Form.Repaint
   End If
   i = i + 1
   Forms![GUG]![Text2] = ""
Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub
Public Sub SANADKHORUGMAVAD(NUMBER As Long, NUMBER2 As Long)
                   On Error Resume Next
 Dim CON, i As Long
 DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
 DoCmd.restore
 DoCmd.OpenForm "GUG"
 Forms![GUG]![Text0] = "n"
 Forms![GUG]![Text0].Visible = True
 Forms![GUG]![Box7].Visible = True
 Forms![GUG].Form.Refresh
 Forms![GUG].Form.Repaint
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, HEDRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, RDD As Integer
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, HEAD_LST.USER_NAME FROM HEAD_LST WHERE ((HEAD_LST.NUMBER >= " & NUMBER & " AND HEAD_LST.NUMBER <=" & NUMBER2 & "  and HEAD_LST.tag = 10 ) )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.MoveLast
  HEDRST.MoveFirst
  Forms![GUG]![SNUM] = HEDRST.RecordCount
   CON = HEDRST.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
   i = 1
  While Not HEDRST.EOF()
        Forms![GUG]![num] = i
        rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Then
          max_ns = 1
        Else
          max_ns = rst.Fields(0) + 1
        End If
        Err.Clear
        SHART = "NO_S = 8 AND N_S = " & HEDRST.Fields("N_S")
        SHRST.Filter = SHART
        Set SHRST = SHRST
        If SHRST.RecordCount = 0 Or Err <> 0 Then
          Err.Clear
          SHRST.AddNew
          SHRST.Fields("N_S") = max_ns
          SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
          SHRST.Fields("SHARH_S") = Left(" حواله خروج مواد از انبار شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
          SHRST.Fields("GHATEI") = 0
          SHRST.Fields("NO_S") = 8
          SHRST.Fields("OKF") = True
          SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
          SHRST.Fields("CRT") = Now()
          SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
          SHRST.update
        Else
          max_ns = SHRST.Fields("N_S")
          SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
          SHRST.Fields("SHARH_S") = Left(" حواله خروج مواد از انبار شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
          SHRST.Fields("GHATEI") = 0
          SHRST.Fields("NO_S") = 8
          SHRST.Fields("OKF") = True
          SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
          SHRST.Fields("CRT") = Now()
          SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
          SHRST.update
        End If
        If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
           HEDRST.Fields("N_S") = max_ns
           HEDRST.update
        End If
        DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 10))")
        Set JST = New ADODB.Recordset
        If Not Forms![BASEKNOW]![FINALS] Then
             JST.Open "SELECT dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.HEAD_MANF.CODE AS COM, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NAM, dbo.HEAD_MANF.NAMES, dbo.HEAD_MANF.N_KOL, dbo.HEAD_MANF.NUMBER, dbo.HEAD_MANF.TNUMBER, dbo.DTL_MANF.SMABl AS SMAB FROM  dbo.STUF_DEF RIGHT OUTER JOIN dbo.HEAD_MANF INNER JOIN dbo.INVO_LST ON dbo.HEAD_MANF.FNUMB = dbo.INVO_LST.N_RASID ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE INNER JOIN dbo.DTL_MANF ON dbo.DTL_MANF.CODE = dbo.INVO_LST.CODE AND dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB AND dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.INVO_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") AND (dbo.INVO_LST.TAG = 10)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             While Not JST.EOF()
                   If JST.Fields("MABL_K") <> 0 Then
                       SDRST.AddNew 'انبار
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                       SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   If JST.Fields("MABL_K") <> 0 Then
                       RDD = GETGRPKALAco(JST.Fields("code"))
                       SDRST.AddNew 'پاياپاي هزينه توليد
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
                       SDRST.Fields("HES_M") = IIf(RDD = 2 Or RDD = 3, 2, 1)
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-" & IIf(RDD = 2 Or RDD = 3, 2, 1) & "-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                 Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                 Forms![GUG].Form.Repaint
                   If JST.Fields("MABL_K") <> 0 Then
                       SDRST.AddNew 'هزينه تولييا سير هزينه هاي توليد
                       SDRST.Fields("N_S") = max_ns
                       If IsNull(JST.Fields("COM")) Then
                           SDRST.Fields("HES_K") = JST.Fields("N_KOL")
                           SDRST.Fields("HES_M") = CLng(JST.Fields("NUMBER"))
                           SDRST.Fields("HES_T") = CLng(JST.Fields("TNUMBER"))
                           SDRST.Fields("hes") = JST.Fields("N_KOL") & "-" & CLng(JST.Fields("NUMBER")) & "-" & CLng(JST.Fields("TNUMBER"))
                       Else
                           SDRST.Fields("HES_K") = Forms![BASEKNOW]![HAZ_TOL]
                           SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                           SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                           SDRST.Fields("hes") = Forms![BASEKNOW]![HAZ_TOL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                       End If
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
                       SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   JAMCH = Round(JST.Fields("MABL_K"))
                   If JST.Fields("SMAB") * JST.Fields("MEGHk") <> 0 Then
                       SDRST.AddNew 'كنترل كالاي در جريان ساخت
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BED") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   If (JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk"))) <> 0 Then
                       SDRST.AddNew 'عملكرد
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                       If JAMCH > (JST.Fields("SMAB") * JST.Fields("MEGHk")) Then
                           SDRST.Fields("BED") = Round(JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk")))
                       Else
                           SDRST.Fields("BES") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk") - JAMCH)
                       End If
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                    End If
                    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                    Forms![GUG].Form.Repaint
                    JST.MoveNext
             Wend
            JST.Close
        Else
             JST.Open "SELECT dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, MAX(dbo.DTL_MANF.SMABL) AS SMAB FROM   dbo.INVO_LST INNER JOIN  dbo.DTL_MANF ON dbo.DTL_MANF.CODE = dbo.INVO_LST.CODE WHERE (dbo.INVO_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") And (dbo.INVO_LST.TAG = 10) GROUP BY dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             While Not JST.EOF()
                   If JST.Fields("MABL_K") <> 0 Then
                       SDRST.AddNew 'انبار
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                       SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   If JST.Fields("MABL_K") <> 0 Then
                       SDRST.AddNew 'پاياپاي هزينه توليد
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
                       SDRST.Fields("HES_M") = 1
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-1" & "-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                 Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                 Forms![GUG].Form.Repaint
                   If JST.Fields("MABL_K") <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![HAZ_TOL]), 99999, CLng(JST.Fields("code"))) Then
                         Err.Clear
                         Call CREATHES(CLng(Forms![BASEKNOW]![HAZ_TOL]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
                         If Err <> 0 Then
                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![HAZ_TOL] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                         End If
                       End If
                       SDRST.AddNew 'هزينه تولييا سير هزينه هاي توليد
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![HAZ_TOL]
                       SDRST.Fields("HES_M") = 99999
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("hes") = Forms![BASEKNOW]![HAZ_TOL] & "-99999-" & CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                       SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   JAMCH = Round(JST.Fields("MABL_K"))
                   If JST.Fields("SMAB") * JST.Fields("MEGHk") <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), 99999, CLng(JST.Fields("code"))) Then
                         Err.Clear
                         Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
                         If Err <> 0 Then
                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![CONKAL] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                         End If
                       End If
                       SDRST.AddNew 'كنترل كالاي در جريان ساخت
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                       SDRST.Fields("HES_M") = 99999
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-99999-" & CLng(JST.Fields("code"))
                       SDRST.Fields("BED") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk"))
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                   End If
                   Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                   Forms![GUG].Form.Repaint
                   If (JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk"))) <> 0 Then
                      If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("code"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![AMALKARD] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        End If
                       End If
                       SDRST.AddNew 'عملكرد
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                       SDRST.Fields("HES_M") = 99999
                       SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                       SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-99999-" & CLng(JST.Fields("code"))
                       SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                       If JAMCH > (JST.Fields("SMAB") * JST.Fields("MEGHk")) Then
                           SDRST.Fields("BED") = Round(JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk")))
                       Else
                           SDRST.Fields("BES") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk") - JAMCH)
                       End If
                       SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 10
                       SDRST.update
                    End If
                    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                    Forms![GUG].Form.Repaint
                    JST.MoveNext
             Wend
        End If
        JST.Close
        Set JST = New ADODB.Recordset
        JST.Open "SELECT  SUM(BED) AS SBED, SUM(BES) AS SBES FROM dbo.DEED_DTL WHERE (N_S = " & max_ns & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If JST.RecordCount > 0 Then
           If Abs(JST.Fields("SBED") - JST.Fields("SBES")) <= 40 And Abs(JST.Fields("SBED") - JST.Fields("SBES")) <> 0 Then
                 If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, 99999) Then
                    Err.Clear
                    Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, 99999, "كسر دهم ريال")
                    If Err <> 0 Then
                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![AMALKARD] & "-99999-99999" & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                    End If
                 End If
                 SDRST.AddNew 'عملكرد
                 SDRST.Fields("N_S") = max_ns
                 SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                 SDRST.Fields("HES_M") = 99999
                 SDRST.Fields("HES_T") = 99999
                 SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-99999-99999"
                 SDRST.Fields("SHARH") = Left("حواله خروج شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
                 If JST.Fields("SBED") - JST.Fields("SBES") > 0 Then
                     SDRST.Fields("BES") = JST.Fields("SBED") - JST.Fields("SBES")
                 Else
                     SDRST.Fields("BED") = Abs(JST.Fields("SBED") - JST.Fields("SBES"))
                 End If
                 SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                 SDRST.Fields("TAG") = 10
                 SDRST.update
            End If
        End If
        JST.Close
        HEDRST.MoveNext
        If (i Mod CON) = 0 Then
          Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
          Forms![GUG].Form.Repaint
        End If
        i = i + 1
        Forms![GUG]![Text2] = ""
        rst.Close
   Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub
Public Sub SANADKHORUGSAYER(NUMBER As Long, NUMBER2 As Long)
                 On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, JSTT As New ADODB.Recordset, HEDRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
   Dim CON, i As Long
   DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
   DoCmd.restore
   DoCmd.OpenForm "GUG"
   Forms![GUG]![Text0] = "n"
   Forms![GUG]![Text0].Visible = True
   Forms![GUG]![Box7].Visible = True
   Forms![GUG].Form.Refresh
   Forms![GUG].Form.Repaint
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, HEAD_LST.USER_NAME FROM HEAD_LST WHERE ((HEAD_LST.NUMBER >= " & NUMBER & " AND HEAD_LST.NUMBER <=" & NUMBER2 & "  and HEAD_LST.tag = 11 ) )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.MoveLast
  HEDRST.MoveFirst
  Forms![GUG]![SNUM] = HEDRST.RecordCount
   CON = HEDRST.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
   i = 1
  While Not HEDRST.EOF()
         Forms![GUG]![num] = i
          Set rst = New ADODB.Recordset
          rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If IsNull(rst.Fields(0)) Then
            max_ns = 1
          Else
            max_ns = rst.Fields(0) + 1
          End If
          Err.Clear
          SHART = "NO_S = 12 AND N_S = " & HEDRST.Fields("N_S")
          SHRST.Filter = SHART
          Set SHRST = SHRST
          If SHRST.RecordCount = 0 Or Err <> 0 Then
            SHRST.AddNew
            SHRST.Fields("N_S") = max_ns
            SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
            SHRST.Fields("SHARH_S") = Left(" حواله خروج ساير مواد از انبار شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & "مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
            SHRST.Fields("GHATEI") = 0
            SHRST.Fields("NO_S") = 12
            SHRST.Fields("OKF") = True
            SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
            SHRST.Fields("CRT") = Now()
            SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
            SHRST.update
          Else
               max_ns = SHRST.Fields("N_S")
                SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
                SHRST.Fields("SHARH_S") = Left(" حواله خروج ساير مواد از انبار شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & "مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
                SHRST.Fields("GHATEI") = 0
                SHRST.Fields("NO_S") = 12
                SHRST.Fields("OKF") = True
                SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
                SHRST.Fields("CRT") = Now()
                SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
                SHRST.update
          End If
          If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
             HEDRST.Fields("N_S") = max_ns
             HEDRST.update
          End If
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Repaint
          DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 11))")
          Set JST = New ADODB.Recordset
            DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 11))")
            JST.Open "SELECT     SANAD_NO, N_RASID, MABL_K, MEGHk, CODE, ANBAR FROM dbo.INVO_LST WHERE     (NUMBER = " & HEDRST.Fields("NUMBER") & ") AND (TAG = 11)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not JST.EOF()
               If IsNumeric(JST.Fields("N_RASID")) Then
                      JAMCH = 0
                      Set JSTT = New ADODB.Recordset
                      JSTT.Open "SELECT     FNUMB, NUMBER, TNUMBER, N_KOL, NAMES FROM dbo.HEAD_MANF WHERE     (FNUMB = " & JST.Fields("N_RASID") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If JST.Fields("MABL_K") <> 0 Then
                             SDRST.AddNew 'كنترل كالاي در جريان سا
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = JSTT.Fields("N_KOL")
                             SDRST.Fields("HES_M") = JSTT.Fields("NUMBER")
                             SDRST.Fields("HES_T") = JSTT.Fields("TNUMBER")
                             SDRST.Fields("hes") = JSTT.Fields("N_KOL") & "-" & JSTT.Fields("NUMBER") & "-" & JSTT.Fields("TNUMBER")
                             SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                             SDRST.Fields("SHARH") = Left("حواله خروج ساير شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JSTT.Fields("NAMES")), 255)
                             SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                             SDRST.Fields("TAG") = 11
                             SDRST.update
                             JAMCH = JST.Fields("MABL_K")
                             SDRST.AddNew 'انبار
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                             SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                             SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                             SDRST.Fields("SHARH") = Left("حواله خروج  ساير  مواد شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                             SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                             SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                             SDRST.Fields("TAG") = 11
                             SDRST.Fields("MHAZ_NO") = JST.Fields("SANAD_NO")
                             SDRST.update
                       End If
                       Set JSTT = New ADODB.Recordset
                Else
                      Call GETTAF3(JST.Fields("N_RASID"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                      If JST.Fields("MABL_K") <> 0 Then
                          SDRST.AddNew 'كنترل كالاي در جريان سا
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = CKOL
                          SDRST.Fields("HES_M") = CMOIN
                          SDRST.Fields("HES_T") = CTAF
                          SDRST.Fields("HES_T2") = CTAF2
                          SDRST.Fields("HES_T3") = CTAF3
                          SDRST.Fields("HES_T4") = CTAF4
                          SDRST.Fields("hes") = JST.Fields("N_RASID")
                          SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                          SDRST.Fields("SHARH") = Left("حواله خروج ساير شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(GETTAFNAME(JST.Fields("N_RASID"))), 255)
                          SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 11
                          SDRST.update
                          JAMCH = JST.Fields("MABL_K")
                          SDRST.AddNew 'انبار
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                          SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                          SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                          SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                          SDRST.Fields("SHARH") = Left("حواله خروج  ساير  مواد شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                          SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                          SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 11
                          SDRST.Fields("MHAZ_NO") = JST.Fields("SANAD_NO")
                          SDRST.update
                      End If
                End If
                JST.MoveNext
            Wend
       JST.Close
      HEDRST.MoveNext
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
      rst.Close
   Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub
Public Sub SANADENTEGHAL(NUMBER As Long, NUMBER2 As Long)
                 On Error Resume Next
   Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, HEDRST As New ADODB.Recordset
   Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String
   Dim CON, i As Long
   Dim arst As New ADODB.Recordset
   DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
   DoCmd.restore
   DoCmd.OpenForm "GUG"
   Forms![GUG]![Text0] = "n"
   Forms![GUG]![Text0].Visible = True
   Forms![GUG]![Box7].Visible = True
   Forms![GUG].Form.Refresh
   Forms![GUG].Form.Repaint
   SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   HEDRST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, HEAD_LST.USER_NAME FROM HEAD_LST WHERE ((HEAD_LST.NUMBER >= " & NUMBER & " AND HEAD_LST.NUMBER <=" & NUMBER2 & "  and HEAD_LST.tag = 5 ) )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   HEDRST.MoveLast
   HEDRST.MoveFirst
   Forms![GUG]![SNUM] = HEDRST.RecordCount
   CON = HEDRST.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
   i = 1
   While Not HEDRST.EOF()
        Forms![GUG]![num] = i
        arst.Open "SELECT  CODE, KIND FROM dbo.TCOD_ANBAR WHERE (CODE = " & HEDRST.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If arst.RecordCount > 0 Then
           If arst.Fields("kind") = 1 Or arst.Fields("kind") = 2 Then
              If Not (Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT])) Then
                   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 5))")
                   arst.Close
                   Set arst = New ADODB.Recordset
                   GoTo MV
              End If
           End If
        End If
        arst.Close
        Set arst = New ADODB.Recordset
          Set rst = New ADODB.Recordset
          rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If IsNull(rst.Fields(0)) Then
            max_ns = 1
          Else
            max_ns = rst.Fields(0) + 1
          End If
          Err.Clear
          SHART = "NO_S = 10 AND N_S = " & HEDRST.Fields("N_S")
          SHRST.Filter = SHART
          Set SHRST = SHRST
          Calendar = vbCalHijri
          If SHRST.RecordCount = 0 Or Err <> 0 Then
            SHRST.AddNew
            SHRST.Fields("N_S") = max_ns
            SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
            SHRST.Fields("SHARH_S") = Left(" حواله انتقالي مواد شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " از انبار " & HEDRST.Fields("ANBAR") & " به " & HEDRST.Fields("ANBARF") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
            SHRST.Fields("GHATEI") = 0
            SHRST.Fields("NO_S") = 10
            SHRST.Fields("OKF") = True
            SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
            SHRST.Fields("CRT") = Now()
            SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
            SHRST.update
          Else
            max_ns = SHRST.Fields("N_S")
            SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
            SHRST.Fields("SHARH_S") = Left(" حواله انتقالي مواد شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " از انبار " & HEDRST.Fields("ANBAR") & " به " & HEDRST.Fields("ANBARF") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
            SHRST.Fields("GHATEI") = 0
            SHRST.Fields("NO_S") = 10
            SHRST.Fields("OKF") = True
            SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
            SHRST.update
          End If
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Repaint
          If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
             HEDRST.Fields("N_S") = max_ns
          End If
          DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 5))")
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
          JST.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, STUF_DEF.NAME, INVO_LST.ANBAR, INVO_LST.CODE, INVO_LST.MEGH, INVO_LST.MEGHk, INVO_LST.MEGH_MAR, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.ANBARF FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & HEDRST.Fields("NUMBER") & ") AND ((INVO_LST.TAG)=5))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
           While Not JST.EOF()
               If JST.Fields("MABL_K") <> 0 Then
                If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(2))
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                   End If
                End If
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
               SDRST.AddNew 'از انبار
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                SDRST.Fields("SHARH") = Left("حواله انتقالي شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 5
                SDRST.update
               End If
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
               If JST.Fields("MABL_K") <> 0 Then
                If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBARF")), CLng(JST.Fields("code"))) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBARF")), CLng(JST.Fields("code")), JST.Fields(2))
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                   End If
                End If
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
                SDRST.AddNew 'به انبار
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                SDRST.Fields("HES_M") = JST.Fields("ANBARF")
                SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBARF") & "-" & CLng(JST.Fields("code"))
                SDRST.Fields("SHARH") = Left("حواله انتقالي شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & "  بابت " & JST.Fields("name"), 255)
                SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 5
                SDRST.update
               End If
               JST.MoveNext
              Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
              Forms![GUG].Form.Repaint
          Wend
MV:      JST.Close
      HEDRST.MoveNext
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
      rst.Close
   Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub
Function LASTAVRAGE(CODE As String, ANBAR As Long, dt As Long) As Double
  Dim RRST As New ADODB.Recordset
    LASTAVRAGE = 0
        If MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) <> "5" Then
            LASTAVRAGE = GETSTANDARDPRICE_KOL(CODE, dt)
        End If
        If LASTAVRAGE = 0 Then
            Set RRST = New ADODB.Recordset
            RRST.Open "SELECT TOP 1 PERCENT ANBAR, CODE, DATE_N, avrage, TAG, NUMBER FROM         dbo.KA_KH(800101) KA_KH  WHERE     (CODE = N'" & CODE & "') AND (ANBAR = " & ANBAR & ")AND (DATE_N <= " & dt & ") AND (TAG = 0 OR  TAG = 1 OR TAG = 7 OR TAG = 9 OR TAG = 6 ) ORDER BY DATE_N DESC, NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RRST.RecordCount > 0 And Not IsNull(RRST.Fields("AVRAGE")) Then
               LASTAVRAGE = RRST.Fields("AVRAGE")
            Else
                    Set RRST = New ADODB.Recordset
                    RRST.Open "SELECT TOP 1 PERCENT dbo.INVO_LST.AVRAGE2, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.INVO_LST.NUMBER      AS MNN FROM dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.HEAD_LST.DATE_N <= " & dt & ") AND (dbo.INVO_LST.CODE = '" & CODE & "') AND (dbo.INVO_LST.ANBARF = " & ANBAR & ")  ORDER BY dbo.HEAD_LST.DATE_N DESC, dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RRST.RecordCount > 0 And Not IsNull(RRST.Fields("AVRAGE2")) Then
                          LASTAVRAGE = RRST.Fields("AVRAGE2")
                    Else
                        Set RRST = New ADODB.Recordset
                        RRST.Open "SELECT FI_A FROM  dbo.STUF_FSK WHERE (CODE =  '" & CODE & "') AND (ANBAR = " & ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If RRST.RecordCount > 0 And Not IsNull(RRST.Fields("FI_A")) Then
                              LASTAVRAGE = RRST.Fields("FI_A")
                        Else
                              LASTAVRAGE = 0
                        End If
                    End If
            End If
        End If

End Function

Function GETKALANAME(CODE As Double) As String
  Dim RRST As New ADODB.Recordset
  RRST.Open "SELECT CODE,NAME FROM STUF_DEF WHERE (CODE =" & CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If RRST.RecordCount > 0 Then
    If IsNull(RRST.Fields("NAME")) Then
        GETKALANAME = " "
    Else
        GETKALANAME = RRST.Fields("NAME")
    End If
  Else
    GETKALANAME = " "
  End If
  RRST.Close
End Function

Public Function GETSTANDARDPRICE(CODE As String) As Double

                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "') GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ORDER BY dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        GETSTANDARDPRICE = rst.Fields("SumOfMABLK") + rst.Fields("SumOfIMBIBE_SAR") + rst.Fields("SumOfIMBIBE_MANF")
   Else
        GETSTANDARDPRICE = 0
   End If
End Function

Function LASTPRICE(CODE As String, dt As Long) As Long
  Dim RRST As New ADODB.Recordset
    RRST.Open "SELECT TOP 1 PERCENT dbo.INVO_LST.AVRAGE, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.MABL FROM dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.HEAD_LST.DATE_N <= " & dt & ") AND dbo.INVO_LST.TAG = 1 GROUP BY dbo.INVO_LST.AVRAGE, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.MABL HAVING (dbo.INVO_LST.CODE = '" & CODE & "') ORDER BY dbo.HEAD_LST.DATE_N DESC, dbo.INVO_LST.MABL DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RRST.RecordCount > 0 Then
      If IsNull(RRST.Fields("MABL")) Then
          LASTPRICE = 0
      Else
          LASTPRICE = RRST.Fields("MABL")
      End If
    Else
      Set RRST = New ADODB.Recordset
      RRST.Open "SELECT     TOP 1 PERCENT CODE, FI_A FROM dbo.STUF_FSK WHERE code = '" & CODE & "' ORDER BY FI_A DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RRST.RecordCount > 0 Then
           If IsNull(RRST.Fields("FI_A")) Then
                LASTPRICE = 0
           Else
              LASTPRICE = RRST.Fields("FI_A")
           End If
        Else
           LASTPRICE = 0
        End If
    End If
  RRST.Close
End Function

Public Function MGETKOL(SHES As String) As Long
    Dim i As Byte
    i = 1
    While MID(SHES, i, 1) <> "-" And i <= Len(SHES)
        i = i + 1
    Wend
    MGETKOL = Left(SHES, i - 1)
End Function
Public Function MGETMOIN(SHES As String) As Long
    Dim i, j As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
    While MID(SHES, i, 1) <> "-"
        i = i + 1
    Wend
    j = i + 1
    While MID(SHES, j, 1) <> "-" And j <= Len(SHES)
        j = j + 1
    Wend
    i = i + 1
    MGETMOIN = MID(SHES, i, j - i)
    End If
End Function
Public Function MGETTAF(SHES As String) As Long
    Dim i, j As Byte
    i = 1
    If (Len(SHES) < 5) Then
    Else
        While MID(SHES, i, 1) <> "-"
            i = i + 1
        Wend
        j = i + 1
        If j = Len(SHES) Then

        Else
            While MID(SHES, j, 1) <> "-" And j < Len(SHES)
                j = j + 1
            Wend
            If j >= Len(SHES) Then
                MGETTAF = 0
            Else
                MGETTAF = MID(SHES, j + 1, Len(SHES))
            End If
         End If
    End If
End Function

Public Function BLOCKED(HKOL As Long, HMOIN As Long, HTAF As Long) As Boolean
                 On Error Resume Next
    Dim RST2 As New ADODB.Recordset, TKOL As Long, TMOIN As Long, TTAF As Long
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            TKOL = MGETKOL(RST2.Fields("HES"))
            TMOIN = MGETMOIN(RST2.Fields("HES"))
            TTAF = MGETTAF(RST2.Fields("HES"))
            If TKOL = HKOL And TMOIN = 0 And TTAF = 0 Then
                 BLOCKED = True
                 Exit Function
            Else
                If TKOL = HKOL And TMOIN = HMOIN And TTAF = 0 Then
                    BLOCKED = True
                    Exit Function
                Else
                    If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF Then
                        BLOCKED = True
                        Exit Function
                    End If
                End If
            End If
            RST2.MoveNext
        Wend
      End If
      BLOCKED = False
End Function

Public Function BLOCKEDM(HKOL As Long, HMOIN As Long, HTAF As Long, Optional HTAF2 As Long) As Boolean
                 On Error Resume Next
    Dim RST2 As New ADODB.Recordset, TKOL As Long, TMOIN As Long, TTAF As Long, TTAF2 As Long
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            TKOL = MGETKOL(RST2.Fields("HES"))
            TMOIN = MGETMOIN(RST2.Fields("HES"))
            TTAF = MGETTAF(RST2.Fields("HES"))
            TTAF2 = GETTAF2(RST2.Fields("HES"))
            If TKOL = HKOL And TMOIN = 0 And TTAF = 0 And TTAF2 = -1 Then
                 BLOCKEDM = True
                 Exit Function
            Else
                If TKOL = HKOL And TMOIN = HMOIN And TTAF = 0 And TTAF2 = -1 Then
                    BLOCKEDM = True
                    Exit Function
                Else
                    If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = -1 Then
                        BLOCKEDM = True
                        Exit Function
                    Else
                        If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 Then
                            BLOCKEDM = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            RST2.MoveNext
        Wend
      End If
      BLOCKEDM = False
End Function
Public Function BLOCKEDCUST(HES As String) As Boolean
    Dim RST2 As New ADODB.Recordset
     RST2.Open "SELECT     HES FROM dbo.BLOCK_CUSTOMER  WHERE     (HES = N'" & HES & "') AND (ENDBLK = 0) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        BLOCKEDCUST = True
        Exit Function
     End If
      BLOCKEDCUST = False

End Function
Function GETFIRSTPRICE(CODE As String) As Long
  Dim RRST As New ADODB.Recordset
    RRST.Open "SELECT     TOP 1 PERCENT CODE, FI_A FROM dbo.STUF_FSK WHERE code = '" & CODE & "' ORDER BY FI_A DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RRST.RecordCount > 0 Then
       If IsNull(RRST.Fields("FI_A")) Then
            GETFIRSTPRICE = 0
       Else
          GETFIRSTPRICE = RRST.Fields("FI_A")
       End If
    Else
       GETFIRSTPRICE = 0
    End If
  RRST.Close
End Function

Public Function Checketebar(HES As String) As Integer
   Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
               On Error Resume Next
   CHK = 0
   TOPETEB = 0
    rst.Open "SELECT * FROM AZAE WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       TOPETEB = rst.Fields("TOPETEB")
       TOPASN = rst.Fields("TOPASN")
       ENDIS = rst.Fields("ENDIS")
    End If
  If ENDIS Then
      Set rst = New ADODB.Recordset
      rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        MAND = 0
      Else
        MAND = rst.Fields(0)
      End If
      If Not Forms![BASEKNOW]![SAGHF] Then
        MAND = 0
      End If
      Set rst = New ADODB.Recordset
      rst.Open " SELECT     SUM(MABL) AS smab FROM         dbo.PAY_GETD WHERE     (N_KOL3 IS NULL) AND (N_KOL2 IS NULL) AND (CUST_NO = N'" & HES & "') AND (N_KOL = " & Forms![BASEKNOW]![BANKHA] & " OR  N_KOL IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
       CHK = 0
      Else
         CHK = rst.Fields(0)
      End If
    'چك صفر شد
    If Not Forms![BASEKNOW]![SAGHF2] Or Not TOPASN Then
      CHK = 0
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM AZAE WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      If TOPETEB - MAND - CHK >= 0 Then
         Checketebar = True
      Else
         Checketebar = False
         DoCmd.OpenForm "mesageform", , , , , acDialog, "مانده حساب : " & Format(MAND, "#,###") & "  مانده اسناد : " & Format(CHK, "#,###") & "  كل اعتبار : " & Format(TOPETEB, "#,###")
      End If
    Else
      Checketebar = 2
    End If
  Else
      Checketebar = True
  End If

End Function
Public Function ChecketebarMEG(HES As String) As Integer
   Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPMEGH As Double, ENDIS As Boolean
                    On Error Resume Next
    TOPMEGH = 0
    rst.Open "SELECT * FROM AZAE WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       TOPMEGH = rst.Fields("TOPMEGH")
       ENDIS = rst.Fields("ENDIS")
    End If
    If ENDIS And TOPMEGH <> 0 Then
          Set rst = New ADODB.Recordset
          rst.Open "SELECT     SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEGH FROM    dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            MAND = 0
          Else
            MAND = rst.Fields(0)
          End If
        If Not Forms![BASEKNOW]![SAGHF] Then
          MAND = 0
        End If
        If TOPMEGH - MAND >= 0 Then
            ChecketebarMEG = True
        Else
            ChecketebarMEG = False
            DoCmd.OpenForm "mesageform", , , , , acDialog, "  كل اعتبار مقداري : " & Format(TOPMEGH, "#,###") & "  كل خريد مقداري : " & Format(MAND, "#,###")
        End If
  Else
      ChecketebarMEG = True
  End If

End Function

Public Sub GENSANADKHAZ(fnum, TNUM As Long)
                 On Error Resume Next
    Dim CON, i As Long
    DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
    DoCmd.OpenForm "GUG"
    Forms![GUG]![Text0] = "n"
    Forms![GUG]![Text0].Visible = True
    Forms![GUG]![Box7].Visible = True
    Forms![GUG].Form.Refresh
    Forms![GUG].Form.Repaint
    Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHRH As String
    Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL As Double, CMOIN As Double, CTAF As Double, CTAF2 As Double, takh As Double, SHART As String
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    HFRST.Open "SELECT     * FROM dbo.PGET_HED WHERE     (ID BETWEEN " & fnum & " AND " & TNUM & ")  ORDER BY ID", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    HFRST.MoveLast
    HFRST.MoveFirst
    Forms![GUG]![SNUM] = HFRST.RecordCount
    CON = HFRST.RecordCount / 50
    If CON < 1 Then
      CON = 1
    End If
    i = 1
    While Not HFRST.EOF
          Forms![GUG]![num] = i
          rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If IsNull(rst.Fields(0)) Then
            max_ns = 1
          Else
            max_ns = rst.Fields(0) + 1
          End If
          SHRH = "خزانه داري شماره " & HFRST.Fields("ID") & "مورخ " & Format(HFRST.Fields("DATE"), "####/##/##")
          If IsNull(HFRST.Fields("N_S")) Then
                SHRST.AddNew
                SHRST.Fields("N_S") = max_ns
                SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
                SHRST.Fields("SHARH_S") = SHRH
                SHRST.Fields("GHATEI") = 0
                SHRST.Fields("NO_S") = 5
                SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
                SHRST.Fields("OKF") = True
                SHRST.Fields("CRT") = Now()
                SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
                SHRST.update
                HFRST.Fields("N_S") = max_ns
                HFRST.update
          Else
              SHART = "NO_S = 5 AND N_S = " & HFRST.Fields("N_S")
              SHRST.Filter = SHART
              Set SHRST = SHRST
              If SHRST.RecordCount = 0 Then
                SHRST.AddNew
                SHRST.Fields("N_S") = max_ns
                SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
                SHRST.Fields("SHARH_S") = SHRH
                SHRST.Fields("GHATEI") = 0
                SHRST.Fields("NO_S") = 5
                SHRST.Fields("OKF") = True
                SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
                SHRST.Fields("CRT") = Now()
                SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
                SHRST.update
              Else
                SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
                SHRST.Fields("SHARH_S") = SHRH
                SHRST.Fields("USER_NAME") = UCurrentUser()
                SHRST.Fields("OKF") = True
                SHRST.update
              End If
          End If
          If Not IsNull(HFRST.Fields("N_S")) Then
            DoCmd.RunSQL ("DELETE FROM dbo.DEED_DTL WHERE (N_S = " & HFRST.Fields("N_S") & ")")
          End If
          DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BED, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     THES_K, THES_M, THES_T, THES_T2, THES_T3, THES_T4, SHARH, MABL, N_SERI, BANK," & HFRST.Fields("N_S") & " AS Expr1,THES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & HFRST.Fields("ID") & ")"
          DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BES, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     FHES_K, FHES_M, FHES_T, FHES_T2, FHES_T3, FHES_T4, SHARH, MABL, N_SERI, BANK," & HFRST.Fields("N_S") & " AS Expr1,FHES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & HFRST.Fields("ID") & ")"
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Refresh
          Forms![GUG]![Text2].Requery
          Forms![GUG].Form.Repaint
          rst.Close
          If (i Mod CON) = 0 Then
              Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
              Forms![GUG].Form.Refresh
              Forms![GUG]![Text0].Requery
              Forms![GUG].Form.Repaint
              Forms![GUG]![Text2] = ""
          End If
          i = i + 1
          HFRST.MoveNext
    Wend
    SHRST.Close
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub


Public Sub GENSANADVD(fnum, TNUM As Long)
                     On Error Resume Next
    Dim CON, i As Long
    DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
    DoCmd.OpenForm "GUG"
    Forms![GUG]![Text0] = "n"
    Forms![GUG]![Text0].Visible = True
    Forms![GUG]![Box7].Visible = True
    Forms![GUG].Form.Refresh
    Forms![GUG].Form.Repaint
    Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHRH As String, CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D
    Dim max_ns, MABL_CHK, JAMF, JAMCH, takh As Double, SHART As String
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    HFRST.Open "SELECT     * FROM dbo.CHKREC_H WHERE     (IDH BETWEEN " & fnum & " AND " & TNUM & ")  ORDER BY IDH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    HFRST.MoveLast
    HFRST.MoveFirst
    Forms![GUG]![SNUM] = HFRST.RecordCount
    CON = HFRST.RecordCount / 50
    If CON < 1 Then
      CON = 1
    End If
    i = 1
    While Not HFRST.EOF
          Forms![GUG]![num] = i
          rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If IsNull(rst.Fields(0)) Then
            max_ns = 1
          Else
            max_ns = rst.Fields(0) + 1
          End If
          If IsNull(HFRST.Fields("N_S")) Then
            SHRST.AddNew
            SHRST.Fields("N_S") = max_ns
            SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
            SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
            SHRST.Fields("GHATEI") = 0
            SHRST.Fields("NO_S") = 6
            SHRST.Fields("OKF") = True
            SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
            SHRST.Fields("CRT") = Now()
            SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
            SHRST.update
          Else
            SHART = "N_S = " & HFRST.Fields("N_S") & " AND  NO_S = 6 "
            SHRST.Filter = SHART
            Set SHRST = SHRST
            If SHRST.RecordCount = 0 Then
                SHRST.AddNew
                SHRST.Fields("N_S") = max_ns
                SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
                SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
                SHRST.Fields("GHATEI") = 0
                SHRST.Fields("NO_S") = 6
                SHRST.Fields("OKF") = True
                SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
                SHRST.Fields("CRT") = Now()
                SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
                SHRST.update
            Else
              SHRST.Fields("DATE_S") = HFRST.Fields("DATE")
              SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
              SHRST.Fields("GHATEI") = 0
              SHRST.Fields("NO_S") = 6
              SHRST.Fields("OKF") = True
              SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
              SHRST.update
            End If
          End If
          If IsNull(HFRST.Fields("N_S")) Then
            HFRST.Fields("N_S") = max_ns
          End If

          If Not IsNull(HFRST.Fields("N_S")) Then
            DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & HFRST.Fields("N_S") & " ))")
          End If
          Set rst = New ADODB.Recordset
          rst.Open "SELECT     dbo.CHKREC_H.idh, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND, dbo.PAY_GETD.HES1   FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.idH = " & HFRST.Fields("IDH") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
             Call GETTAF3([Forms]![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
          End If
          If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
             Call GETTAF3([Forms]![BASEKNOW]![ADV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
          End If
          While Not rst.EOF
              SDRST.AddNew
              SDRST.Fields("N_S") = HFRST.Fields("N_S")
              If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                  SDRST.Fields("HES_K") = CKOL
                  SDRST.Fields("HES_M") = CMOIN
                  SDRST.Fields("HES_T") = CTAF
                  SDRST.Fields("HES_T2") = CTAF2
                  SDRST.Fields("HES_T3") = CTAF3
                  SDRST.Fields("HES_T4") = CTAF4
                  SDRST.Fields("hes") = [Forms]![BASEKNOW]![ADA]
              Else
                  SDRST.Fields("HES_K") = CKOLV
                  SDRST.Fields("HES_M") = CMOINV
                  SDRST.Fields("HES_T") = CTAFV
                  SDRST.Fields("HES_T2") = CTAF2V
                  SDRST.Fields("HES_T3") = CTAF3V
                  SDRST.Fields("HES_T4") = CTAF4V
                  SDRST.Fields("hes") = [Forms]![BASEKNOW]![ADV]
              End If
              SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
              SDRST.Fields("BANK") = rst.Fields("BANK")
              SDRST.Fields("BES") = rst.Fields("MABL")
              SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
              SDRST.update
              Call GETTAF3(rst.Fields("HES1"), CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D)
              SDRST.AddNew
              SDRST.Fields("N_S") = HFRST.Fields("N_S")
              SDRST.Fields("HES_K") = CKOLD
              SDRST.Fields("HES_M") = CMOIND
              SDRST.Fields("HES_T") = CTAFD
              SDRST.Fields("HES_T2") = CTAF2D
              SDRST.Fields("HES_T3") = CTAF3D
              SDRST.Fields("HES_T4") = CTAF4D
              SDRST.Fields("hes") = rst.Fields("HES1")
              SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
              SDRST.Fields("BANK") = rst.Fields("BANK")
              SDRST.Fields("BED") = rst.Fields("MABL")
              SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
              SDRST.update
              rst.Fields("N_S") = HFRST.Fields("N_S")
              rst.update
              rst.MoveNext
          Wend
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Refresh
          Forms![GUG]![Text2].Requery
          Forms![GUG].Form.Repaint
          rst.Close
          If (i Mod CON) = 0 Then
              Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
              Forms![GUG].Form.Refresh
              Forms![GUG]![Text0].Requery
              Forms![GUG].Form.Repaint
              Forms![GUG]![Text2] = ""
          End If
          i = i + 1
          HFRST.MoveNext
    Wend
    SHRST.Close
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub

Public Sub GENSANADKHAREED(fnum, TNUM As Long)
                 On Error Resume Next
    Dim CON, i As Long, K As Long
    DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
    DoCmd.OpenForm "GUG"
    Forms![GUG]![Text0] = "n"
    Forms![GUG]![Text0].Visible = True
    Forms![GUG]![Box7].Visible = True
    Forms![GUG].Form.Refresh
    Forms![GUG].Form.Repaint
    Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, SARST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHSH As String
    Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, takh As Double, SHART As String, KHMAVAV  As Double, KHNIM  As Double, KHSAKHT  As Double, KHSAY As Double, BAZAR As Double, HS(7) As Double
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  '  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Set SDRST = Forms![recording].Form.RecordsetClone
    HFRST.Open "SELECT     * FROM dbo.HEAD_LST WHERE     (NUMBER BETWEEN " & fnum & " AND " & TNUM & ") AND (TAG = 12) ORDER BY NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    HFRST.MoveLast
    HFRST.MoveFirst
    Forms![GUG]![SNUM] = HFRST.RecordCount
    CON = HFRST.RecordCount / 50
    If CON < 1 Then
      CON = 1
    End If
    i = 1
    While Not HFRST.EOF
      Forms![GUG]![num] = i
      If Not IsNull(HFRST.Fields("CUST_NO")) Then
         Call GETTAF3(HFRST.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
      SHSH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي  خريد  " & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255), Left(" فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " خريدار: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255))
      If Forms![BASEKNOW]![SNDKH] Then 'سند روزانه است
        If Not IsNull(HFRST.Fields("N_S")) Then 'فاکتور سند دارد
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT   BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and n_s = " & HFRST.Fields("N_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then  'اگرسند  فاکتورهست
                If SARST.Fields("DATE_S") = HFRST.Fields("DATE_N") Then 'تاريخ سند و فاکتوريکي است
                    max_ns = HFRST.Fields("N_S")
                Else
SEJ:
                    Set SARST = New ADODB.Recordset: SARST.Open "SELECT    BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and DATE_s = " & HFRST.Fields("DATE_N"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If SARST.RecordCount > 0 Then   'اگرسند به تاريخ فاکتورهست
                        max_ns = SARST.Fields("N_S")
                    Else
                        max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 1, -1, HFRST.Fields("USER_NAME"))
                        HFRST.Fields("N_S") = max_ns
                    End If
                End If
            Else
                GoTo SEJ ' چک کن اگه نيست صادر کن
            End If
        Else
           GoTo SEJ ' چک کن اگه نيست صادر کن
        End If
      Else 'تک سندي
        If Not IsNull(HFRST.Fields("N_S")) Then 'فاکتور سند دارد
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT    n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and N_s = " & HFRST.Fields("N_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then   'اگرسند فاکتورهست
                If SARST.Fields("DATE_S") <> HFRST.Fields("DATE_N") Then 'تاريخ سند و فاکتوريکي است
                    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & HFRST.Fields("DATE_N") & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 1,OKF=-1,USER_NAME ='" & HFRST.Fields("USER_NAME") & "' WHERE N_S =" & HFRST.Fields("N_S"))
                End If
                max_ns = HFRST.Fields("N_S")
            Else
                max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 1, -1, HFRST.Fields("USER_NAME"))
                HFRST.Fields("N_S") = max_ns
            End If
        Else
            max_ns = Createsanad(HFRST.Fields("DATE_N"), SHSH, 0, 1, -1, HFRST.Fields("USER_NAME"))
            HFRST.Fields("N_S") = max_ns
        End If
      End If

    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      If IsNull(HFRST.Fields("N_S")) Or HFRST.Fields("N_S") <> max_ns Then
         HFRST.Fields("N_S") = max_ns
         HFRST.update
      End If
     JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & HFRST.Fields("NUMBER") & " ) AND ((INVO_LST.TAG)=1))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          JAMF = Round(JST.Fields(0))
     Else
        JAMF = 0
     End If
     JST.Close
     Set JST = New ADODB.Recordset
     JST.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.TAG)=1) AND ((PAY_GETP.NUMBER)= " & HFRST.Fields("NUMBER") & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
           JAMCH = JST.Fields(0)
     Else
        JAMCH = 0
     End If
     KHMAVAV = 0
     KHNIM = 0
     KHSAKHT = 0
     KHSAY = 0
     BAZAR = 0
     HS(1) = 0
     HS(2) = 0
     HS(3) = 0
     HS(4) = 0
     HS(5) = 0
     HS(6) = 0
     HS(7) = 0
     JST.Close
     Set JST = New ADODB.Recordset
    DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HFRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 12))")
    JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, dbo.STUF_DEF.RADAH FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & HFRST.Fields("NUMBER") & ") AND ((INVO_LST.TAG)=1)) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
    Forms![GUG].Form.Refresh
    Forms![GUG]![Text2].Requery
    Forms![GUG].Form.Repaint
    While Not JST.EOF()
      If JST.Fields("MABL_K") <> 0 Then
           If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE"))) Then
              Err.Clear
              Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE")), JST.Fields(4))
              If Err <> 0 Then
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              End If
           End If
           SDRST.AddNew 'خريد
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
           SDRST.Fields("HES_M") = JST.Fields("ANBAR")
           SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
           SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("CODE"))
           SDRST.Fields("SHARH") = RIGHT("خريدفاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
           SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12
           SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           Select Case JST.Fields("RADAH")
               Case 1: KHMAVAV = KHMAVAV + Round(JST.Fields("MABL_K"))
               Case 2: KHNIM = KHNIM + Round(JST.Fields("MABL_K"))
               Case 3: KHSAKHT = KHSAKHT + Round(JST.Fields("MABL_K"))
               Case 4: BAZAR = BAZAR + Round(JST.Fields("MABL_K"))
               Case 5: HS(1) = HS(1) + Round(JST.Fields("MABL_K"))
               Case 6: HS(2) = HS(2) + Round(JST.Fields("MABL_K"))
               Case 7: HS(3) = HS(3) + Round(JST.Fields("MABL_K"))
               Case 8: HS(4) = HS(4) + Round(JST.Fields("MABL_K"))
               Case 9: HS(5) = HS(5) + Round(JST.Fields("MABL_K"))
               Case 10: HS(6) = HS(6) + Round(JST.Fields("MABL_K"))
               Case Else: KHSAY = KHSAY + Round(JST.Fields("MABL_K"))
           End Select
           SDRST.update
       End If
       JST.MoveNext
     Wend
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("MABL_HAZ") <> 0 Then
       SDRST.AddNew 'كرايه حمل يا غيره
       SDRST.Fields("N_S") = max_ns
       If Not IsNull(HFRST.Fields("MOIN_HAZ")) Then
             Call GETTAF3(HFRST.Fields("MOIN_HAZ"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
       End If
       SDRST.Fields("HES_K") = HKOL
       SDRST.Fields("HES_M") = HMOIN
       SDRST.Fields("HES_T") = HTAF
       SDRST.Fields("HES_T2") = HTAF2
       SDRST.Fields("HES_T3") = HTAF3
       SDRST.Fields("HES_T4") = HTAF4
       SDRST.Fields("hes") = HFRST.Fields("MOIN_HAZ")
       SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " - " & GETTAFNAME(HFRST.Fields("MOIN_HAZ")), 255)
       SDRST.Fields("BED") = HFRST.Fields("MABL_HAZ")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
     If JAMCH <> 0 Then 'چكهاي دريافتي
       CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       CHRST.MoveLast
       CHRST.MoveFirst
       CHRST.Filter = "NUMBER = " & HFRST.Fields("NUMBER") & " AND TAG = 1"
       Set CHRST = CHRST
       Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
       Forms![GUG].Form.Refresh
       Forms![GUG]![Text2].Requery
       Forms![GUG].Form.Repaint
       If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
       Do Until CHRST.EOF
           MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
           SDRST.AddNew 'اسناد پرداختني
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![APA])
           SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![APA])
           SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![APA])
           SDRST.Fields("hes") = Forms![BASEKNOW]![APA]
           SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
           SDRST.Fields("BES") = CHRST.Fields("MABL")
           SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
           SDRST.Fields("BANK") = CHRST.Fields("BANK")
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.update
           SDRST.AddNew 'چكهاي پرداختي
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("ف.خ." & HFRST.Fields("NUMBER1") & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
           SDRST.Fields("BED") = CHRST.Fields("MABL")
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.update
           CHRST.MoveNext
       Loop
       Else
       End If
       CHRST.Close
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If JAMF <> 0 Then
           SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("SHARH") = RIGHT("فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##") & " " & HFRST.Fields("MOLAH"), 255)
           SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
           SDRST.Fields("BES") = JAMF + HFRST.Fields("MBAA")
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.Fields("RADIF") = HFRST.Fields("NUMBER")
           SDRST.update
           If KHMAVAV <> 0 Then
               'كنترل خريد '
               SDRST.AddNew 'خريد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
               SDRST.Fields("HES_M") = 1
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-1-1"
               SDRST.Fields("SHARH") = RIGHT("خريد مواد اوليه فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
               SDRST.Fields("BED") = KHMAVAV
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
               SDRST.update
           End If
           If KHNIM <> 0 Then
               'كنترل خريد '
               SDRST.AddNew 'خريد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
               SDRST.Fields("HES_M") = 2
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-2-1"
               SDRST.Fields("SHARH") = RIGHT("خريد نيمه ساخته فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
               SDRST.Fields("BED") = KHNIM
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
               SDRST.update
           End If
           If KHSAKHT <> 0 Then
               'كنترل خريد '
               SDRST.AddNew 'خريد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
               SDRST.Fields("HES_M") = 3
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-3-1"
               SDRST.Fields("SHARH") = RIGHT("خريد ساخته شده فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
               SDRST.Fields("BED") = KHSAKHT
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
               SDRST.update
           End If
           If BAZAR <> 0 Then
               'كنترل خريد '
               SDRST.AddNew 'خريد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
               SDRST.Fields("HES_M") = 4
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-4-1"
               SDRST.Fields("SHARH") = RIGHT("خريد بازرگاني  فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
               SDRST.Fields("BED") = BAZAR
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
               SDRST.update
           End If
           If KHSAY <> 0 Then
                    'كنترل خريد '
                If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), 11, 1) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), 11, 1, "ساير 2")
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-11-1" & " اشكال دارد لطفا بررسي كنيد"
                   End If
                End If
               SDRST.AddNew 'خريد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
               SDRST.Fields("HES_M") = 11
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-11-1"
               SDRST.Fields("SHARH") = RIGHT("خريد ساير فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
               SDRST.Fields("BED") = KHSAY
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
               SDRST.update
           End If
            For K = 1 To 6
                If HS(K) <> 0 Then
                    'كنترل خريد '
                If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), K + 4, 1) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), K + 4, 1, GETGRPKALA(K + 4))
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-" & K + 4 & "-1" & " اشكال دارد لطفا بررسي كنيد"
                   End If
                End If
                    SDRST.AddNew 'خريد
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
                    SDRST.Fields("HES_M") = K + 4
                    SDRST.Fields("HES_T") = 1
                    SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-" & K + 4 & "-1"
                    SDRST.Fields("SHARH") = RIGHT("خريد " & GETGRPKALA(K + 4) & " فاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
                    SDRST.Fields("BED") = HS(K)
                    HS(7) = HS(7) + HS(K)
                    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
                    SDRST.update
                End If
            Next K
           SDRST.AddNew 'پاياپاي خريد
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![PKHARID]
           SDRST.Fields("HES_M") = 1
           SDRST.Fields("HES_T") = 1
           SDRST.Fields("hes") = Forms![BASEKNOW]![PKHARID] & "-1-1"
           SDRST.Fields("SHARH") = RIGHT("خريدفاكتورشماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & "فروشنده: " & GETTAFNAME(HFRST.Fields("CUST_NO")), 255)
           SDRST.Fields("BES") = KHSAY + KHSAKHT + KHNIM + KHMAVAV + BAZAR + HS(7)
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("MABL_HAZ") <> 0 Then
           SDRST.AddNew 'كل بستانكاري شخص بابت خدمات
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BES") = HFRST.Fields("MABL_HAZ")
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("M_NAGHD") <> 0 Then
       SDRST.AddNew 'مبلغ نقدشخص
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = CKOL
       SDRST.Fields("HES_M") = CMOIN
       SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
       SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
       SDRST.Fields("BED") = HFRST.Fields("M_NAGHD")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
  '----------------------ْحواله واريزي
     If HFRST.Fields("MOIN_HAV") <> 0 Then
        If HFRST.Fields("MOIN_HAV") <> 0 Then
           SDRST.AddNew 'مبلغ واريزي شخص
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = CKOL
           SDRST.Fields("HES_M") = CMOIN
           SDRST.Fields("HES_T") = CTAF
           SDRST.Fields("HES_T2") = CTAF2
           SDRST.Fields("HES_T3") = CTAF3
           SDRST.Fields("HES_T4") = CTAF4
           SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
           SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
           SDRST.Fields("BED") = HFRST.Fields("MABL_HAV")
           SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
           SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
           SDRST.update
       End If
       If Not IsNull(HFRST.Fields("MOIN_HAV")) Then
            Call GETTAF3(HFRST.Fields("MOIN_HAV"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            SDRST.AddNew 'مبلغ حواله
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = HFRST.Fields("MOIN_HAV")
            SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
            SDRST.Fields("BEs") = HFRST.Fields("MABL_HAV")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.update
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي مبلغ حواله مشخص نشده است"
        End If
     End If

     If HFRST.Fields("MABL_VAR") <> 0 Then
        SDRST.AddNew 'مبلغ واريزي شخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BEd") = HFRST.Fields("MABL_VAR")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
        SDRST.update
     End If
     If HFRST.Fields("MABL_VAR") <> 0 Then
        If Not IsNull(HFRST.Fields("MOIN_VAR")) Then
             Call GETTAF3(HFRST.Fields("MOIN_VAR"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            SDRST.AddNew 'مبلغ واريزي
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = HFRST.Fields("MOIN_VAR")
            SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
            SDRST.Fields("BEs") = HFRST.Fields("MABL_VAR")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
            SDRST.update
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي مبلغ حواله مشخص نشده است"
        End If
     End If

  '----------------------
  If HFRST.Fields("M_NAGHD") <> 0 Then
       SDRST.AddNew 'مبلغ نقدصندوق
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
       SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
       SDRST.Fields("HES_T") = HFRST.Fields("SHIFT")
       SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & HFRST.Fields("DEPATMAN") & "-" & HFRST.Fields("SHIFT")
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
       SDRST.Fields("BES") = HFRST.Fields("M_NAGHD")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("TAKHFIF") <> 0 Then
       SDRST.AddNew 'مبلغ تخفيف شخص
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = CKOL
       SDRST.Fields("HES_M") = CMOIN
       SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
       SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
       SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
       SDRST.Fields("BED") = HFRST.Fields("TAKHFIF")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("TAKHFIF") <> 0 Then
       SDRST.AddNew 'تخفيف خريد
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = Forms![BASEKNOW]![TKHARID]
       SDRST.Fields("HES_M") = 1
       SDRST.Fields("HES_T") = 1
       SDRST.Fields("hes") = Forms![BASEKNOW]![TKHARID] & "-1-1"
       SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خريد  شماره " & HFRST.Fields("NUMBER1") & "-" & HFRST.Fields("FNUMCO") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
       SDRST.Fields("BES") = HFRST.Fields("TAKHFIF")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
     If HFRST.Fields("MBAA") <> 0 Then
       SDRST.AddNew 'مالليات بر ارزش افزوده
       SDRST.Fields("N_S") = max_ns
       If Not IsNull(HFRST.Fields("HMBAA")) Then
          Call GETTAF3(HFRST.Fields("HMBAA"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
       End If
       SDRST.Fields("HES_K") = HKOL
       SDRST.Fields("HES_M") = HMOIN
       SDRST.Fields("HES_T") = HTAF
       SDRST.Fields("HES_T2") = HTAF2
       SDRST.Fields("HES_T3") = HTAF3
       SDRST.Fields("HES_T4") = HTAF4
       SDRST.Fields("hes") = HFRST.Fields("HMBAA")
       SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور خريد شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
       SDRST.Fields("BED") = HFRST.Fields("MBAA")
       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
       SDRST.Fields("TAG") = 12:  SDRST.Fields("ARZD") = IIf(IsNull(HFRST.Fields("ARZD")), 1, HFRST.Fields("ARZD"))
       SDRST.update
     End If
'      RST.Close
      Set rst = New ADODB.Recordset
      Set JST = New ADODB.Recordset
      HFRST.MoveNext
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Refresh
        Forms![GUG]![Text0].Requery
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
   Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub

Public Function BLOCKEDMK(HES As String) As Boolean
                  On Error Resume Next
    Dim RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, TKOL, TMOIN, TTAF, TTAF2, TTAF3, TTAF4
    Dim HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4
     rst3.Open "SELECT USERCO,HES FROM BLOCKNON_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst3.RecordCount > 0 Then
        While Not rst3.EOF
            Call GETTAF3(rst3.Fields("HES"), TKOL, TMOIN, TTAF, TTAF2, TTAF3, TTAF4)
            Call GETTAF3(HES, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            If TKOL = HKOL Then
                If TKOL = HKOL And IsNull(TMOIN) Then
                     BLOCKEDMK = False
                     Exit Function
                Else
                    If TKOL = HKOL And TMOIN = HMOIN And IsNull(TTAF) Then
                        BLOCKEDMK = False
                        Exit Function
                    Else
                        If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And IsNull(TTAF2) Then
                            BLOCKEDMK = False
                            Exit Function
                        Else
                            If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And IsNull(TTAF3) Then
                                BLOCKEDMK = False
                                Exit Function
                            Else
                                If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And TTAF3 = HTAF3 And IsNull(TTAF4) Then
                                    BLOCKEDMK = False
                                    Exit Function
                                Else
                                    If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And TTAF3 = HTAF3 And TTAF4 = HTAF4 Then
                                        BLOCKEDMK = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            rst3.MoveNext
        Wend
      End If
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            Call GETTAF3(RST2.Fields("HES"), TKOL, TMOIN, TTAF, TTAF2, TTAF3, TTAF4)
         '   TKOL = MGETKOL(rst2.Fields("HES"))
         '   TMOIN = MGETMOIN(rst2.Fields("HES"))
         '   TTAF = MGETTAF(rst2.Fields("HES"))
            Call GETTAF3(HES, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
         '   HKOL = MGETKOL(HES)
         '   HMOIN = MGETMOIN(HES)
         '   HTAF = MGETTAF(HES)
            If TKOL = HKOL Then
                If TKOL = HKOL And IsNull(TMOIN) Then
                     BLOCKEDMK = True
                     Exit Function
                Else
                    If TKOL = HKOL And TMOIN = HMOIN And IsNull(TTAF) Then
                        BLOCKEDMK = True
                        Exit Function
                    Else
                        If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And IsNull(TTAF2) Then
                            BLOCKEDMK = True
                            Exit Function
                        Else
                            If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And IsNull(TTAF3) Then
                                BLOCKEDMK = True
                                Exit Function
                            Else
                                If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And TTAF3 = HTAF3 And IsNull(TTAF4) Then
                                    BLOCKEDMK = True
                                    Exit Function
                                Else
                                    If TKOL = HKOL And TMOIN = HMOIN And TTAF = HTAF And TTAF2 = HTAF2 And TTAF3 = HTAF3 And TTAF4 = HTAF4 Then
                                        BLOCKEDMK = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            RST2.MoveNext
        Wend
      End If
      BLOCKEDMK = False
End Function

Public Function ISTAF(HES As String) As Boolean
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    Dim HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4
    Call GETTAF3(HES, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
    If IsNull(HKOL) Or IsNull(HMOIN) Then
        ISTAF = False
        Exit Function
    End If
    If IsNull(HTAF) Then
        rst.Open "SELECT   TNUMBER FROM dbo.TDETA_HES WHERE (N_KOL = " & HKOL & ") And (NUMBER = " & HMOIN & ") GROUP BY TNUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            ISTAF = True
        Else
            ISTAF = False
        End If
    Else
        If IsNull(HTAF2) Then
            rst.Open "SELECT   TNUMBER2 FROM dbo.TDETA_HES2 WHERE (N_KOL = " & HKOL & ") And (NUMBER = " & HMOIN & ") And (TNUMBER = " & HTAF & ") GROUP BY TNUMBER2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                ISTAF = True
            Else
                ISTAF = False
            End If
        Else
            If IsNull(HTAF3) Then
                rst.Open "SELECT   TNUMBER3 FROM dbo.TDETA_HES3 WHERE (N_KOL = " & HKOL & ") And (NUMBER = " & HMOIN & ") And (TNUMBER = " & HTAF & ") And (TNUMBER2 = " & HTAF2 & ") GROUP BY TNUMBER3", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    ISTAF = True
                Else
                    ISTAF = False
                End If
            Else
                If IsNull(HTAF4) Then
                    rst.Open "SELECT   TNUMBER4 FROM dbo.TDETA_HES4 WHERE (N_KOL = " & HKOL & ") And (NUMBER = " & HMOIN & ") And (TNUMBER = " & HTAF & ") And (TNUMBER2 = " & HTAF2 & ") And (TNUMBER3 = " & HTAF3 & ") GROUP BY TNUMBER4", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        ISTAF = True
                    Else
                        ISTAF = False
                    End If
                Else
                    ISTAF = False
                End If
            End If
        End If
    End If
End Function

Public Function GETGRPKALA(CC As Integer) As String
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     NAMES  FROM dbo.TCOD_STUFGROUP WHERE     (CODE = " & CC & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETGRPKALA = rst.Fields("NAMES")
    End If
End Function


Public Function IFHESAB(HES As String) As Boolean
                     On Error Resume Next
    Dim i As Integer
    For i = 1 To Len(HES)
        If MID(HES, i, 1) = "-" Then
            IFHESAB = True
            Exit Function
        End If
    Next i
    IFHESAB = False
End Function

Public Sub GENSANADANBARGARD(NUMBER As Long, NUMBER2 As Long)
                   On Error Resume Next
 Dim CON, i As Long
 DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
 DoCmd.restore
 DoCmd.OpenForm "GUG"
 Forms![GUG]![Text0] = "n"
 Forms![GUG]![Text0].Visible = True
 Forms![GUG]![Box7].Visible = True
 Forms![GUG].Form.Refresh
 Forms![GUG].Form.Repaint
  Dim HEDRST As New ADODB.Recordset, rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
  Dim NEWR As Boolean
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.Open "SELECT     GRD_NUM, GRD_DATE, GRD_ANBAR, GRD_HES, N_S, COMMENT, USER_NAME FROM     dbo.ANBGRD_HEAD WHERE ((GRD_NUM >= " & NUMBER & " AND GRD_NUM <=" & NUMBER2 & " ) and not(N_S IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HEDRST.MoveLast
  HEDRST.MoveFirst
  Forms![GUG]![SNUM] = HEDRST.RecordCount
   CON = HEDRST.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
   i = 1
  While Not HEDRST.EOF()
      Forms![GUG]![num] = i
        rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Then
          max_ns = 1
        Else
          max_ns = rst.Fields(0) + 1
        End If
        Err.Clear
        SHART = "NO_S = 17 AND N_S = " & HEDRST.Fields("N_S")
        SHRST.Filter = SHART
        Set SHRST = SHRST
        If SHRST.RecordCount = 0 Or Err <> 0 Then
          Err.Clear
          SHRST.AddNew
          SHRST.Fields("N_S") = max_ns
          SHRST.Fields("DATE_S") = HEDRST.Fields("GRD_DATE")
          SHRST.Fields("SHARH_S") = Left(" انبار گرداني شماره " & HEDRST.Fields("GRD_NUM") & " از انبار " & HEDRST.Fields("GRD_ANBAR") & " مورخ " & Format(HEDRST.Fields("GRD_DATE"), "####/##/##"), 255)
          SHRST.Fields("GHATEI") = 0
          SHRST.Fields("NO_S") = 17
          SHRST.Fields("OKF") = True
          SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
          SHRST.Fields("CRT") = Now()
          SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
          SHRST.update
        Else
          max_ns = SHRST.Fields("N_S")
          SHRST.Fields("DATE_S") = HEDRST.Fields("GRD_DATE")
          SHRST.Fields("SHARH_S") = Left(" انبار گرداني شماره " & HEDRST.Fields("GRD_NUM") & " از انبار " & HEDRST.Fields("GRD_ANBAR") & " مورخ " & Format(HEDRST.Fields("GRD_DATE"), "####/##/##"), 255)
          SHRST.Fields("GHATEI") = 0
          SHRST.Fields("NO_S") = 17
          SHRST.Fields("OKF") = True
          SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
          SHRST.update
        End If
        If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
           HEDRST.Fields("N_S") = max_ns
           HEDRST.update
        End If
        Set JST = New ADODB.Recordset
        If Not IsNull(HEDRST.Fields("GRD_HES")) Then
           Call GETTAF3(HEDRST.Fields("GRD_HES"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
        End If
        If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
           HEDRST.Fields("N_S") = max_ns
        End If
        DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE N_S = " & HEDRST.Fields("N_S"))
        JST.Open "SELECT  dbo.ANBGRD_LST.*, MOG - NUM3 AS EKH FROM dbo.ANBGRD_LST WHERE  (MOG - NUM2 <> 0) AND (MOG - NUM1 <> 0) AND GRD_NUM = " & HEDRST.Fields("GRD_NUM"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        JAMF = 0
        While Not JST.EOF()
              If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), HEDRST.Fields("GRD_ANBAR"), CLng(JST.Fields("CODE"))) Then
                 Err.Clear
                 Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), HEDRST.Fields("GRD_ANBAR"), CLng(JST.Fields("CODE")), JST.Fields("CODE"))
                 If Err <> 0 Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                 End If
              End If
              If Round(LASTAVRAGE(JST.Fields("CODE"), HEDRST.Fields("GRD_ANBAR"), HEDRST.Fields("GRD_DATE")) * JST.Fields("EKH")) <> 0 Then
                  SDRST.AddNew 'از انبار
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                  SDRST.Fields("HES_M") = HEDRST.Fields("GRD_ANBAR")
                  SDRST.Fields("HES_T") = JST.Fields("CODE")
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & HEDRST.Fields("GRD_ANBAR") & "-" & JST.Fields("CODE")
                  If JST.Fields("EKH") > 0 Then
                      SDRST.Fields("SHARH") = Left(" انبار گرداني شماره " & HEDRST.Fields("GRD_NUM") & " از انبار " & HEDRST.Fields("GRD_ANBAR") & " مورخ " & Format(HEDRST.Fields("GRD_DATE"), "####/##/##") & " به مقدار" & JST.Fields("EKH"), 255)
                      SDRST.Fields("BES") = Round(JST.Fields("mabl") * JST.Fields("EKH"))
                  Else
                      SDRST.Fields("SHARH") = Left(" انبار گرداني شماره " & HEDRST.Fields("GRD_NUM") & " از انبار " & HEDRST.Fields("GRD_ANBAR") & " مورخ " & Format(HEDRST.Fields("GRD_DATE"), "####/##/##") & " به مقدار" & JST.Fields("EKH") * -1, 255)
                      SDRST.Fields("BED") = Round(JST.Fields("mabl") * JST.Fields("EKH") * -1)
                  End If
                  SDRST.update
              End If
              'jst.Fields("MABL") = LASTAVRAGE(jst.Fields("CODE"), HEDRST.Fields("GRD_ANBAR"), HEDRST.Fields("GRD_DATE"))
              'jst.update
              JAMF = JAMF + Round(JST.Fields("mabl") * JST.Fields("EKH"))
              JST.MoveNext
            Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
            Forms![GUG].Form.Refresh
            Forms![GUG]![Text2].Requery
            Forms![GUG].Form.Repaint
        Wend
        If JAMF <> 0 Then
              SDRST.AddNew 'به انبار
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = CKOL
              SDRST.Fields("HES_M") = CMOIN
              SDRST.Fields("HES_T") = CTAF
              SDRST.Fields("HES_T2") = CTAF2
              SDRST.Fields("HES_T3") = CTAF3
              SDRST.Fields("HES_T4") = CTAF4
              SDRST.Fields("hes") = HEDRST.Fields("GRD_HES")
              SDRST.Fields("SHARH") = Left("انبار گرداني شماره " & HEDRST.Fields("GRD_NUM") & " از انبار " & HEDRST.Fields("GRD_ANBAR") & " مورخ " & Format(HEDRST.Fields("GRD_DATE"), "####/##/##"), 255)
              If JAMF > 0 Then
                  SDRST.Fields("BED") = JAMF
              Else
                  SDRST.Fields("BES") = JAMF * -1
              End If
              SDRST.update
        End If
        JAMF = 0
        HEDRST.MoveNext
        If (i Mod CON) = 0 Then
          Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
          Forms![GUG].Form.Refresh
          Forms![GUG]![Text0].Requery
          Forms![GUG].Form.Repaint
        End If
        i = i + 1
        Forms![GUG]![Text2] = ""
        Set rst = New ADODB.Recordset
    Wend
   DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
End Sub
Public Sub gensanadbargashfroosh(fnum As Long, TNUM As Long)
                On Error Resume Next
  Dim CON, i As Long, a, fs
  DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
  DoCmd.OpenForm "GUG"
  Forms![GUG]![Text0] = "n"
  Forms![GUG]![Text0].Visible = True
  Forms![GUG]![Box7].Visible = True
  Forms![GUG].Form.Refresh
  Forms![GUG].Form.Repaint
  Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, JAMFKH As Double, MBL As Double, CMOIN As Double, CTAF As Double, takh As Double, SHART As String
  Dim ii As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long, PRST As New ADODB.Recordset, TAKHF As Double, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, MAVAD As Double, DAST As Double, SAR As Double
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HFRST.Open "SELECT     * FROM dbo.HEAD_LST WHERE     (NUMBER BETWEEN " & fnum & " AND " & TNUM & ") AND (TAG = 4)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HFRST.MoveLast
  HFRST.MoveFirst
  Forms![GUG]![SNUM] = HFRST.RecordCount
  CON = HFRST.RecordCount / 50
  If CON < 1 Then
    CON = 1
  End If
  i = 1
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OPENTextFile("c:\LOGbrfr.txt", 8, 1)
    If Err <> 0 Then
        Set a = fs.CREATETextFile("c:\LOGbrfr.txt", True)
    End If
    a.writeline ("شروع باز سازي از فاكتور شماره : " & fnum & " تا فاكتور شماره :" & TNUM)
    a.writeline (Now())
  While Not HFRST.EOF
      Forms![GUG]![num] = i
      If Not IsNull(HFRST.Fields("CUST_NO")) Then
           Call GETTAF3(HFRST.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
      rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If IsNull(rst.Fields(0)) Then
        max_ns = 1
      Else
        max_ns = rst.Fields(0) + 1
      End If
      Err.Clear
      SHART = "NO_S = 4 AND N_S = " & HFRST.Fields("N_S")
      SHRST.Filter = SHART: Set SHRST = SHRST
      If SHRST.RecordCount = 0 Or Err <> 0 Then
        Err.Clear
        SHRST.AddNew
        SHRST.Fields("N_S") = max_ns
        SHRST.Fields("DATE_S") = HFRST.Fields("DATE_N")
        SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 4
        SHRST.Fields("OKF") = True
        SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
        SHRST.update
      Else
         If IsNull(HFRST.Fields("N_S")) Then
            max_ns = SHRST.Fields("N_S")
          Else
            max_ns = HFRST.Fields("N_S")
          End If
      End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      If IsNull(HFRST.Fields("N_S")) Or HFRST.Fields("N_S") <> max_ns Then
         HFRST.Fields("N_S") = max_ns
         HFRST.update
      End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
   JST.Open "SELECT Sum([MEGH_MAR]*[mabl]) AS mabk  FROM dbo.INVO_LST WHERE     (NUMBER = " & HFRST.Fields("NUMBER1") & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
      JAMF = JST.Fields(0)
   Else
      JAMF = 0
   End If
  '  Set JST = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   JST.Open "SELECT Sum([MEGH_MAR]*[mabl]) AS mabk FROM dbo.INVO_LST WHERE     (NUMBER = " & HFRST.Fields("NUMBER1") & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
      JAMFKH = JST.Fields(0)
   Else
      JAMFKH = 0
   End If
'    Set JST = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=4) AND ((PAY_GETD.NUMBER)= " & HFRST.Fields("NUMBER") & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
   Else
        JAMCH = 0
  End If
   TAKHF = 0
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HFRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= " & 4 & "))")
   Set JST = New ADODB.Recordset
   JST.Open "SELECT  dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.MABL AS MABL_K, dbo.INVO_LST_TAKH.MEGH_MAR, dbo.INVO_LST_TAKH.CODE, dbo.INVO_LST_TAKH.ANBAR, dbo.STUF_DEF.NAME, dbo.INVO_LST_TAKH.CUST_KIND, ISNULL(dbo.TAKHPERS.TAFPER, 0) AS TFP, ROUND(ISNULL(dbo.TAKHPERS.TAFPER, 0) * dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.mabl / 100, 0) As takh,dbo.INVO_LST_TAKH.AVRAGE   FROM  dbo.STUF_DEF INNER JOIN  dbo.INVO_LST_TAKH ON dbo.STUF_DEF.CODE = dbo.INVO_LST_TAKH.CODE LEFT OUTER JOIN   dbo.TAKHPERS ON dbo.INVO_LST_TAKH.CUST_KIND = dbo.TAKHPERS.CUST_CO AND    dbo.INVO_LST_TAKH.CODE = dbo.TAKHPERS.TAKH_COD WHERE     (dbo.INVO_LST_TAKH.MEGH_MAR  <> 0) AND  (dbo.INVO_LST_TAKH.NUMBER = " & HFRST.Fields("NUMBER1") & ") AND (dbo.INVO_LST_TAKH.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            SDRST.AddNew 'فروش
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
            SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-4-" & CLng(JST.Fields("CODE"))
            SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
            SDRST.Fields("BED") = JST.Fields("MABL_K")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("TAG") = 4
       Else
        If JST.Fields("anbar") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            If JST.Fields("MABL_K") > 0 Then
                SDRST.AddNew 'فروش
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]

                If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                  SDRST.Fields("HES_M") = 1
                  SDRST.Fields("HES_T") = 1
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-1-1"
                Else
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                End If
                SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                SDRST.Fields("BED") = JST.Fields("MABL_K")
                SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 4
                SDRST.update
             End If
        Else
            If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code"))) Then
              Err.Clear
              Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code")), JST.Fields(4))
              If Err <> 0 Then
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              End If
               End If
               SDRST.AddNew 'در آمد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
               SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
               SDRST.Fields("HES_T") = JST.Fields("CODE")
               SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & HFRST.Fields("DEPATMAN") & "-" & JST.Fields("CODE")
               SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
               SDRST.Fields("BEd") = JST.Fields("MABL_K")
'               jamkh = jamkh + JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 4
              SDRST.update
          End If

      End If

       If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or True Then
                  MAVAD = Round(GETSTANDARDPRICE_MAVAD(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGH_MAR"))
                  DAST = Round(GETSTANDARDPRICE_DAST(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGH_MAR"))
                  SAR = Round(GETSTANDARDPRICE_SAR(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGH_MAR"))
                  If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                     End If
                  End If
                  If (MAVAD + DAST + SAR) <> 0 And MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) <> "5" Then
                      SDRST.AddNew 'انبار محصول
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                      SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                      SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                      SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                      SDRST.Fields("BED") = MAVAD + DAST + SAR
                      SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                      SDRST.Fields("TAG") = 4
                      SDRST.update
                        If MID(Forms![BASEKNOW]![tindata], 9, 1) <> 1 Then
                          If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                             Err.Clear
                             Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                             If Err <> 0 Then
                                DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                             End If
                          End If
                        End If
                      If MAVAD > 0 Then
                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = MAVAD
                          SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 4
                          SDRST.update
                      End If
                     If DAST <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999, "دستمزد " & JST.Fields(4))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                          End If
                       End If
                       SDRST.AddNew 'قيمت تمام شده
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                       SDRST.Fields("HES_T") = 9999999
                       SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                       SDRST.Fields("BES") = DAST
                       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 4
                       SDRST.update
                     End If
                     If SAR <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998, "سربار " & JST.Fields(4))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                          End If
                       End If
                       SDRST.AddNew 'قيمت تمام شده
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                       SDRST.Fields("HES_T") = 9999998
                       SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                       SDRST.Fields("BES") = SAR
                       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 4
                       SDRST.update
                    End If
                Else
                      If JST.Fields("avrage") > 0 Then
                            SDRST.AddNew 'انبار محصول
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                            SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGH_MAR"))
                            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                            SDRST.Fields("TAG") = 4
                            SDRST.update

                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGH_MAR"))
                          SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 4
                          SDRST.update
                      End If
                End If
        End If
     JST.MoveNext
   Wend

 '   Set JST = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGH_MAR, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, [SumOfMABLK]+[IMBIBE_MANF]+[IMBIBE_SAR] AS GHT FROM (INVO_LST INNER JOIN MANF_JAMK ON INVO_LST.CODE = MANF_JAMK.CODE) INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (((INVO_LST.NUMBER)=" & HFRST.Fields("NUMBER1") & ") AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If HFRST.Fields("MABL_HAZ") <> 0 Then
    If IsNull(HFRST.Fields("MOIN_HAZ")) Then
           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل يا غيره
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("MOIN_HAZ")) Then
             Call GETTAF3(HFRST.Fields("MOIN_HAZ"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("MOIN_HAZ")
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش  شماره " & HFRST.Fields("NUMBER") & " - " & GETTAFNAME(HFRST.Fields("MOIN_HAZ")), 255)
        SDRST.Fields("BED") = HFRST.Fields("MABL_HAZ")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.update
     End If
  End If
  If JAMCH <> 0 Then 'چكهاي پرداختي
    CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER = " & HFRST.Fields("NUMBER") & " AND TAG = 4"
    Set CHRST = CHRST
    If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
    Do Until CHRST.EOF
        MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
        SDRST.AddNew 'اسناد پرداختني
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![APA])
        SDRST.Fields("hes") = Forms![BASEKNOW]![APA]
        SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BES") = CHRST.Fields("MABL")
        SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
        SDRST.Fields("BANK") = CHRST.Fields("BANK")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.update
        SDRST.AddNew 'چكهاي پرداختي
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("ف.ب.ف." & HFRST.Fields("NUMBER") & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BED") = CHRST.Fields("MABL")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.update
        CHRST.MoveNext
    Loop
    Else
    End If
  End If
    If HFRST.Fields("TAKHFIF") <> 0 Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM(dbo.INVO_LST.n_moin / dbo.INVO_LST.meghk * dbo.INVO_LST.MEGH_MAR) AS JAMT, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.INVO_LST.NUMBER = " & HFRST.Fields("NUMBER1") & " ) And (dbo.INVO_LST.TAG = 2) GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          TAKHF = 0
          While Not rst.EOF()
             If Round(rst.Fields("JAMT")) <> 0 Then
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        End If
                    End If
                     SDRST.AddNew 'تخفيف فروش
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                     SDRST.Fields("HES_M") = 3
                     SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش فاكتور  شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                     SDRST.Fields("TAG") = 4
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                 Else
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE"))) Then
                      Err.Clear
                      Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                      End If
                     End If
                     SDRST.AddNew 'تخفيف فروش
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                     SDRST.Fields("HES_M") = CLng(rst.Fields("CUST_KIND"))
                     SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش فاكتور  شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & CLng(rst.Fields("CUST_KIND")) & "-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                     SDRST.Fields("TAG") = 4
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                  End If
                End If
                rst.MoveNext
           Wend
        End If
        If HFRST.Fields("TAKHFIF") <> TAKHF Then
           HFRST.Fields("TAKHFIF") = TAKHF
        End If
   End If
      If HFRST.Fields("MBAA") <> 0 Then
        SDRST.AddNew 'مالليات بر ارزش افزوده
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("MBAA")) Then
             Call GETTAF3(HFRST.Fields("HMBAA"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("HMBAA")
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BED") = HFRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.update
     End If
  If JAMF + HFRST.Fields("MABL_HAZ") - HFRST.Fields("TAKHFIF") > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش  شماره " & HFRST.Fields("NUMBER") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##") & HFRST.Fields("MOLAH"), 255)
        SDRST.Fields("BES") = JAMF + HFRST.Fields("MABL_HAZ") - HFRST.Fields("TAKHFIF") + HFRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.Fields("RADIF") = HFRST.Fields("NUMBER")
        SDRST.update
   End If
'  If HFRST.Fields("MABL_HAZ") <> 0 Then
'    SDRST.AddNew 'كل بستانكاري شخص بابت خدمات  فاكتور
'    SDRST.Fields("N_S") = MAX_NS
'    SDRST.Fields("HES_K") = CKOL
'    SDRST.Fields("HES_M") = CMOIN
'    SDRST.Fields("HES_t") = CTAF
'    SDRST.Fields("SHARH") = Right("خدمات فاكتور برگشت فروش  شماره " & HFRST.Fields("NUMBER") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
'    SDRST.Fields("BES") = HFRST.Fields("MABL_HAZ")
'    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
'    SDRST.Fields("TAG") = 4
'    SDRST.update
'    End If
  If HFRST.Fields("M_NAGHD") <> 0 Then
    SDRST.AddNew 'مبلغ نقدشخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_t") = CTAF
    SDRST.Fields("HES_T2") = CTAF2
    SDRST.Fields("HES_T3") = CTAF3
    SDRST.Fields("HES_T4") = CTAF4
    SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش  شماره " & HFRST.Fields("NUMBER") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
    SDRST.Fields("BED") = HFRST.Fields("M_NAGHD")
    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
    SDRST.Fields("TAG") = 4
    SDRST.update
  End If
  If HFRST.Fields("M_NAGHD") <> 0 Then
    SDRST.AddNew 'مبلغ نقدصندوق
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
    SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
    SDRST.Fields("HES_T") = HFRST.Fields("SHIFT")
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & HFRST.Fields("DEPATMAN") & "-" & HFRST.Fields("SHIFT")
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش  شماره " & HFRST.Fields("NUMBER") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
    SDRST.Fields("BES") = HFRST.Fields("M_NAGHD")
    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
    SDRST.Fields("TAG") = 4
    SDRST.update
  End If
  Dim visitorn As String
  JAMP = 0
  PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & HFRST.Fields("NUMBER") & ") AND (TAG = 4) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not PRST.EOF
        visitorn = GETTAFNAME(PRST.Fields("CUST_NO"))
        Dim PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4
        If Not IsNull(PRST.Fields("CUST_NO")) Then
             Call GETTAF3(PRST.Fields("CUST_NO"), PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4)
        End If
        SDRST.AddNew 'پورسانت
        SDRST.Fields("HES_K") = PHKOL
        SDRST.Fields("HES_M") = PHMOIN
        SDRST.Fields("HES_T") = PHTAF
        SDRST.Fields("HES_T2") = PHTAF2
        SDRST.Fields("HES_T3") = PHTAF3
        SDRST.Fields("HES_T4") = PHTAF4
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("hes") = PRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT(" فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER") & " بابت " & PRST.Fields("DARSAD") & "درصد سهم پورسانت " & GETTAFNAME(PRST.Fields("CUST_NO")) & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")), 255)
        If Not PRST.Fields("STAT") Then
            If Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100) <> PRST.Fields("PURSANT") Then
                PRST.Fields("PURSANT") = Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100)
                PRST.update
            End If
        Else
            If PRST.Fields("DARSAD") <> PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100 Then
                PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100
                PRST.update
            End If
        End If
        SDRST.Fields("BEd") = PRST.Fields("PURSANT")
        JAMP = JAMP + PRST.Fields("PURSANT")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        If PRST.Fields("PURSANT") <> 0 Then
            SDRST.update
        End If
        PRST.MoveNext
    Wend
    If JAMP > 0 Then
        SDRST.AddNew 'پورسانت
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("hes") = Forms![BASEKNOW]![HPOR]
        SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER") & " " & visitorn, 255)
        SDRST.Fields("BES") = JAMP
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 4
        SDRST.update
    End If

    Set PRST = New ADODB.Recordset
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      rst.Close
      Set rst = New ADODB.Recordset
      Set JST = New ADODB.Recordset
      HFRST.MoveNext
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Refresh
        Forms![GUG]![Text0].Requery
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
   Wend
   Err.Clear
     a.writeline (Now())
  DoCmd.Close acForm, "GUG"
   DoCmd.Close acForm, "GENSANADFROOSH"
   Call gensanadbargashfroosh2(fnum, TNUM)
End Sub

Public Sub gensanadbargashfroosh2(fnum As Long, TNUM As Long)
                  On Error Resume Next
  Dim CON, i As Long, a, fs
 ' DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
  DoCmd.OpenForm "GUG"
  Forms![GUG]![Text0] = "n"
  Forms![GUG]![Text0].Visible = True
  Forms![GUG]![Box7].Visible = True
  Forms![GUG].Form.Refresh
  Forms![GUG].Form.Repaint
  Dim rst As New ADODB.Recordset, HFRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, JAMFKH As Double, MBL As Double, CMOIN As Double, CTAF As Double, takh As Double, SHART As String
  Dim ii As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long, PRST As New ADODB.Recordset, TAKHF As Double, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, MAVAD As Double, DAST As Double, SAR As Double
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HFRST.Open "SELECT     * FROM dbo.HEAD_LST WHERE     (NUMBER BETWEEN " & fnum & " AND " & TNUM & ") AND (TAG = 25)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  HFRST.MoveLast
  HFRST.MoveFirst
  Forms![GUG]![SNUM] = HFRST.RecordCount
  CON = HFRST.RecordCount / 50
  If CON < 1 Then
    CON = 1
  End If
  i = 1
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OPENTextFile("c:\LOGbrfr2.txt", 8, 1)
    If Err <> 0 Then
        Set a = fs.CREATETextFile("c:\LOGbrfr2.txt", True)
    End If
    a.writeline ("شروع باز سازي از فاكتور شماره : " & fnum & " تا فاكتور شماره :" & TNUM)
    a.writeline (Now())
  While Not HFRST.EOF
      Forms![GUG]![num] = i
        If Not IsNull(HFRST.Fields("CUST_NO")) Then
           Call GETTAF3(HFRST.Fields("CUST_NO"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
        End If
      rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If IsNull(rst.Fields(0)) Then
        max_ns = 1
      Else
        max_ns = rst.Fields(0) + 1
      End If
      Err.Clear
      SHART = "NO_S = 4 AND N_S = " & HFRST.Fields("N_S")
      SHRST.Filter = SHART: Set SHRST = SHRST
      If SHRST.RecordCount = 0 Or Err <> 0 Then
        Err.Clear
        SHRST.AddNew
        SHRST.Fields("N_S") = max_ns
        SHRST.Fields("DATE_S") = HFRST.Fields("DATE_N")
        SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 4
        SHRST.Fields("OKF") = True
        SHRST.Fields("USER_NAME") = HFRST.Fields("USER_NAME")
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
        SHRST.update
      Else
         If IsNull(HFRST.Fields("N_S")) Then
            max_ns = SHRST.Fields("N_S")
          Else
            max_ns = HFRST.Fields("N_S")
          End If
      End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      If IsNull(HFRST.Fields("N_S")) Or HFRST.Fields("N_S") <> max_ns Then
         HFRST.Fields("N_S") = max_ns
         HFRST.update
      End If
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE ((INVO_LST.NUMBER)= " & HFRST.Fields("NUMBER") & ") AND (TAG = 24) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
      JAMF = JST.Fields(0)
   Else
      JAMF = 0
   End If
  '  Set JST = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE ((INVO_LST.NUMBER)= " & HFRST.Fields("NUMBER") & ") AND (TAG = 24) AND (ANBAR = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
      JAMFKH = JST.Fields(0)
   Else
      JAMFKH = 0
   End If
 '   Set JST = New ADODB.Recordset
   Set JST = New ADODB.Recordset
   JST.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.TAG)=24) AND ((PAY_GETP.NUMBER)= " & HFRST.Fields("NUMBER") & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
   Else
        JAMCH = 0
  End If
   TAKHF = 0
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (DEED_DTL.NUMBER= " & HFRST.Fields("NUMBER") & ") AND (DEED_DTL.TAG= 24 OR DEED_DTL.TAG= 25)")

   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, INVO_LST.avrage fROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & HFRST.Fields("NUMBER") & ") AND ((INVO_LST.TAG)=24)) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            SDRST.AddNew 'فروش
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
            SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-4-" & CLng(JST.Fields("CODE"))
            SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
            SDRST.Fields("BED") = JST.Fields("MABL_K")
            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
            SDRST.Fields("TAG") = 25
       Else
        If JST.Fields("anbar") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش. وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            If JST.Fields("MABL_K") > 0 Then
                SDRST.AddNew 'فروش
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]

                If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                  SDRST.Fields("HES_M") = 1
                  SDRST.Fields("HES_T") = 1
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-1-1"
                Else
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                End If
                SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                SDRST.Fields("BED") = JST.Fields("MABL_K")
                SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 25
                SDRST.update
             End If
        Else
            If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code"))) Then
              Err.Clear
              Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(HFRST.Fields("DEPATMAN")), CLng(JST.Fields("code")), JST.Fields(4))
              If Err <> 0 Then
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              End If
               End If
               SDRST.AddNew 'در آمد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
               SDRST.Fields("HES_M") = HFRST.Fields("DEPATMAN")
               SDRST.Fields("HES_T") = JST.Fields("CODE")
               SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & HFRST.Fields("DEPATMAN") & "-" & JST.Fields("CODE")
               SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
               SDRST.Fields("BEd") = JST.Fields("MABL_K")
'               jamkh = jamkh + JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
               SDRST.Fields("TAG") = 25
              SDRST.update
          End If

      End If
       If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or True Then
                  MAVAD = Round(GETSTANDARDPRICE_MAVAD(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
                  DAST = Round(GETSTANDARDPRICE_DAST(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
                  SAR = Round(GETSTANDARDPRICE_SAR(JST.Fields("code"), HFRST.Fields("DATE_N")) * JST.Fields("MEGHK"))
                  If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                     End If
                  End If
                  If (MAVAD + DAST + SAR) <> 0 And MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) <> "5" Then
                      SDRST.AddNew 'انبار محصول
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                      SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                      SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                      SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                      SDRST.Fields("BED") = MAVAD + DAST + SAR
                      SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                      SDRST.Fields("TAG") = 25
                      SDRST.update
                        If MID(Forms![BASEKNOW]![tindata], 9, 1) <> 1 Then
                          If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                             Err.Clear
                             Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                             If Err <> 0 Then
                                DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                             End If
                          End If
                        End If
                      If MAVAD > 0 Then
                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = MAVAD
                          SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 25
                          SDRST.update
                      End If
                     If DAST <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999, "دستمزد " & JST.Fields(4))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                          End If
                       End If
                       SDRST.AddNew 'قيمت تمام شده
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                       SDRST.Fields("HES_T") = 9999999
                       SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                       SDRST.Fields("BES") = DAST
                       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 25
                       SDRST.update
                     End If
                     If SAR <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998, "سربار " & JST.Fields(4))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."

                          End If
                       End If
                       SDRST.AddNew 'قيمت تمام شده
                       SDRST.Fields("N_S") = max_ns
                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                       SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                       SDRST.Fields("HES_T") = 9999998
                       SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                       SDRST.Fields("BES") = SAR
                       SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                       SDRST.Fields("TAG") = 25
                       SDRST.update
                    End If
                Else
                      If JST.Fields("avrage") > 0 Then
                            SDRST.AddNew 'انبار محصول
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                            SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("برگشت. فروش  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                            SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                            SDRST.Fields("TAG") = 25
                            SDRST.update

                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & HFRST.Fields("NUMBER") & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                          SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                          SDRST.Fields("TAG") = 25
                          SDRST.update
                      End If
                End If
        End If
     JST.MoveNext
   Wend
   Set JST = New ADODB.Recordset
  If HFRST.Fields("MABL_HAZ") <> 0 Then
    If IsNull(HFRST.Fields("MOIN_HAZ")) Then
     '      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل يا غيره
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("MOIN_HAZ")) Then
             Call GETTAF3(HFRST.Fields("MOIN_HAZ"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("MOIN_HAZ")
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش.  شماره" & HFRST.Fields("NUMBER") & "- " & GETTAFNAME(HFRST.Fields("MOIN_HAZ")), 255)
        SDRST.Fields("BED") = HFRST.Fields("MABL_HAZ")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.update
     End If
  End If
  If JAMCH <> 0 Then 'چكهاي پرداختي
    CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER =" & HFRST.Fields("NUMBER") & "AND TAG = " & 24
    Set CHRST = CHRST
    If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
    Do Until CHRST.EOF
        MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
        SDRST.AddNew 'اسناد پرداختني
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![APA])
        SDRST.Fields("hes") = Forms![BASEKNOW]![APA]
        SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BES") = CHRST.Fields("MABL")
        SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
        SDRST.Fields("BANK") = CHRST.Fields("BANK")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.update
        SDRST.AddNew 'چكهاي پرداختي
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("ف.ب.ف." & HFRST.Fields("NUMBER") & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BED") = CHRST.Fields("MABL")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.update
        CHRST.MoveNext
    Loop
    End If
  End If
    If HFRST.Fields("TAKHFIF") <> 0 Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM(dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGHk / 100) AS JAMT, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG-1 WHERE (dbo.INVO_LST.NUMBER =" & HFRST.Fields("NUMBER") & ") And (dbo.INVO_LST.TAG = 24) GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          TAKHF = 0
          While Not rst.EOF()
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        End If
                    End If
                     SDRST.AddNew 'تخفيف فروش
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                     SDRST.Fields("HES_M") = 3
                     SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش. فاكتور  شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                     SDRST.Fields("TAG") = 25
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                 Else
                    If Round(rst.Fields("JAMT")) <> 0 Then
                         If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE"))) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                         End If
                         SDRST.AddNew 'تخفيف فروش
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                         SDRST.Fields("HES_M") = CLng(rst.Fields("CUST_KIND"))
                         SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                         SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش. فاكتور  شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & CLng(rst.Fields("CUST_KIND")) & "-" & CLng(rst.Fields("CODE"))
                         SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                         SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
                         SDRST.Fields("TAG") = 25
                         TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                         SDRST.update
                    End If
                 End If
                rst.MoveNext
           Wend
        End If
        If HFRST.Fields("TAKHFIF") <> TAKHF Then
            HFRST.Fields("TAKHFIF") = TAKHF
        End If
   End If
      If HFRST.Fields("MBAA") <> 0 Then
        SDRST.AddNew 'مالليات بر ارزش افزوده
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(HFRST.Fields("HMBAA")) Then
             Call GETTAF3(HFRST.Fields("HMBAA"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = HFRST.Fields("HMBAA")
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور برگشت فروش شماره " & HFRST.Fields("NUMBER1") & " مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
        SDRST.Fields("BED") = HFRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.update
     End If
  If JAMF + HFRST.Fields("MABL_HAZ") - HFRST.Fields("TAKHFIF") + HFRST.Fields("MBAA") > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش.  شماره" & HFRST.Fields("NUMBER") & "مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##") & HFRST.Fields("MOLAH"), 255)
        SDRST.Fields("BES") = JAMF + HFRST.Fields("MABL_HAZ") - HFRST.Fields("TAKHFIF") + HFRST.Fields("MBAA")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.Fields("RADIF") = HFRST.Fields("NUMBER")
        SDRST.update
   End If

  If HFRST.Fields("M_NAGHD") <> 0 Then
    SDRST.AddNew 'مبلغ نقدشخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_t") = CTAF
    SDRST.Fields("HES_T2") = CTAF2
    SDRST.Fields("HES_T3") = CTAF3
    SDRST.Fields("HES_T4") = CTAF4
    SDRST.Fields("hes") = HFRST.Fields("CUST_NO")
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش.  شماره" & HFRST.Fields("NUMBER") & "مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
    SDRST.Fields("BED") = HFRST.Fields("M_NAGHD")
    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
    SDRST.Fields("TAG") = 25
    SDRST.update
  End If
  If HFRST.Fields("M_NAGHD") <> 0 Then
    SDRST.AddNew 'مبلغ نقدصندوق
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
    SDRST.Fields("HES_M") = 1
    SDRST.Fields("HES_T") = 1
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-1-1"
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش.  شماره" & HFRST.Fields("NUMBER") & "مورخ" & Format(HFRST.Fields("DATE_N"), "####/##/##"), 255)
    SDRST.Fields("BES") = HFRST.Fields("M_NAGHD")
    SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
    SDRST.Fields("TAG") = 25
    SDRST.update
  End If
  JAMP = 0
  Dim visitorn As String
    PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & HFRST.Fields("NUMBER") & ") AND (TAG = 24) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not PRST.EOF
        visitorn = GETTAFNAME(PRST.Fields("CUST_NO"))
        Dim PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4
        If Not IsNull(PRST.Fields("CUST_NO")) Then
             Call GETTAF3(PRST.Fields("CUST_NO"), PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4)
        End If
        SDRST.AddNew 'پورسانت
        SDRST.Fields("HES_K") = PHKOL
        SDRST.Fields("HES_M") = PHMOIN
        SDRST.Fields("HES_T") = PHTAF
        SDRST.Fields("HES_T2") = PHTAF2
        SDRST.Fields("HES_T3") = PHTAF3
        SDRST.Fields("HES_T4") = PHTAF4
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("hes") = PRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT(" فاكتور برگشت فروش . شماره" & HFRST.Fields("NUMBER") & "بابت " & PRST.Fields("DARSAD") & "درصد سهم پورسانت " & GETTAFNAME(PRST.Fields("CUST_NO")) & " مورخ " & Format(HFRST.Fields("DATE_N"), "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")), 255)
        If Not PRST.Fields("STAT") Then
            If Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100) <> PRST.Fields("PURSANT") Then
                PRST.Fields("PURSANT") = Round((JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * PRST.Fields("DARSAD") / 100)
                PRST.update
            End If
        Else
            If PRST.Fields("DARSAD") <> PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100 Then
                PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / (JAMF - HFRST.Fields("TAKHFIF") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, HFRST.Fields("mbaa"), 0)) * 100
                PRST.update
            End If
        End If
        SDRST.Fields("BEd") = PRST.Fields("PURSANT")
        JAMP = JAMP + PRST.Fields("PURSANT")
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        If PRST.Fields("PURSANT") <> 0 Then
            SDRST.update
        End If
        PRST.MoveNext
    Wend
    If JAMP > 0 Then
        SDRST.AddNew 'پورسانت
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("hes") = Forms![BASEKNOW]![HPOR]
        SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور برگشت فروش . شماره" & HFRST.Fields("NUMBER") & "" & visitorn, 255)
        SDRST.Fields("BES") = JAMP
        SDRST.Fields("NUMBER") = HFRST.Fields("NUMBER")
        SDRST.Fields("TAG") = 25
        SDRST.update
    End If
    Set PRST = New ADODB.Recordset
      Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
      Forms![GUG].Form.Refresh
      Forms![GUG]![Text2].Requery
      Forms![GUG].Form.Repaint
      rst.Close
      Set rst = New ADODB.Recordset
      Set JST = New ADODB.Recordset
      HFRST.MoveNext
      If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Refresh
        Forms![GUG]![Text0].Requery
        Forms![GUG].Form.Repaint
      End If
      i = i + 1
      Forms![GUG]![Text2] = ""
   Wend
   Err.Clear
    a.writeline (Now())
   DoCmd.Close acForm, "GUG"
End Sub



Public Function GETLASTFR(co As String, dt As Long) As Double
    Dim rst As New ADODB.Recordset, FNN As Long
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.N_KOL FROM dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.INVO_LST.TAG = 9) AND (dbo.INVO_LST.CODE = N'" & co & "') AND (dbo.HEAD_LST.DATE_N <= " & dt & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If IsNull(rst.Fields("N_KOL")) Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     TOP 100 PERCENT FNUMB FROM dbo.HEAD_MANF WHERE     (CODE = N'" & co & "') ORDER BY FNUMB DESC ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                GETLASTFR = rst.Fields("FNUMB")
            Else
                GETLASTFR = 0
            End If
        Else
            FNN = rst.Fields("N_KOL")
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     TOP 100 PERCENT CODE,FNUMB FROM dbo.HEAD_MANF WHERE     (FNUMB = " & FNN & " and CODE = N'" & co & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                GETLASTFR = rst.Fields("FNUMB")
            Else
                Set rst = New ADODB.Recordset
                rst.Open "SELECT     TOP 100 PERCENT FNUMB FROM dbo.HEAD_MANF WHERE     (CODE = N'" & co & "') ORDER BY FNUMB DESC ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    GETLASTFR = rst.Fields("FNUMB")
                Else
                    GETLASTFR = 0
                End If
             End If
        End If
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT FNUMB FROM dbo.HEAD_MANF WHERE     (CODE = N'" & co & "') ORDER BY FNUMB DESC ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            GETLASTFR = rst.Fields("FNUMB")
        Else
            GETLASTFR = 0
        End If
    End If
End Function


Public Sub REFRESHMOG(CODE As String, ANBAR As Integer, frm As String)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset, MAND As Double
    rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & CODE & "') AND (dbo.STUF_FSK.ANBAR = " & ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        MAND = rst.Fields("MAND")
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & CODE & "' AND ANBAR = " & ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          rst.Fields("MOGODI") = MAND
          rst.Fields("MOGODI_A") = 0
          rst.update
        End If
        Forms(frm).[MOGU] = MAND
    End If

End Sub
'تخفيفات پيشرفته را برحسب نوع مشتري به فاكتور اضافه ميكند
Public Sub ADDTAKH(CUST_KIND As Long, NUMBER As Long, KIND As Long)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
If MID(Forms![BASEKNOW]![OPTIONSS], 58, 1) = "5" Then
    DoCmd.RunSQL ("DELETE FROM dbo.TAKHFIF_APLAY WHERE (NUMBER = " & NUMBER & ") AND (KIND = " & KIND & ")")
    Set rst = New ADODB.Recordset
    Set RST2 = New ADODB.Recordset
    RST2.Open "TAKHFIF_APLAY", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SELECT TID FROM dbo.CUSTKIND_TF WHERE  (ACTIV = 0) AND (CUST_COD = " & CUST_KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        RST2.AddNew
        RST2.Fields("TID") = rst.Fields("TID")
        RST2.Fields("NUMBER") = NUMBER
        RST2.Fields("KIND") = KIND
        RST2.update
        rst.MoveNext
    Wend
End If
End Sub
Public Sub APLAYTAKH(NUMBER As Long, KIND As Long, M_NAGHD As Double, MABL_VAR As Double, MABL_HAV As Double, TICMBAA As Boolean)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, JAMMAB As Double, per As Double, OOK As Boolean, JST As New ADODB.Recordset, JAMF As Double, JAMCH As Double
If MID(Forms![BASEKNOW]![OPTIONSS], 58, 1) = "5" Then
  If Not IsNull(NUMBER) Then
        JST.Open "SELECT  SUM(MABL_K - ROUND(N_KOL/100 * MABL_K, 0))  AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & NUMBER & " ) AND ((INVO_LST.TAG)=" & KIND & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
            JAMF = JST.Fields(0)
        Else
             JAMF = 0
        End If
        Set JST = New ADODB.Recordset
        JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (PAY_GETD.DATE >= PAY_GETD.DATE_S AND PAY_GETD.TAG = " & KIND & " AND PAY_GETD.NUMBER= " & NUMBER & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
           JAMCH = JST.Fields(0)
        Else
           JAMCH = 0
        End If
        per = 0
        rst.Open "SELECT     dbo.TAKHFIF_APLAY.*, dbo.TAKHFIF_DEF.TKIND FROM dbo.TAKHFIF_APLAY INNER JOIN  dbo.TAKHFIF_DEF ON dbo.TAKHFIF_APLAY.TID = dbo.TAKHFIF_DEF.TID WHERE (NUMBER = " & NUMBER & ") AND (KIND = " & KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            Select Case rst.Fields("TKIND")
                Case 1: JAMMAB = M_NAGHD + MABL_VAR + MABL_HAV + JAMCH
                Case 2: JAMMAB = JAMF
            End Select
            RST2.Open "SELECT     TID, AZMAB, TAMAB, TPERS, TMAB FROM dbo.TAKHFIF_DEF_DTL WHERE     TID = " & rst.Fields("TID"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            OOK = False
            While Not RST2.EOF And Not OOK
                  If JAMMAB >= RST2.Fields("AZMAB") And JAMMAB <= RST2.Fields("TAMAB") Then
                    If JAMF > 0 Then
                      per = per + RST2.Fields("TPERS") * JAMMAB / JAMF
                      OOK = True
                    Else
                      per = per + RST2.Fields("TPERS")
                      OOK = True
                    End If
                  End If
                  RST2.MoveNext
            Wend
            Set RST2 = New ADODB.Recordset
            rst.MoveNext
        Wend
        If JAMF > 0 Then
            Set RST2 = New ADODB.Recordset
            RST2.Open "SELECT     TKHN,code,megh,meghk,mabl,mabl_k,n_kol,n_moin,imbaa FROM dbo.invo_lst WHERE     tag = " & KIND & " and NUMBER =  " & NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not RST2.EOF
              RST2.Fields("TKHN") = per
              RST2.Fields("N_MOIN") = Round(RST2.Fields("N_KOL") * RST2.Fields("MABL_K") / 100) + Round((RST2.Fields("MABL_K") - Round(RST2.Fields("N_KOL") * RST2.Fields("MABL_K") / 100)) * per / 100)
              If TICMBAA Then
                  Set rst = New ADODB.Recordset
                rst.Open "select CMBAA ,code from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("CMBAA") Then
                        If RST2.Fields("IMBAA") <> Fix((RST2.Fields("MABL_K") - RST2.Fields("N_MOIN")) * GetArzesh(rst.Fields("code")) / 100) Then
                             RST2.Fields("IMBAA") = Fix((RST2.Fields("MABL_K") - RST2.Fields("N_MOIN")) * GetArzesh(rst.Fields("code")) / 100)
                        End If
                    Else
                       If RST2.Fields("IMBAA") <> 0 Then
                            Forms![BASEKNOW]![Text44] = False
                            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                            If Forms![BASEKNOW]![Text44] Then
                                RST2.Fields("IMBAA") = 0
                            End If
                       End If
                    End If
                End If
              Else
               RST2.Fields("IMBAA") = 0
              End If
              RST2.update
              RST2.MoveNext
          Wend
        Else
        '    If KIND = 14 Then
         '        Set rst2 = New ADODB.Recordset
          '       rst2.Open "SELECT NUMBER, TAG, TAKHFIF FROM dbo.HEAD_LST WHERE (TAG = 14)  and NUMBER =  " & NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           '      If rst2.RecordCount > 0 Then
            '        rst2.Fields("TAKHFIF") = Round(PER * JAMMAB / 100)
             '       rst2.update
              '   End If
               '  rst2.Close
            ' End If
        End If
   End If
End If
End Sub

Public Function GETSTANDARDPRICE_MAVAD(CODE As String, dt As Long) As Double

                    On Error Resume Next
   Dim rst As New ADODB.Recordset, fnum As Double
   fnum = GETLASTFR(CODE, dt)
   If fnum = 0 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "') GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ORDER BY dbo.HEAD_MANF.FNUMB", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_MAVAD = rst.Fields("SumOfMABLK")
        Else
             GETSTANDARDPRICE_MAVAD = 0
        End If
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "' and dbo.HEAD_MANF.FNUMB = " & fnum & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_MAVAD = rst.Fields("SumOfMABLK")
        Else
             GETSTANDARDPRICE_MAVAD = 0
        End If
   End If
End Function

Public Function GETSTANDARDPRICE_DAST(CODE As String, dt As Long) As Double

                    On Error Resume Next
   Dim rst As New ADODB.Recordset, fnum As Double
   fnum = GETLASTFR(CODE, dt)
   If fnum = 0 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "') GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ORDER BY dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_DAST = rst.Fields("SumOfIMBIBE_MANF")
        Else
             GETSTANDARDPRICE_DAST = 0
        End If
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "' and dbo.HEAD_MANF.FNUMB = " & fnum & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_DAST = rst.Fields("SumOfIMBIBE_MANF")
        Else
             GETSTANDARDPRICE_DAST = 0
        End If
   End If
End Function

Public Function GETSTANDARDPRICE_SAR(CODE As String, dt As Long) As Double

                    On Error Resume Next
   Dim rst As New ADODB.Recordset, fnum As Double
   fnum = GETLASTFR(CODE, dt)
   If fnum = 0 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "') GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ORDER BY dbo.HEAD_MANF.FNUMB", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_SAR = rst.Fields("SumOfIMBIBE_SAR")
        Else
             GETSTANDARDPRICE_SAR = 0
        End If
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "' and dbo.HEAD_MANF.FNUMB = " & fnum & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_SAR = rst.Fields("SumOfIMBIBE_SAR")
        Else
             GETSTANDARDPRICE_SAR = 0
        End If
   End If
End Function


Public Sub SANADVORUDSAKHT(NUMBER As Long, NUMBER2 As Long)
                 On Error Resume Next
   Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, HEDRST As New ADODB.Recordset
   Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, a, fs
   Dim CON, i As Long
   DoCmd.OpenForm "GENSANADFROOSH", , , , , acHidden
   DoCmd.restore
   DoCmd.OpenForm "GUG"
   Forms![GUG]![Text0] = "n"
   Forms![GUG]![Text0].Visible = True
   Forms![GUG]![Box7].Visible = True
   Forms![GUG].Form.Refresh
   Forms![GUG].Form.Repaint
   SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   HEDRST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.DEPATMAN, HEAD_LST.SHIFT, HEAD_LST.CUST_KIND, HEAD_LST.USER_NAME FROM HEAD_LST WHERE ((HEAD_LST.NUMBER >= " & NUMBER & " AND HEAD_LST.NUMBER <=" & NUMBER2 & "  and HEAD_LST.tag = 9 ) )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   HEDRST.MoveLast
   HEDRST.MoveFirst
   Forms![GUG]![SNUM] = HEDRST.RecordCount
   CON = HEDRST.RecordCount / 50
   If CON < 1 Then
      CON = 1
   End If
   i = 1
     Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OPENTextFile("c:\LOGSANADVORUDSAKHT.txt", 8, 1)
    If Err <> 0 Then
        Set a = fs.CREATETextFile("c:\LOGSANADVORUDSAKHT.txt", True)
    End If
    a.writeline ("شروع باز سازي از سند شماره : " & NUMBER & " تا سند شماره :" & NUMBER2)
    a.writeline (Now())
   While Not HEDRST.EOF()
      Forms![GUG]![num] = i
      Set rst = New ADODB.Recordset
      If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
          rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If IsNull(rst.Fields(0)) Then
            max_ns = 1
          Else
            max_ns = rst.Fields(0) + 1
          End If
          Err.Clear
          SHART = "NO_S = 9 AND N_S =" & HEDRST.Fields("N_S")
          SHRST.Filter = SHART
          Set SHRST = SHRST
          If SHRST.RecordCount = 0 Or Err <> 0 Then
            Err.Clear
            SHRST.AddNew
            SHRST.Fields("N_S") = max_ns
            SHRST.Fields("DATE_S") = HEDRST.Fields("DATE_N")
            SHRST.Fields("SHARH_S") = Left(" برگه ورود كالا به انبار شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & "مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##"), 255)
            SHRST.Fields("GHATEI") = 0
            SHRST.Fields("NO_S") = 9
            SHRST.Fields("OKF") = True
            SHRST.Fields("USER_NAME") = HEDRST.Fields("USER_NAME")
            SHRST.Fields("CRT") = Now()
            SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
            SHRST.update
          Else
               max_ns = SHRST.Fields("N_S")
          End If
          If IsNull(HEDRST.Fields("N_S")) Or HEDRST.Fields("N_S") <> max_ns Then
             HEDRST.Fields("N_S") = max_ns
          End If
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Repaint
          DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 9))")
          CHRST.Open "SELECT     dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH) AS SumOfMEGH, SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR,SUM(dbo.INVO_LST.MABL) AS SumOfMABL, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF,dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME FROM dbo.STUF_DEF INNER JOIN  dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE AND dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE GROUP BY dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME " & _
                      " HAVING (dbo.INVO_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") AND (dbo.INVO_LST.TAG = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Repaint
       If MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) <> "5" Then
          While Not CHRST.EOF()
            JAMCH = 0
            JST.Open "SELECT DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT] AS MEGHM, INVO_LST.ANBAR  FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.ANBAR HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & HEDRST.Fields("NUMBER") & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "') AND ((INVO_LST.ANBAR)=" & CHRST.Fields("ANBAR") & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not JST.EOF()
               If JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK") <> 0 Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code"))) Then
                      Err.Clear
                      Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code")), GETKALANAME(JST.Fields("code")))
                      If Err <> 0 Then
                          a.writeline ("خطا در برگه شماره :" & HEDRST.Fields("NUMBER") & " نوع :" & HEDRST.Fields("tag") & "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد.")
                      End If
                    End If
                    SDRST.AddNew ' كنترل كالاي در جريان
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                    SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                    SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                    SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                    SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHM") * CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(CHRST.Fields("NAME")), 255)
                    SDRST.Fields("BES") = Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                    JAMCH = JAMCH + Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                    SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 9
                    SDRST.update
                    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                    Forms![GUG].Form.Repaint
               End If
                JST.MoveNext
            Wend
          '   Set JST = New ADODB.Recordset
            Set JST = New ADODB.Recordset
            JST.Open "SELECT STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.MEGHk, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR, INVO_LST.CODE FROM STUF_DEF INNER JOIN (INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) ON STUF_DEF.CODE = INVO_LST.CODE WHERE (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & HEDRST.Fields("NUMBER") & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If JST.RecordCount > 0 Then
                If JST.Fields("IMBIBE_SAR") * JST.Fields("MEGHK") > 0 Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999998) Then
                      Err.Clear
                      Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999998, "سربار")
                      If Err <> 0 Then
                          a.writeline ("خطا در برگه شماره :" & HEDRST.Fields("NUMBER") & " نوع :" & HEDRST.Fields("tag") & "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد.")
                      End If
                    End If
                    SDRST.AddNew ' جذب  سربار
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                    SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                    SDRST.Fields("HES_T") = 99999998
                    SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(JST.Fields("NAME")), 255)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999998"
                    SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                    JAMCH = JAMCH + Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                    SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 9
                    SDRST.update
                    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                    Forms![GUG].Form.Repaint
                 End If
                 If JST.Fields("IMBIBE_MANF") * JST.Fields("MEGHK") > 0 Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999999) Then
                      Err.Clear
                      Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999999, "دستمزد")
                      If Err <> 0 Then
                          a.writeline ("خطا در برگه شماره :" & HEDRST.Fields("NUMBER") & " نوع :" & HEDRST.Fields("tag") & "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد.")
                      End If
                    End If
                    SDRST.AddNew ' جذب  دستمزد
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                    SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                    SDRST.Fields("HES_T") = 99999999
                    SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(JST.Fields("NAME")), 255)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999999"
                    SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                    JAMCH = JAMCH + Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                    SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                    SDRST.Fields("TAG") = 9
                    SDRST.update
                    Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                    Forms![GUG].Form.Repaint
                  End If
            End If
            If JAMCH <> 0 Then
                SDRST.AddNew ' انبار محصول
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                SDRST.Fields("HES_M") = CHRST.Fields("ANBAR")
                SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & CHRST.Fields("ANBAR") & "-" & CLng(CHRST.Fields("CODE"))
                SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(CHRST.Fields("NAME")), 255)
                SDRST.Fields("BED") = Round(JAMCH)
                SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                SDRST.Fields("TAG") = 9
                SDRST.update
                Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
                Forms![GUG].Form.Repaint
            End If
            CHRST.MoveNext
           '  Set JST = New ADODB.Recordset
            Set JST = New ADODB.Recordset
          Wend
        Else
          Set CHRST = New ADODB.Recordset
          CHRST.Open "SELECT     dbo.INVO_LST.NUMBER, dbo.INVO_LST.N_KOL, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH) AS SumOfMEGH, SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR,SUM(dbo.INVO_LST.MABL) AS SumOfMABL, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF,dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME FROM dbo.STUF_DEF INNER JOIN  dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE AND dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE GROUP BY dbo.INVO_LST.NUMBER, dbo.INVO_LST.N_KOL, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME " & _
                           " HAVING (dbo.INVO_LST.NUMBER = " & HEDRST.Fields("NUMBER") & ") AND (dbo.INVO_LST.TAG = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          Forms![GUG]![Text2] = Forms![GUG]![Text2] + "n"
          Forms![GUG].Form.Repaint
               While Not CHRST.EOF()
                 JAMCH = 0
                 Set JST = New ADODB.Recordset
                 JST.Open "SELECT dbo.HEAD_MANF.FNUMB,DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT] AS MEGHM, INVO_LST.anbar  FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY  dbo.HEAD_MANF.FNUMB ,DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & HEDRST.Fields("NUMBER") & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "') AND ((INVO_LST.anbar)=" & CHRST.Fields("anbar") & " AND  (dbo.HEAD_MANF.FNUMB = " & IIf(IsNull(CHRST.Fields("N_KOL")), 0, CHRST.Fields("N_KOL")) & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 While Not JST.EOF()
                    If JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK") <> 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code"))) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code")), CHRST.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                          End If
                         End If
                         SDRST.AddNew ' كنترل كالاي در جريان
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                         SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & JST.Fields("MEGHM") * CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("BES") = Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                         SDRST.Fields("TAG") = 9
                         SDRST.update
                    End If
                    JST.MoveNext
                 Wend
                 Set JST = New ADODB.Recordset
                 JST.Open "SELECT     IMBIBE_MANF, IMBIBE_SAR, CODE FROM dbo.HEAD_MANF WHERE     (FNUMB = " & IIf(IsNull(CHRST.Fields("N_KOL")), 0, CHRST.Fields("N_KOL")) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If JST.RecordCount > 0 Then
                     If JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK") > 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999998) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999998, "سربار")
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999998"
                          End If
                        End If
                         SDRST.AddNew ' جذب  سربار
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                         SDRST.Fields("HES_T") = 99999998
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999998"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                         SDRST.Fields("TAG") = 9
                         SDRST.update
                      End If
                      If JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK") > 0 Then
                         SDRST.AddNew ' جذب  دستمزد
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                         SDRST.Fields("HES_T") = 99999999
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                         SDRST.Fields("TAG") = 9
                         SDRST.update
                       End If
                 End If
                 If JAMCH <> 0 Then
                     SDRST.AddNew ' انبار محصول
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                     SDRST.Fields("HES_M") = CLng(CHRST.Fields("ANBAR"))
                     SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & CLng(CHRST.Fields("ANBAR")) & "-" & CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("SHARH") = Left("برگه ورود شماره " & HEDRST.Fields("NUMBER") & "-" & HEDRST.Fields("FNUMCO") & " مورخ " & Format(HEDRST.Fields("DATE_N"), "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                     SDRST.Fields("BED") = JAMCH
                     SDRST.Fields("NUMBER") = HEDRST.Fields("NUMBER")
                     SDRST.Fields("TAG") = 9
                     SDRST.update
                 End If
                 CHRST.MoveNext
                 Set JST = New ADODB.Recordset
               Wend
        End If
    Else
          DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & HEDRST.Fields("NUMBER") & ") AND ((DEED_DTL.TAG)= 9))")
     End If
     CHRST.Close
     HEDRST.MoveNext
     If (i Mod CON) = 0 Then
        Forms![GUG]![Text0] = Forms![GUG]![Text0] + "n"
        Forms![GUG].Form.Repaint
     End If
     i = i + 1
     Forms![GUG]![Text2] = ""
     rst.Close
   Wend
    a.writeline (Now())
   DoCmd.Close acForm, "GUG"

End Sub
Public Function GETSTANDARDPRICE_KOL(CODE As String, dt As Long) As Double

                    On Error Resume Next
   Dim rst As New ADODB.Recordset, fnum As Double
   fnum = GETLASTFR(CODE, dt)
   If fnum = 0 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "') GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ORDER BY dbo.HEAD_MANF.FNUMB", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_KOL = rst.Fields("SumOfMABLK") + rst.Fields("SumOfIMBIBE_MANF") + rst.Fields("SumOfIMBIBE_SAR")
        Else
             GETSTANDARDPRICE_KOL = 0
        End If
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     TOP 100 PERCENT SUM(dbo.DTL_MANF.MABLK) AS SumOfMABLK, dbo.HEAD_MANF.IMBIBE_MANF AS SumOfIMBIBE_MANF,  dbo.HEAD_MANF.IMBIBE_SAR AS SumOfIMBIBE_SAR, dbo.HEAD_MANF.FNUMB FROM dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.HEAD_MANF.CODE = N'" & CODE & "' and dbo.HEAD_MANF.FNUMB = " & fnum & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.FNUMB ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
             GETSTANDARDPRICE_KOL = rst.Fields("SumOfMABLK") + rst.Fields("SumOfIMBIBE_MANF") + rst.Fields("SumOfIMBIBE_SAR")
        Else
             GETSTANDARDPRICE_KOL = 0
        End If
   End If
End Function

Function LASTPRICENZ(CODE As String, dt As Long) As Long
  Dim RRST As New ADODB.Recordset
    RRST.Open "SELECT     TOP 1  PERCENT dbo.INVO_LST.MABL FROM dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.INVO_LST.MABL <> 0 ) AND (dbo.HEAD_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & CODE & "') AND (dbo.HEAD_LST.DATE_N <= " & dt & ") ORDER BY dbo.HEAD_LST.DATE_N DESC, dbo.INVO_LST.MABL DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RRST.RecordCount > 0 Then
      If IsNull(RRST.Fields("MABL")) Then
          LASTPRICENZ = 0
      Else
          LASTPRICENZ = RRST.Fields("MABL")
      End If
    Else
      Set RRST = New ADODB.Recordset
      RRST.Open "SELECT     TOP 1 PERCENT CODE, FI_A FROM dbo.STUF_FSK WHERE code = '" & CODE & "' ORDER BY FI_A DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RRST.RecordCount > 0 Then
           If IsNull(RRST.Fields("FI_A")) Or RRST.Fields("FI_A") = 0 Then
                Set RRST = New ADODB.Recordset
                RRST.Open "SELECT     SMABL FROM dbo.DTL_MANF WHERE      (CODE = N'" & CODE & "') ORDER BY SMABL DESC ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RRST.RecordCount > 0 Then
                    LASTPRICENZ = RRST.Fields("SMABL")
                Else
                    LASTPRICENZ = 0
               End If
           Else
              LASTPRICENZ = RRST.Fields("FI_A")
           End If
        Else
           LASTPRICENZ = 0
        End If
    End If
  RRST.Close
End Function

Public Function Frisok(HES As String) As Boolean
                 On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, a, fs
   If MID(Forms![BASEKNOW]![OPTIONSS], 64, 1) = "5" Then
     CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "frsnd" & Forms![BASEKNOW]![USERCOD] & "')   DROP view " & "frsnd" & Forms![BASEKNOW]![USERCOD])
     CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "frinv" & Forms![BASEKNOW]![USERCOD] & "')   DROP view " & "frinv" & Forms![BASEKNOW]![USERCOD])
     DoCmd.RunSQL ("create  view frsnd" & Forms![BASEKNOW]![USERCOD] & " as SELECT     SUM(BED - BES) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 13) AND (HES = N'" & HES & "') GROUP BY NUMBER")
     DoCmd.RunSQL ("create  view frinv" & Forms![BASEKNOW]![USERCOD] & " as SELECT     dbo.HEAD_LST.NUMBER,dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) AS Expr1 FROM         dbo.HEAD_LST INNER JOIN dbo.JAMFACTPRS ON dbo.HEAD_LST.NUMBER = dbo.JAMFACTPRS.NUMBER AND dbo.HEAD_LST.TAG = dbo.JAMFACTPRS.TAG + 11 LEFT OUTER JOIN dbo.jamchkfact ON dbo.JAMFACTPRS.NUMBER = dbo.jamchkfact.NUMBER WHERE     (dbo.HEAD_LST.TAG = 13) AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') GROUP BY dbo.HEAD_LST.NUMBER,dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) HAVING      (dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) <> 0)")

     rst.Open "SELECT     dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER, dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 FROM   dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & " RIGHT OUTER JOIN  dbo.frinv" & Forms![BASEKNOW]![USERCOD] & " ON dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".NUMBER = dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER WHERE     (dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 > 0) AND (dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 IS NULL) OR (round(dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 - dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2,0) <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.OPENTextFile("c:\errorfr.txt", 8, 1)
        If Err <> 0 Then
            Set a = fs.CREATETextFile("c:\errorfr.txt", True)
        End If
        a.writeline (rst.Fields("NUMBER"))
        a.writeline (Now())
        a.Close
        For i = 1 To rst.RecordCount
           Call GENSANADFROOSH(rst.Fields("NUMBER"), rst.Fields("NUMBER"))
           rst.MoveNext
        Next i
        Frisok = False
     Else
        Frisok = True
     End If
     rst.Requery
     rst.Requery
     rst.Requery
     rst.Requery
     If rst.RecordCount > 0 Then
        Frisok = False
     Else
        Frisok = True
     End If
  Else
     Frisok = True
  End If

End Function


Public Function Khisok(HES As String) As Boolean
            On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, a, fs
   If MID(Forms![BASEKNOW]![OPTIONSS], 64, 1) = "5" Then
     CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "frsnd" & Forms![BASEKNOW]![USERCOD] & "')   DROP view " & "frsnd" & Forms![BASEKNOW]![USERCOD])
     CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "frinv" & Forms![BASEKNOW]![USERCOD] & "')   DROP view " & "frinv" & Forms![BASEKNOW]![USERCOD])
     DoCmd.RunSQL ("create  view frsnd" & Forms![BASEKNOW]![USERCOD] & " as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'" & HES & "') GROUP BY NUMBER")
   '  DoCmd.RunSQL ("create  view frinv" & Forms![BASEKNOW]![USERCOD] & " as SELECT     dbo.INVO_LST.NUMBER, SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS Expr1, dbo.HEAD_LST.MBAA FROM         dbo.INVO_LST INNER JOIN    dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG - 11 WHERE     (dbo.HEAD_LST.TAG = 12) AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') GROUP BY dbo.INVO_LST.NUMBER, dbo.HEAD_LST.MBAA HAVING      (SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) <> 0)")
     DoCmd.RunSQL ("create  view frinv" & Forms![BASEKNOW]![USERCOD] & " as SELECT     dbo.HEAD_LST.NUMBER,dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) AS Expr1 FROM         dbo.HEAD_LST INNER JOIN dbo.JAMFACTPRS ON dbo.HEAD_LST.NUMBER = dbo.JAMFACTPRS.NUMBER AND dbo.HEAD_LST.TAG = dbo.JAMFACTPRS.TAG + 11 LEFT OUTER JOIN dbo.jamchkfact ON dbo.JAMFACTPRS.NUMBER = dbo.jamchkfact.NUMBER WHERE     (dbo.HEAD_LST.TAG = 12) AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') GROUP BY dbo.HEAD_LST.NUMBER,dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) HAVING      (dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ - ISNULL(dbo.jamchkfact.mabch,0) <> 0)")
'     RST.Open "SELECT     dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER, dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 FROM   dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & " RIGHT OUTER JOIN  dbo.frinv" & Forms![BASEKNOW]![USERCOD] & " ON dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".NUMBER = dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER WHERE     (dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 + dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".MBAA > 0) AND (dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 IS NULL) OR (round(dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 + dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".MBAA - dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2,0) <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Open "SELECT     dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER, dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 FROM   dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & " RIGHT OUTER JOIN  dbo.frinv" & Forms![BASEKNOW]![USERCOD] & " ON dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".NUMBER = dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".NUMBER WHERE     (dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 > 0) AND (dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2 IS NULL) OR (round(dbo.frinv" & Forms![BASEKNOW]![USERCOD] & ".Expr1 - dbo.frsnd" & Forms![BASEKNOW]![USERCOD] & ".Expr2,0) <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.OPENTextFile("c:\errorkh.txt", 8, 1)
        If Err <> 0 Then
            Set a = fs.CREATETextFile("c:\errorkh.txt", True)
        End If
        a.writeline (Now())
        a.Close
        For i = 1 To rst.RecordCount
           a.writeline (rst.Fields("NUMBER"))
           Call GENSANADKHAREED(rst.Fields("NUMBER"), rst.Fields("NUMBER"))
           rst.MoveNext
        Next i
        Khisok = False
     Else
        Khisok = True
     End If
     rst.Requery
     rst.Requery
     rst.Requery
     rst.Requery
     If rst.RecordCount > 0 Then
        Khisok = False
     Else
        Khisok = True
     End If
  Else
     Khisok = True
  End If

End Function




Public Function Signed(tg As Integer, num As Long) As Boolean
             On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, a, fs
     rst.Open "SELECT     sgn2,sgn3 from head_lst  WHERE     (tag = " & tg & ") AND ( NUMBER =" & num & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        If rst.Fields("SGN2") Or rst.Fields("SGN3") Then
            Signed = True
        Else
            Signed = False
        End If
     Else
       Signed = False
     End If

End Function

Public Sub UpdateGHeymatFF(numb As Integer, tgg As Integer, PEPID As Long, PEID As Long, MODAT_PPID As Integer, TICMBAA As Integer, CUST_KIND As Integer)
               On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, TFF As Double, stf As Double, MLBAA As Double, IMBA As Double
    If MODAT_PPID <> 0 Then
        Call HAVEprice(numb, IIf(tgg = 13, 2, tgg), PEPID)
        RST2.Open "SELECT        dbo.PRICE_ELAMIE_DTL.PRICE1, dbo.PRICE_ELAMIE_DTL.PGID, dbo.STUF_DEF.CODE, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA,dbo.PRICE_ELAMIE_DTL.PEPID FROM dbo.PRICE_ELAMIE_DTL INNER JOIN dbo.STUF_DEF ON dbo.PRICE_ELAMIE_DTL.PGID = dbo.STUF_DEF.PGID INNER JOIN dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE WHERE        (dbo.INVO_LST.TAG = " & IIf(tgg = 13, 2, tgg) & ") AND dbo.INVO_LST.jay = 0  AND (dbo.INVO_LST.NUMBER =  " & numb & ") AND (dbo.PRICE_ELAMIE_DTL.PEPID =  " & PEPID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
            RST2.Fields("mabl") = RST2.Fields("price1")
            RST2.Fields("mabl_k") = Round(RST2.Fields("price1") * RST2.Fields("meghk"))
            RST2.update
            RST2.MoveNext
        Wend
        stf = 0
        MLBAA = 0
        IMBA = 0
        Set rst = New ADODB.Recordset
        Set RST2 = New ADODB.Recordset
        rst3.Open "SELECT        TOP (100) PERCENT PEID,TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & PEID & ") And (PPID = " & MODAT_PPID & ") And (CUSTCODE  = " & CUST_KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst3.RecordCount > 0 Then
          RST2.Open "SELECT        TOP (100) PERCENT dbo.INVO_LST.CODE,  dbo.INVO_LST.TKHN, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA , dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.STUF_DEF.CMBAA FROM            dbo.INVO_LST INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & numb & ") AND dbo.INVO_LST.jay = 0  And (dbo.INVO_LST.TAG = " & IIf(tgg = 13, 2, tgg) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          While Not RST2.EOF
                RST2.Fields("n_kol") = rst3.Fields("tf1")
                RST2.Fields("TKHN") = rst3.Fields("tf2")
                TFF = Round(RST2.Fields("MABL_K") * rst3.Fields("tf1") / 100) + Round((RST2.Fields("MABL_K") - Round(RST2.Fields("MABL_K") * rst3.Fields("tf1") / 100)) * rst3.Fields("tf2") / 100)
                stf = stf + TFF
                RST2.Fields("n_moin") = TFF
                If TICMBAA Then
                   If RST2.Fields("CMBAA") Then
                      IMBA = Fix((RST2.Fields("MABL_K") - TFF) * GetArzesh(RST2.Fields("code")) / 100)
                   Else
                      IMBA = 0
                    End If
                Else
                   IMBA = 0
                End If
                RST2.Fields("IMBAA") = IMBA
                MLBAA = MLBAA + IMBA
                RST2.update
                RST2.MoveNext
          Wend
        End If
        DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
        If tgg = 13 Then
             DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = 2")
        End If

    End If
End Sub


Public Sub UpdateGHeymat(numb As Integer, tgg As Integer, DTT As Long, MODAT_PPID As Integer, CUST_KIND As Integer, DEPATMAN As Long, TICMBAA As Integer)
               On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, TFF As Double, stf As Double, MLBAA As Double, IMBA As Double
    If MODAT_PPID <> 0 Then
        rst.Open "SELECT  TOP (100) PERCENT PEPID FROM dbo.PRICE_ELAMIE WHERE (PEPDATE <= " & DTT & ") And (PEPDEPART = " & DEPATMAN & ")ORDER BY PEPID DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Call HAVEprice(numb, IIf(tgg = 13, 2, tgg), rst.Fields("pepid"))
            RST2.Open "SELECT        dbo.PRICE_ELAMIE_DTL.PRICE1, dbo.PRICE_ELAMIE_DTL.PGID, dbo.STUF_DEF.CODE, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA,dbo.PRICE_ELAMIE_DTL.PEPID FROM dbo.PRICE_ELAMIE_DTL INNER JOIN dbo.STUF_DEF ON dbo.PRICE_ELAMIE_DTL.PGID = dbo.STUF_DEF.PGID INNER JOIN dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE WHERE        (dbo.INVO_LST.TAG = " & IIf(tgg = 13, 2, tgg) & ") AND (dbo.INVO_LST.NUMBER =  " & numb & ") AND dbo.INVO_LST.jay = 0  AND (dbo.PRICE_ELAMIE_DTL.PEPID =  " & rst.Fields("pepid") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            DoCmd.RunSQL ("UPDATE head_lst  SET PEPID  = " & rst.Fields("pepid") & " WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
            If tgg = 13 Then
                DoCmd.RunSQL ("UPDATE head_lst  SET PEPID  = " & rst.Fields("pepid") & " WHERE  NUMBER = " & numb & "  AND TAG = 2")
            End If
            While Not RST2.EOF
                RST2.Fields("mabl") = RST2.Fields("price1")
                RST2.Fields("mabl_k") = Round(RST2.Fields("price1") * RST2.Fields("meghk"))
                RST2.update
                RST2.MoveNext
            Wend
        Else
            DoCmd.RunSQL ("UPDATE INVO_LST SET IMBAA  = 0 , N_KOL = 0 ,N_MOIN = 0,TKHN = 0 , MABL_K =0 , MABL = 0   WHERE  NUMBER = " & numb & "  AND TAG = " & IIf(tgg = 13, 2, tgg))
            DoCmd.OpenForm "mesag", , , , , acDialog, " توجه براي اين واحد و اين نوع مشتري و اين نحوه پرداخت اعلاميه قيمت و تخفيف مشخص نشده است."
            Exit Sub
        End If
        stf = 0
        MLBAA = 0
        IMBA = 0
        Set rst = New ADODB.Recordset
        Set RST2 = New ADODB.Recordset
        rst.Open "SELECT        TOP (100) PERCENT PEID FROM dbo.PRICE_ELAMIETF WHERE (PEDATE <= " & DTT & ") And (PEPDEPART = " & DEPATMAN & ") ORDER BY PEID DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst3.Open "SELECT        TOP (100) PERCENT PEID,TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & rst.Fields("peid") & ") And (CUSTCODE = " & CUST_KIND & ") And (PPID = " & MODAT_PPID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            DoCmd.RunSQL ("UPDATE head_lst  SET PEID  = " & rst.Fields("PEID") & " WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
            If tgg = 13 Then
                DoCmd.RunSQL ("UPDATE head_lst  SET PEID  = " & rst.Fields("peid") & " WHERE  NUMBER = " & numb & "  AND TAG = 2")
            End If
            If rst3.RecordCount > 0 Then
              RST2.Open "SELECT        TOP (100) PERCENT dbo.INVO_LST.CODE,  dbo.INVO_LST.TKHN, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA , dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.STUF_DEF.CMBAA FROM            dbo.INVO_LST INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & numb & ") AND dbo.INVO_LST.jay = 0  And (dbo.INVO_LST.TAG = " & IIf(tgg = 13, 2, tgg) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              While Not RST2.EOF
                    RST2.Fields("n_kol") = rst3.Fields("tf1")
                    RST2.Fields("TKHN") = rst3.Fields("tf2")
                    TFF = Round(RST2.Fields("MABL_K") * rst3.Fields("tf1") / 100) + Round((RST2.Fields("MABL_K") - Round(RST2.Fields("MABL_K") * rst3.Fields("tf1") / 100)) * rst3.Fields("tf2") / 100)
                    stf = stf + TFF
                    RST2.Fields("n_moin") = TFF
                    If TICMBAA Then
                       If RST2.Fields("CMBAA") Then
                          IMBA = Fix((RST2.Fields("MABL_K") - TFF) * GetArzesh(RST2.Fields("code")) / 100)
                       Else
                          IMBA = 0
                        End If
                    Else
                       IMBA = 0
                    End If
                    RST2.Fields("IMBAA") = IMBA
                    MLBAA = MLBAA + IMBA
                    RST2.update
                    RST2.MoveNext
              Wend
            Else
                DoCmd.RunSQL ("UPDATE INVO_LST  SET N_KOL = 0 ,N_MOIN = 0,TKHN = 0  WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
                DoCmd.OpenForm "mesag", , , , , acDialog, " توجه براي اين واحد و اين نوع مشتري و اين نحوه پرداخت اعلاميه  تخفيف مشخص نشده است."
            End If
        End If
        DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
        If tgg = 13 Then
             DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = 2")
        End If

    End If
End Sub

Public Function GETGHeymatKala(numb As Integer, tgg As Integer, DTT As Long, MODAT_PPID As Integer, CUST_KIND As Integer, DEPATMAN As Integer, TICMBAA As Integer, CODECC As String, PEID As Long, PEPID As Long)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, TFF As Double, stf As Double, MLBAA As Double, IMBA As Double
    If MODAT_PPID <> 0 Then
        RST2.Open "SELECT        dbo.PRICE_ELAMIE_DTL.PRICE1, dbo.PRICE_ELAMIE_DTL.PGID, dbo.STUF_DEF.CODE, dbo.PRICE_ELAMIE_DTL.PEPID FROM         dbo.PRICE_ELAMIE_DTL INNER JOIN   dbo.STUF_DEF ON dbo.PRICE_ELAMIE_DTL.PGID = dbo.STUF_DEF.PGID WHERE     (dbo.STUF_DEF.CODE = '" & CODECC & "') AND (dbo.PRICE_ELAMIE_DTL.PEPID = " & PEPID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            GETGHeymatKala = RST2.Fields("price1")
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, " براي اين کالا قيمت يا گروه قيمتي تعريف نشده "
        End If
    End If
End Function
Public Function GETTaghfifKala1_old(numb As Integer, tgg As Integer, DTT As Long, MODAT_PPID As Integer, CUST_KIND As Integer, DEPATMAN As Integer, TICMBAA As Integer, CODECC As String, PEID As Long, PEPID As Long)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset
    If MODAT_PPID <> 0 Then
        rst3.Open "SELECT        TOP (100) PERCENT TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & PEID & ") And (CUSTCODE = " & CUST_KIND & ") And (PPID = " & MODAT_PPID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst3.RecordCount > 0 Then
            GETTaghfifKala1_old = rst3.Fields("TF1")
        Else
            DoCmd.RunSQL ("UPDATE INVO_LST  SET N_KOL = 0 ,N_MOIN = 0,TKHN = 0  WHERE  NUMBER = " & numb & "  AND CODE = '" & CODECC & "'  AND TAG = " & IIf(tgg = 13, 2, tgg))
            DoCmd.OpenForm "mesag", , , , , acDialog, " توجه براي اين واحد و اين نوع مشتري و اين نحوه پرداخت و اين کالا اعلاميه  تخفيف مشخص نشده است."
        End If
    End If
End Function


Public Function GETTaghfifKala2_old(numb As Integer, tgg As Integer, DTT As Long, MODAT_PPID As Integer, CUST_KIND As Integer, DEPATMAN As Integer, TICMBAA As Integer, CODECC As String, PEID As Long, PEPID As Long)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset
    If MODAT_PPID <> 0 Then
        rst3.Open "SELECT        TOP (100) PERCENT TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & PEID & ") And (CUSTCODE = " & CUST_KIND & ") And (PPID = " & MODAT_PPID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst3.RecordCount > 0 Then
            GETTaghfifKala2_old = rst3.Fields("TF2")
        Else
            DoCmd.RunSQL ("UPDATE INVO_LST  SET N_KOL = 0 ,N_MOIN = 0,TKHN = 0  WHERE  NUMBER = " & numb & "  AND CODE = '" & CODECC & "'  AND TAG = " & IIf(tgg = 13, 2, tgg))
            DoCmd.OpenForm "mesag", , , , , acDialog, " توجه براي اين واحد و اين نوع مشتري و اين نحوه پرداخت و اين کالا اعلاميه  تخفيف مشخص نشده است."
        End If
    End If
End Function
Public Function GETTaghfifKala1(ByVal numb As Integer, ByVal tgg As Integer, ByVal CUST_KIND As Integer, ByVal MODAT_PPID As Integer, ByVal PEID As Long, ByVal CODECC As String) As Single
     On Error Resume Next ' بهتر است مديريت خطاي دقيق‌تري داشته باشيد
    Dim rstRule As ADODB.Recordset
    Dim rstException As ADODB.Recordset
    Dim strSQL As String
    Dim sngTF1 As Single
    Dim lngPETID As Long

    sngTF1 = 0 ' مقدار پيش‌فرض

    If MODAT_PPID <> 0 And PEID <> 0 Then
        ' 1. ابتدا قانون تخفيف عمومي و PETID آن را پيدا کن
        Set rstRule = New ADODB.Recordset
        strSQL = "SELECT PETID, TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & PEID & ") AND (CUSTCODE = " & CUST_KIND & ") AND (PPID = " & MODAT_PPID & ")"
        rstRule.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        If Not rstRule.EOF Then
            lngPETID = rstRule.Fields("PETID").value
            sngTF1 = rstRule.Fields("TF1").value ' تخفيف عمومي را به عنوان پيش‌فرض در نظر بگير

            ' 2. حالا براي اين PETID و CODECC، در جدول استثنائات جستجو کن
            Set rstException = New ADODB.Recordset
            strSQL = "SELECT EXCEPTION_TF1 FROM dbo.PRICE_ELAMIETF_EXCEPTION WHERE (PETID = " & lngPETID & ") AND (CODE = '" & CODECC & "')"
            rstException.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

            If Not rstException.EOF Then
                ' اگر استثنا پيدا شد، از تخفيf استثنا استفاده کن
                sngTF1 = rstException.Fields("EXCEPTION_TF1").value
            End If

            If Not rstException Is Nothing Then
                If rstException.State = adStateOpen Then rstException.Close
                Set rstException = Nothing
            End If
        Else
            ' اگر قانون تخفيف عمومي يافت نشد، پيغام خطا و پاک کردن تخفيفات قبلي (اختياري)
             DoCmd.RunSQL ("UPDATE INVO_LST SET N_KOL = 0 ,N_MOIN = 0,TKHN = 0 WHERE NUMBER = " & numb & " AND CODE = '" & CODECC & "' AND TAG = " & IIf(tgg = 13, 2, tgg))
             DoCmd.OpenForm "mesag", , , , , acDialog, "توجه: براي اين واحد، نوع مشتري، نحوه پرداخت و کالا، اعلاميه تخفيف عمومي مشخص نشده است."
        End If

        If Not rstRule Is Nothing Then
            If rstRule.State = adStateOpen Then rstRule.Close
            Set rstRule = Nothing
        End If
    End If

    GETTaghfifKala1 = sngTF1

Exit_Function:
    If Not rstRule Is Nothing Then
        If rstRule.State = adStateOpen Then rstRule.Close
        Set rstRule = Nothing
    End If
    If Not rstException Is Nothing Then
        If rstException.State = adStateOpen Then rstException.Close
        Set rstException = Nothing
    End If
    Exit Function
'ErrorHandler:
    ' MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in GETTaghfifKala1"
    ' Resume Exit_Function
End Function
Public Function GETTaghfifKala2(ByVal numb As Integer, ByVal tgg As Integer, ByVal CUST_KIND As Integer, ByVal MODAT_PPID As Integer, ByVal PEID As Long, ByVal CODECC As String) As Single
     On Error Resume Next
    Dim rstRule As ADODB.Recordset
    Dim rstException As ADODB.Recordset
    Dim strSQL As String
    Dim sngTF2 As Single
    Dim lngPETID As Long

    sngTF2 = 0 ' مقدار پيش‌فرض

    If MODAT_PPID <> 0 And PEID <> 0 Then
        ' 1. ابتدا قانون تخفيف عمومي و PETID آن را پيدا کن
        Set rstRule = New ADODB.Recordset
        strSQL = "SELECT PETID, TF1, TF2 FROM dbo.PRICE_ELAMIETF_DTL WHERE (PEID = " & PEID & ") AND (CUSTCODE = " & CUST_KIND & ") AND (PPID = " & MODAT_PPID & ")"
        rstRule.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        If Not rstRule.EOF Then
            lngPETID = rstRule.Fields("PETID").value
            sngTF2 = rstRule.Fields("TF2").value ' تخفيف عمومي را به عنوان پيش‌فرض در نظر بگير

            ' 2. حالا براي اين PETID و CODECC، در جدول استثنائات جستجو کن
            Set rstException = New ADODB.Recordset
            strSQL = "SELECT EXCEPTION_TF2 FROM dbo.PRICE_ELAMIETF_EXCEPTION WHERE (PETID = " & lngPETID & ") AND (CODE = '" & CODECC & "')"
            rstException.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic

            If Not rstException.EOF Then
                ' اگر استثنا پيدا شد، از تخفيf استثنا استفاده کن
                sngTF2 = rstException.Fields("EXCEPTION_TF2").value
            End If

            If Not rstException Is Nothing Then
                If rstException.State = adStateOpen Then rstException.Close
                Set rstException = Nothing
            End If
        Else
            ' اگر قانون تخفيف عمومي يافت نشد (پيغام خطا مشابه تابع قبل)
             DoCmd.RunSQL ("UPDATE INVO_LST SET N_KOL = 0 ,N_MOIN = 0,TKHN = 0 WHERE NUMBER = " & numb & " AND CODE = '" & CODECC & "' AND TAG = " & IIf(tgg = 13, 2, tgg))
             DoCmd.OpenForm "mesag", , , , , acDialog, "توجه: براي اين واحد، نوع مشتري، نحوه پرداخت و کالا، اعلاميه تخفيف عمومي مشخص نشده است."
        End If

        If Not rstRule Is Nothing Then
            If rstRule.State = adStateOpen Then rstRule.Close
            Set rstRule = Nothing
        End If
    End If

    GETTaghfifKala2 = sngTF2

Exit_Function:
    If Not rstRule Is Nothing Then
        If rstRule.State = adStateOpen Then rstRule.Close
        Set rstRule = Nothing
    End If
    If Not rstException Is Nothing Then
        If rstException.State = adStateOpen Then rstException.Close
        Set rstException = Nothing
    End If
    Exit Function
'ErrorHandler:
    ' MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in GETTaghfifKala2"
    ' Resume Exit_Function
End Function
Public Function Getmodat(MODAT_PPID As Integer) As Integer
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT PPID, MODAT FROM PRICE_PAYNO where ppid = " & MODAT_PPID, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Getmodat = rst.Fields("modat")
    Else
        Getmodat = 0
    End If
End Function



Public Function GetDayPrice(COD As String, dt As Long) As Double
               On Error Resume Next
   Dim rst As New ADODB.Recordset, sql As String, PRGID As Long
   rst.Open "PRGHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.AddNew
   rst.Fields("pcode") = COD
   rst.Fields("NUM") = 1
   rst.Fields("grp") = 1010
   rst.Fields("PRG_DATE") = FARSIDATE(DATE)
   rst.update
   PRGID = rst.Fields("PRGID")
   Set rst = New ADODB.Recordset
   sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH, dbo.DTL_MANF.MEGHk,dbo.DTL_MANF.PERT, dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.FNUMB = " & GETLASTFR(COD, dt) & ")"
   DoCmd.RunSQL sql
   sql = "update dbo.PROGPAS1 Set PASED = -1  WHERE     (RADAH <> 2) AND (RADAH <> 3) AND (PRGID = " & PRGID & ")"
   DoCmd.RunSQL sql
   rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  rst.MoveLast
 '  rst.MoveFirst
   While rst.RecordCount > 0
        While Not rst.EOF()
           sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH * " & rst.Fields("meghk") & " as megh , dbo.DTL_MANF.MEGHk * " & rst.Fields("meghk") & " as meghk ,dbo.DTL_MANF.PERT * " & rst.Fields("meghk") & " AS PERT , (dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT) * " & rst.Fields("meghk") & " AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK * " & rst.Fields("meghk") & " AS MABLK ,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.FNUMB = " & GETLASTFR(rst.Fields("codB"), FARSIDATE(Now())) & ")"
           DoCmd.RunSQL sql
           rst.Fields("pased") = True
           rst.Fields("FR") = GETLASTFR(rst.Fields("codB"), dt)
           rst.update
           rst.MoveNext
        Wend
        Set rst = New ADODB.Recordset
        rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        rst.MoveLast
      '  rst.MoveFirst
    Wend
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     CODB,PRGID, MEGHK, MABL, MABLK, RADAH FROM dbo.PROGPAS1 WHERE     (PRGID = " & PRGID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
       rst.Fields("MABL") = LASTPRICENZ(rst.Fields("CODB"), dt)
       rst.Fields("MABLK") = rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), dt)
       rst.update
        rst.MoveNext
    Wend
    Dim MAVAD As Double
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     dbo.QPROGPAS2.PRGID, dbo.QPROGPAS2.CODB, dbo.QPROGPAS2.VAHED, dbo.QPROGPAS2.SumOfMEGH AS MEGH,   dbo.QPROGPAS2.SumOfMEGHK AS MEGHK, dbo.QPROGPAS2.SumOfPERT AS PERT, dbo.QPROGPAS2.SumOfKOLMAV AS KOLMAV, dbo.QPROGPAS2.PASED , dbo.QPROGPAS2.RADAH, dbo.STUF_DEF.NAME FROM  dbo.QPROGPAS2 INNER JOIN dbo.STUF_DEF ON dbo.QPROGPAS2.CODB = dbo.STUF_DEF.CODE WHERE     (dbo.QPROGPAS2.PRGID = " & PRGID & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        MAVAD = MAVAD + rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), dt)
        rst.MoveNext
    Wend
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.HEAD_MANF.IMBIBE_MANF * dbo.PROGPAS1.MEGHK) AS MANF, SUM(dbo.HEAD_MANF.IMBIBE_SAR * dbo.PROGPAS1.MEGHK)   AS SAR FROM         dbo.PROGPAS1 INNER JOIN    dbo.HEAD_MANF ON dbo.PROGPAS1.FR = dbo.HEAD_MANF.FNUMB WHERE     (dbo.PROGPAS1.RADAH <> 1) AND (dbo.PROGPAS1.PRGID = " & PRGID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MANF")) Then
        GetDayPrice = Round(GETSTANDARDPRICE_SAR(COD, dt) + rst.Fields("SAR") + GETSTANDARDPRICE_DAST(COD, dt) + rst.Fields("MANF") + MAVAD)
    Else
        GetDayPrice = Round(GETSTANDARDPRICE_SAR(COD, dt) + GETSTANDARDPRICE_DAST(COD, dt) + MAVAD)
    End If
 End Function

Public Function GETFROOSH(DT1, DT2 As Long) As Double
  Dim MAHLIST As String, XXD As Long, XXM As Long, XXS As Long, MH As Long, mah As Long, SAL As Long, SH As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Long
  rst.Open "SELECT     SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS FR FROM     dbo.DEED_DTL INNER JOIN   dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.DEED_HED.DATE_S BETWEEN " & DT1 & " AND " & DT2 & ") AND (dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![FROSH] & " OR dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![MFROSH] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 And Not IsNull(rst.Fields("FR")) Then
    GETFROOSH = rst.Fields("FR")
  Else
    GETFROOSH = 0
  End If
End Function

Public Function GETBSCHEK(SAL As Long, mah As Long) As Double
  Dim rst As New ADODB.Recordset, DT1, DT2 As Long
  DT1 = SAL * 10000 + mah * 100
  DT2 = SAL * 10000 + mah * 100 + 31
  rst.Open "SELECT     SUM(MABL) AS JCH FROM         dbo.PAY_GETD WHERE     (N_KOL IS NULL OR    N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL) AND (N_KOL3 IS NULL)  AND (DATE_S BETWEEN " & DT1 & " AND " & DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 And Not IsNull(rst.Fields("JCH")) Then
    GETBSCHEK = rst.Fields("JCH")
  Else
    GETBSCHEK = 0
  End If
End Function

'جمع چکهاي دريافتي در محدوده تاريخي با سررسيد در محدوده تاريخي اول
Public Function GETCHEKin(DT1, DT2 As Long, dt11, dt22 As Long) As Double
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT     SUM(MABL) AS JCH FROM         dbo.PAY_GETD WHERE     (DATE_S BETWEEN " & DT1 & " AND " & DT2 & ") and (DATE BETWEEN " & dt11 & " AND " & dt22 & ") and (N_KOL <> 911 or N_KOL IS  NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 And Not IsNull(rst.Fields("JCH")) Then
    GETCHEKin = rst.Fields("JCH")
  Else
    GETCHEKin = 0
  End If
End Function
Public Sub HAVEprice(numb As Integer, tgg As Integer, PEID As Long)
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
  RST2.Open "SELECT   CODE FROM INVO_LST WHERE     (dbo.INVO_LST.TAG =" & tgg & ") AND (dbo.INVO_LST.NUMBER = " & numb & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not RST2.EOF
        rst.Open "SELECT     dbo.STUF_DEF.CODE FROM         dbo.STUF_DEF LEFT OUTER JOIN    dbo.PRICE_ELAMIE_DTL ON dbo.STUF_DEF.PGID = dbo.PRICE_ELAMIE_DTL.PGID WHERE     (dbo.PRICE_ELAMIE_DTL.PEPID = " & PEID & ") AND (dbo.STUF_DEF.CODE = N'" & RST2.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields("CODE")) Then
            DoCmd.OpenForm "mesag", , , , , acDialog, " کالاي : '" & RST2.Fields("CODE") & "' - '" & GETKALANAME(RST2.Fields("CODE")) & Chr(10) & " داراي گروه بندي قيميتي نيست يا  گروه آن  در اعلاميه قيمت تعريف نشده. "
        End If
        Set rst = New ADODB.Recordset
        RST2.MoveNext
  Wend
End Sub


Public Sub TR(tbl As String, wxs As String, dt As Date, FLAGU As Integer)
       On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, FLS As String, trtx As String
    FLS = GETfldlist(tbl)
AGAINR:    Call TREXIXTCREATE(tbl)
    If FLAGU = 1 Then
       trtx = FARSIDATE(Now()) & " AS Expr2  , " & CDbl(dt) & " AS Expr1,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'"
       Err.Clear
       DoCmd.RunSQL "INSERT INTO dbo.TR_" & tbl & "(" & FLS & " , UP_DATE,UP_TIME, UP_USER_NAME,PC_NAME,IPADD) SELECT  " & FLS & " , " & trtx & " FROM dbo." & tbl & " WHERE " & wxs
       If Err <> 0 Then
            Err.Clear
            CurrentProject.Connection.Execute ("sp_rename  'TR_" & tbl & "', 'TR_" & tbl & FARSIDATE(Now()) & "'")
            If Err <> 0 Then
                Exit Sub
            End If
            GoTo AGAINR
       End If
    Else
       trtx = FARSIDATE(Now()) & " AS Expr2  , " & CDbl(dt) & " AS Expr1"
       CurrentProject.Connection.Execute ("SET IDENTITY_INSERT dbo.TR_" & tbl & " OFF  INSERT INTO dbo.TR_" & tbl & "(" & FLS & " ,UP_DATE ,UP_TIME ) SELECT  " & FLS & " , " & trtx & " FROM dbo." & tbl & " WHERE " & wxs)
       If Err <> 0 Then
            Err.Clear
            CurrentProject.Connection.Execute ("sp_rename  'TR_" & tbl & "', 'TR_" & tbl & FARSIDATE(Now()) & "'")
            If Err <> 0 Then
                Exit Sub
            End If
            GoTo AGAINR
       End If
    End If

End Sub

Public Sub TREXIXTCREATE(tbl As String)
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, FLS As String
   rst.Open "SELECT     sys.sysobjects.name AS clmn FROM         sys.sysobjects WHERE     (sys.sysobjects.name = N'TR_" & tbl & "') ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        FLS = "CREATE TABLE [dbo].[TR_" & tbl & "] ("
        RST2.Open "SELECT     TOP (100) PERCENT sys.syscolumns.name, sys.systypes.name AS TYP, sys.syscolumns.xtype, sys.syscolumns.length FROM         sys.syscolumns INNER JOIN  sys.systypes ON sys.syscolumns.xtype = sys.systypes.xtype INNER JOIN   sys.sysobjects ON sys.syscolumns.id = sys.sysobjects.id WHERE     (sys.sysobjects.name = N'" & tbl & "') AND (sys.systypes.name <> N'sysname') ORDER BY sys.syscolumns.colid", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
            If RST2.Fields("xtype") = 99 Or RST2.Fields("xtype") = 35 Or RST2.Fields("xtype") = 231 Or RST2.Fields("xtype") = 167 Or RST2.Fields("xtype") = 175 Or RST2.Fields("xtype") = 239 Then
                FLS = FLS & "[" & RST2.Fields("NAME") & "] [" & RST2.Fields("TYP") & "] (" & RST2.Fields("length") & ") NULL,"
            Else
                FLS = FLS & "[" & RST2.Fields("NAME") & "] [" & RST2.Fields("TYP") & "] NULL,"
            End If
            RST2.MoveNext
        Wend
        FLS = FLS & "[UP_DATE] [bigint] NOT NULL,[UP_TIME] [float] NOT NULL,[UP_USER_NAME] [nvarchar](40) NULL, [PC_NAME] [nvarchar](50) NULL, [IPADD] [nvarchar](50) NULL,[TRIDD] [int] IDENTITY(1,1) NOT NULL, PRIMARY KEY CLUSTERED([TRIDD] ASC) ON [PRIMARY]) ON [PRIMARY]"
        DoCmd.RunSQL FLS
    End If
End Sub



Public Function GETfldlist(tbl As String) As String
  Dim rst As New ADODB.Recordset, fldl As String
  rst.Open "SELECT     sys.syscolumns.name AS clmn FROM         sys.sysobjects INNER JOIN   sys.syscolumns ON sys.sysobjects.id = sys.syscolumns.id WHERE     (sys.sysobjects.name = N'" & tbl & "') and  (sys.sysobjects.uid = 1)  ORDER BY sys.syscolumns.colid", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not rst.EOF
    If fldl <> "" Then
        fldl = fldl & ",[" & rst.Fields("clmn") & "]"
    Else
        fldl = "[" & rst.Fields("clmn") & "]"
    End If
    rst.MoveNext
  Wend
  GETfldlist = fldl
End Function

Function GetMandMog(CODE As String, ANBAR As Long, dt As Long) As Double
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(" & dt & "," & ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(" & dt & "," & ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & CODE & "') AND (dbo.STUF_FSK.ANBAR = " & ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        GetMandMog = rst.Fields("MAND")
  Else
        GetMandMog = 0
  End If

End Function
Public Sub Updatearzesh(numb As Long, tgg As Integer, TICMBAA As Integer)
    Dim RST2 As New ADODB.Recordset, TFF As Double, stf As Double, MLBAA As Double, IMBA As Double
    RST2.Open "SELECT        TOP (100) PERCENT dbo.INVO_LST.CODE,  dbo.INVO_LST.TKHN, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA , dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.STUF_DEF.CMBAA FROM            dbo.INVO_LST INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & numb & ") And (dbo.INVO_LST.TAG = " & IIf(tgg = 13, 2, tgg) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RST2.EOF
          TFF = Round(RST2.Fields("MABL_K") * RST2.Fields("n_kol") / 100) + Round((RST2.Fields("MABL_K") - Round(RST2.Fields("MABL_K") * RST2.Fields("n_kol") / 100)) * RST2.Fields("TKHN") / 100)
          stf = stf + TFF
          RST2.Fields("n_moin") = TFF
          If TICMBAA Then
             If RST2.Fields("CMBAA") Then
                IMBA = Fix((RST2.Fields("MABL_K") - TFF) * GetArzesh(RST2.Fields("code")) / 100)
             Else
                IMBA = 0
              End If
          Else
             IMBA = 0
          End If
          RST2.Fields("IMBAA") = IMBA
          MLBAA = MLBAA + IMBA
          RST2.update
          RST2.MoveNext
    Wend
    DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = " & tgg)
    If tgg = 13 Then
         DoCmd.RunSQL ("UPDATE HEAD_LST SET MBAA  =  " & MLBAA & " , TAKHFIF = " & stf & "   WHERE  NUMBER = " & numb & "  AND TAG = 2")
    End If
End Sub

Public Function GETGRPKALAco(CC As String) As Integer
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     radah  FROM dbo.stuf_def WHERE     (CODE = '" & CC & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETGRPKALAco = rst.Fields("radah")
    End If
End Function


Public Function GetHesLevel(HES As String) As String
    Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
    Call GETTAF3(HES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    If Not IsNull(TAF4) Then
        GetHesLevel = "TDETA_HES4"
    Else
        If Not IsNull(TAF3) Then
            GetHesLevel = "TDETA_HES3"
        Else
            If Not IsNull(TAF2) Then
                GetHesLevel = "TDETA_HES2"
            Else
                If Not IsNull(taf) Then
                    GetHesLevel = "TDETA_HES"
                Else
                End If
            End If
        End If
    End If

End Function
