Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
   Dim rst As New ADODB.Recordset, KOL As Long, MOIN As Long, taf As Long, NSCON As String, NSCONP As String
   Dim SHART As String
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
    While MID(SHES, i, 1) <> "-" And i < 100
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

Private Sub Command117_Click()
   DoCmd.OpenReport "SUDVAZIAN", acViewPreview
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   rst.Open "select * from surat_mali WHERE CODG = 0" & " and prg = " & Forms![surat_mali]![prg], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        If NSCON = "" Or IsNull(NSCON) Then
            NSCON = "(N_S = " & rst.Fields("HESAB") & ")"
        Else
            NSCON = NSCON & " OR (N_S = " & rst.Fields("HESAB") & ")"
        End If
        rst.MoveNext
   Wend
   Set rst = New ADODB.Recordset
   rst.Open "select * from surat_mali WHERE CODG = 25" & " and prg = " & Forms![surat_mali]![prg], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
       NSCONP = "(N_S <= " & rst.Fields("HESAB") & ")"
   Else
       NSCONP = "(N_S < 999999999)"
   End If
   SHART = GETSHART(1)
   'موجودي پايان دوره

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MOAVPA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(1)

   Set rst = New ADODB.Recordset
   'موجودي مواد اوليه
   SHART = "(" & SHART & ") AND  " & NSCON
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MOMAEB = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND"))
   'خريد طي دوره
   SHART = GETSHART(2)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.KHTEDO = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(2)
   'دستمزد مستقيم
   SHART = GETSHART(3)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.DASTO = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(3)
   'ساير هزينه هاي توليد
   SHART = GETSHART(4)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.SAHAZTO = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(4)
   'موجودي كالاي درجريان ساخت اول دوره:
   SHART = GETSHART(5)

   Set rst = New ADODB.Recordset
   SHART = "(" & SHART & ") AND  " & NSCON
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MOKADJAVDO = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(5)
   'موجودي كالاي در جريان ساخت پايان دوره:
   SHART = GETSHART(5)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MODJSAPA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(5)
   'موجودي كالا در ابتداي دوره:
   SHART = GETSHART(6)
   SHART = "(" & SHART & ") AND  " & NSCON

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MOKALAV = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(6)
   'خريد  كالا طي دوره:
   SHART = GETSHART(7)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.KHTYDO = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(7)
   'ساير
   SHART = GETSHART(8)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.SAYER = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(8)
   'موجودي كالا در پايان دوره:
   SHART = GETSHART(6)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.MOKALPA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(6)
   'فروش ناخالص:
   SHART = GETSHART(12)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.FRNAKH = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(12)
   'كسر مي شود برگشت از فروش و تخفيفات:
   SHART = GETSHART(13)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.BRFRT = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(13)
   'درآمد حاصل از خدمات:
   SHART = GETSHART(14)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.DRHAKHA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(14)
   'درآمد حاصل از معاملات بانكي و سرمايه گذاري:
   SHART = GETSHART(15)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.DRMOBA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(15)
   'درآمد حاصل از اجاره:
   SHART = GETSHART(16)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.DREJA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(16)
   'درآمدهاي غير مترقبه:
   SHART = GETSHART(17)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.DRGMOTAR = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(17)
   'ساير دريافتي ها و درآمدها:
   SHART = GETSHART(18)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.SAYERDA = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) * -1 + REPSUM(18)
   'حقوق و دستمزد:
   SHART = GETSHART(19)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.HOGH = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(19)
   'هزينه هاي اداري:
   SHART = GETSHART(20)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.HAZED = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(20)
   'هزينه هاي مالي:
   SHART = GETSHART(21)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.HAZMAL = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(21)
   'هزينه هاي توزيع و فروش:
   SHART = GETSHART(22)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.HAZTOZI = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(22)
   'هزينه هاي استهلاك(غير توليدي):
   SHART = GETSHART(23)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.HAZEST = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(23)
   'ساير هزينه ها:
   SHART = GETSHART(24)

   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS mand FROM dbo.DEED_DTL WHERE " & SHART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.SAYERHAZ = IIf(IsNull(rst.Fields("MAND")), 0, rst.Fields("MAND")) + REPSUM(24)
End Sub

Public Function GETSHART(RADIF As Long) As String
                  On Error Resume Next
   Dim SHARTLO As String, SRST As New ADODB.Recordset, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
   SHARTLO = ""
   SRST.Open "select * from surat_mali WHERE CODG = " & RADIF & " and prg = " & Forms![surat_mali]![prg], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not SRST.EOF
     If Not IsNull(SRST.Fields("HESAB")) And SRST.Fields("HESAB") <> "" Then
       If Not IsNull(SRST.Fields("HESAB")) Then
           Call GETTAF3(SRST.Fields("HESAB"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
       If CKOL <> 0 Then
         If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " OR ( HES_K = " & CKOL & ")"
         Else
            SHARTLO = "( HES_K = " & CKOL & ")"
         End If
       End If
       If CMOIN <> 0 Then
          If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " AND (HES_M = " & CMOIN & ")"
          Else
            SHARTLO = "(HES_M = " & CMOIN & ")"
          End If
       End If
       If CTAF <> 0 Then
          If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " AND (HES_T = " & CTAF & ")"
          Else
            SHARTLO = "(HES_T = " & CTAF & ")"
          End If
       End If
       If CTAF2 <> 0 Then
          If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " AND (HES_T2 = " & CTAF2 & ")"
          Else
            SHARTLO = "(HES_T2 = " & CTAF2 & ")"
          End If
       End If
       If CTAF3 <> 0 Then
          If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " AND (HES_T3 = " & CTAF3 & ")"
          Else
            SHARTLO = "(HES_T3 = " & CTAF3 & ")"
          End If
       End If
       If CTAF4 <> 0 Then
          If Not IsNull(SHARTLO) And SHARTLO <> "" Then
            SHARTLO = SHARTLO & " AND (HES_T4 = " & CTAF4 & ")"
          Else
            SHARTLO = "(HES_T4 = " & CTAF4 & ")"
          End If
       End If
     End If
     SRST.MoveNext
   Wend
   If SHARTLO = "" Then
        GETSHART = "HES_K = -11"
   Else
        GETSHART = "(" & SHARTLO & ") AND " & NSCONP
   End If
End Function
Public Function REPSUM(RADIF As Long) As Double
                  On Error Resume Next
   Dim SUMR As Double, SRST As New ADODB.Recordset
    Dim RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset
   SHARTLO = ""
   SRST.Open "select * from surat_mali WHERE CODG = " & RADIF & " and prg = " & Forms![surat_mali]![prg], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not SRST.EOF
        If Not (SRST.Fields("TOSIF") = "" Or IsNull(SRST.Fields("TOSIF"))) Then
            RST2.Open "SELECT  SQLST  FROM dbo.SQLState WHERE     (IDD = " & SRST.Fields("TOSIF") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If Not (RST2.Fields("SQLST") = "" Or IsNull(RST2.Fields("SQLST"))) Then
               CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "SURAT" & Forms![BASEKNOW]![USERCOD] & "')   DROP view " & "SURAT" & Forms![BASEKNOW]![USERCOD])
               CurrentProject.Connection.Execute ("CREATE  VIEW " & "SURAT" & Forms![BASEKNOW]![USERCOD] & " AS " & RST2.Fields("SQLST"))
               rst3.Open "SELECT * FROM SURAT" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst3.RecordCount > 0 Then
                  SUMR = SUMR + rst3.Fields(0)
               End If
            End If
            Set RST2 = New ADODB.Recordset
            Set rst3 = New ADODB.Recordset
        End If
        SRST.MoveNext
   Wend
   REPSUM = SUMR
End Function
