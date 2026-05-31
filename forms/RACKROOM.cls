Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                       On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
  rst.Open "SELECT dbo.ROOM.ROOM_NUM, dbo.ROOM.RFLAG, dbo.ROOM.TAKHT, dbo.ROOM.VASAYEL, dbo.ROOM.COMMENT,  dbo.[ROOM SUB].MNAME +  ' ' + dbo.[ROOM SUB].MFAMILY AS NAM FROM         dbo.ROOM LEFT OUTER JOIN   dbo.[ROOM SUB] ON dbo.ROOM.ROOM_NUM = dbo.[ROOM SUB].ROOMN ORDER BY dbo.ROOM.ROOM_NUM", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  i = 1
   While Not rst.EOF()
     Me("R" & i).CAPTION = rst.Fields("ROOM_NUM")
     If IsNull(rst.Fields("NAM")) And (rst.Fields("RFLAG") = 3 Or rst.Fields("RFLAG") = 1) Then
       RST2.Open "SELECT     ROOMN, MAX(IDH) AS Expr1, SAZMANM FROM dbo.MEHMAN GROUP BY ROOMN, IDH, SAZMANM, KDATE HAVING (ROOMN = " & rst.Fields("ROOM_NUM") & ")AND (NOT (SAZMANM IS NULL))  And (Not (IDH Is Null)) And (KDATE Is Null)  ORDER BY MAX(IDH) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If RST2.RecordCount > 0 Then
             Me("L" & i).CAPTION = IIf(IsNull(RST2.Fields("SAZMANM")), " ", RST2.Fields("SAZMANM"))
       End If
       RST2.Close
     Else
         Me("L" & i).CAPTION = IIf(IsNull(rst.Fields("NAM")), " ", rst.Fields("NAM"))
     End If
     Select Case rst.Fields("RFLAG")
       Case 1: Me("R" & i).BackColor = 255 ' اشغال شده
       Case 2: Me("R" & i).BackColor = 6723891 'خالي
       Case 3: Me("R" & i).BackColor = 128  'واگذار شده شده
       Case 4: Me("R" & i).BackColor = 16737843 ' در حال سرويس
     End Select
     i = i + 1
     rst.MoveNext
   Wend
   'نمايش اتاقهاي داراي شماره
  For i = 1 To 50
     If Me("R" & i).CAPTION = "0" Then
        Me("R" & i).Visible = False
        Me("L" & i).Visible = False
     Else
        Me("R" & i).Visible = True
        Me("L" & i).Visible = True
     End If
  Next i
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  End Sub

Private Sub SETRN(VALR As String)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer, TF As Long
   rst.Open "ROOM", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "ROOM_NUM = " & Me(VALR).CAPTION
   Set rst = rst
   If rst.RecordCount = 0 Then
      Exit Sub
   Else

   End If

    Select Case Me.Frame321
      Case 1:
            If Me(VALR).BackColor <> 255 Then
              Me(VALR).BackColor = 255 ' اشغال شده
              rst.Fields("RFLAG") = Me.Frame321
              rst.update
            End If
      Case 2:
            If Me(VALR).BackColor <> 255 Then
              Me(VALR).BackColor = 6723891 'خالي
              rst.Fields("RFLAG") = Me.Frame321
              rst.update
            End If
      Case 3:
            If Me(VALR).BackColor <> 255 Then
                Me(VALR).BackColor = 128  'واگذار شده شده
                rst.Fields("RFLAG") = Me.Frame321
                rst.update
            End If
      Case 4:
            If Me(VALR).BackColor <> 255 Then
              Me(VALR).BackColor = 16737843 ' در حال سرويس
              rst.Fields("RFLAG") = Me.Frame321
              rst.update
            End If
      Case 5:  'فاكتور غذا
              RST2.Open "SELECT MEHMAN.TAFZ, MEHMAN.ROOMN FROM MEHMAN WHERE ((Not (MEHMAN.TAFZ) Is Null) AND ((MEHMAN.KDATE) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              RST2.Filter = "ROOMN =" & Me(VALR).CAPTION
              Set RST2 = RST2
              If RST2.RecordCount = 0 Then
              Else
                DoCmd.OpenForm "HEAD_LST_FROOSH"
                Forms![HEAD_LST_FROOSH]![CUST_NO] = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & RST2.Fields(0)
                If Hour(Now()) >= 6 And Hour(Now()) <= 11 Then
                    Forms![HEAD_LST_FROOSH]![MOLAH] = "صبحانه"
                Else
                    If Hour(Now()) > 11 And Hour(Now()) <= 16 Then
                        Forms![HEAD_LST_FROOSH]![MOLAH] = "نهار"
                    Else
                        If Hour(Now()) > 16 And Hour(Now()) <= 23 Then
                            Forms![HEAD_LST_FROOSH]![MOLAH] = "شام"
                         End If
                    End If
                 End If
              End If

      Case 6:  'فاكتور خدمات
              RST2.Open "SELECT MEHMAN.TAFZ, MEHMAN.ROOMN FROM MEHMAN WHERE ((Not (MEHMAN.TAFZ) Is Null) AND ((MEHMAN.KDATE) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              RST2.Filter = "ROOMN =" & Me(VALR).CAPTION
              Set RST2 = RST2
              If RST2.RecordCount = 0 Then
              Else
                DoCmd.OpenForm "HEAD_LST_KHADAMAT"
                Forms![HEAD_LST_KHADAMAT]![CUST_NO] = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & RST2.Fields(0)
              End If
      Case 7: 'صورتحساب
              RST2.Open "SELECT MEHMAN.TAFZ, MEHMAN.ROOMN, MEHMAN.VDATE FROM MEHMAN WHERE ((Not (MEHMAN.TAFZ) Is Null) AND ((MEHMAN.KDATE) Is Null)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              RST2.Filter = "ROOMN =" & Me(VALR).CAPTION
              Set RST2 = RST2
              If RST2.RecordCount = 0 Then
              Else
                DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", , , , , acHidden
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HKOL] = Forms![BASEKNOW]![BEDEHKAR]
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HMOIN] = 1
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HTAF] = RST2.Fields(0)
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HHKOL] = Forms![BASEKNOW]![BEDEHKAR]
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HHMOIN] = 1
                Forms![F_MENU_KOL_MOIN_TAFZIL]![HHTAF] = RST2.Fields(0)
                Forms![F_MENU_KOL_MOIN_TAFZIL]![DT1] = RST2.Fields("VDATE")
                DoCmd.OpenReport "R_SURAT_HESAB", acViewPreview
              End If
      Case 8: 'تحويل اتاق
              RST2.Open "SELECT MEHMAN.TAFZ, MEHMAN.ROOMN, MEHMAN.VDATE, MEHMAN.KDATE,MEHMAN.KTIME FROM MEHMAN WHERE ((Not (MEHMAN.TAFZ) Is Null) AND ((MEHMAN.KDATE) Is Null)) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              RST2.Filter = "ROOMN =" & Me(VALR).CAPTION
              Set RST2 = RST2
              If RST2.RecordCount = 0 Then
              Else
                 TF = RST2.Fields(0)
                 RST2.Fields("KDATE") = FARSIDATE(DATE)
                 RST2.Fields("KTIME") = Hour(Now()) * 100 + Minute(Now())
                 RST2.update
                 RST2.Open "SELECT TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & ") AND ((TDETA_HES.NUMBER)=1) AND ((TDETA_HES.TNUMBER)=" & TF & ")); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If RST2.RecordCount = 0 Then
                 Else
                   RST2.Fields("CODE_E") = True
                   RST2.update
                   '***************************
                   Me(VALR).BackColor = 16737843
                   rst.Fields("RFLAG") = 4
                   rst.update
                 End If
                 Call Form_Load
              End If
              'واگذار شده
              RST2.Close
              Set RST2 = New ADODB.Recordset
              RST2.Open "SELECT     TAFZ, ROOMN, VDATE, KDATE, KTIME FROM dbo.MEHMAN WHERE     (TAFZ IS NULL) AND (KDATE IS NULL) AND (ROOMN = " & Me(VALR).CAPTION & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If RST2.RecordCount = 0 Then
              Else
                 While Not RST2.EOF()
                    RST2.Fields("KDATE") = FARSIDATE(DATE)
                    RST2.Fields("KTIME") = Hour(Now()) * 100 + Minute(Now())
                    RST2.update
                    RST2.MoveNext
                 Wend
                 Me(VALR).BackColor = 16737843
                 rst.Fields("RFLAG") = 4
                 rst.update
                 Call Form_Load
              End If
    End Select
    rst.Close
End Sub

Private Sub r1_Click()
                     On Error Resume Next
  SETRN ("R1")
End Sub


Private Sub r2_Click()
                      On Error Resume Next
   SETRN ("R2")
End Sub
Private Sub r3_Click()
                      On Error Resume Next
   SETRN ("R3")
End Sub
Private Sub r4_Click()
                      On Error Resume Next
   SETRN ("R4")
End Sub
Private Sub r5_Click()
                     On Error Resume Next
  SETRN ("R5")
End Sub
Private Sub r6_Click()
                      On Error Resume Next
   SETRN ("R6")
End Sub
Private Sub r7_Click()
                     On Error Resume Next
  SETRN ("R7")
End Sub
Private Sub r8_Click()
                      On Error Resume Next
   SETRN ("R8")
End Sub
Private Sub r9_Click()
                      On Error Resume Next
   SETRN ("R9")
End Sub
Private Sub r10_Click()
                      On Error Resume Next
   SETRN ("R10")
End Sub
Private Sub r11_Click()
                      On Error Resume Next
   SETRN ("R11")
End Sub
Private Sub r12_Click()
                      On Error Resume Next
   SETRN ("R12")
End Sub
Private Sub r13_Click()
                      On Error Resume Next
   SETRN ("R13")
End Sub
Private Sub r14_Click()
                     On Error Resume Next
  SETRN ("R14")
End Sub
Private Sub r15_Click()
                      On Error Resume Next
   SETRN ("R15")
End Sub
Private Sub r16_Click()
                      On Error Resume Next
   SETRN ("R16")
End Sub
Private Sub r17_Click()
                      On Error Resume Next
   SETRN ("R17")
End Sub
Private Sub r18_Click()
                      On Error Resume Next
   SETRN ("R18")
End Sub
Private Sub r19_Click()
                     On Error Resume Next
  SETRN ("R19")
End Sub
Private Sub r20_Click()
                      On Error Resume Next
   SETRN ("R20")
End Sub
Private Sub r21_Click()
                     On Error Resume Next
  SETRN ("R21")
End Sub
Private Sub r22_Click()
                      On Error Resume Next
   SETRN ("R22")
End Sub
Private Sub r23_Click()
                     On Error Resume Next
  SETRN ("R23")
End Sub
Private Sub r24_Click()
                     On Error Resume Next
  SETRN ("R24")
End Sub
Private Sub r25_Click()
                     On Error Resume Next
  SETRN ("R25")
End Sub
Private Sub r26_Click()
                     On Error Resume Next
  SETRN ("R26")
End Sub
Private Sub r27_Click()
                      On Error Resume Next
   SETRN ("R27")
End Sub
Private Sub r28_Click()
                     On Error Resume Next
  SETRN ("R28")
End Sub
Private Sub r29_Click()
                      On Error Resume Next
   SETRN ("R29")
End Sub
Private Sub r30_Click()
                     On Error Resume Next
  SETRN ("R30")
End Sub
Private Sub r31_Click()
                     On Error Resume Next
  SETRN ("R31")
End Sub
Private Sub r32_Click()
                     On Error Resume Next
  SETRN ("R32")
End Sub
Private Sub r33_Click()
                      On Error Resume Next
   SETRN ("R33")
End Sub
Private Sub r34_Click()
                     On Error Resume Next
  SETRN ("R34")
End Sub
Private Sub r35_Click()
                     On Error Resume Next
  SETRN ("R35")
End Sub
Private Sub r36_Click()
                      On Error Resume Next
   SETRN ("R36")
End Sub
Private Sub r37_Click()
                     On Error Resume Next
  SETRN ("R37")
End Sub
Private Sub r38_Click()
                    On Error Resume Next
 SETRN ("R38")
End Sub
Private Sub r39_Click()
                     On Error Resume Next
  SETRN ("R39")
End Sub
Private Sub r40_Click()
                      On Error Resume Next
   SETRN ("R40")
End Sub
Private Sub r41_Click()
                     On Error Resume Next
  SETRN ("R41")
End Sub
Private Sub r42_Click()
                      On Error Resume Next
   SETRN ("R42")
End Sub
Private Sub r43_Click()
                     On Error Resume Next
  SETRN ("R43")
End Sub
Private Sub r44_Click()
                      On Error Resume Next
   SETRN ("R44")
End Sub
Private Sub r45_Click()
                     On Error Resume Next
  SETRN ("R45")
End Sub
Private Sub r46_Click()
                     On Error Resume Next
  SETRN ("R46")
End Sub
Private Sub r47_Click()
                      On Error Resume Next
   SETRN ("R47")
End Sub
Private Sub r48_Click()
                     On Error Resume Next
  SETRN ("R48")
End Sub
Private Sub r49_Click()
                     On Error Resume Next
  SETRN ("R49")
End Sub
Private Sub r50_Click()
                      On Error Resume Next
   SETRN ("R50")
End Sub

Private Sub Command331_Click()
                       On Error Resume Next
    DoCmd.OpenForm "MEHMAN"
End Sub
