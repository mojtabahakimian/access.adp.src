Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command14_Click()
                      On Error Resume Next
   DoCmd.Close
End Sub

Private Sub Command2_Click()
                      On Error Resume Next
  Dim BEDEH, BESTA, max_ns As Double
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, SHRST As New ADODB.Recordset, rst As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset
   Dim CON, i As Long, JAZB As Double, JAZBK  As Double, JBES As Double
   rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  rst.Close
  Set rst = New ADODB.Recordset
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SHRST.AddNew
  SHRST.Fields("N_S") = max_ns
  SHRST.Fields("DATE_S") = FARSIDATE(DATE)
  SHRST.Fields("SHARH_S") = RIGHT(" سند عملكرد كالاها تا تاريخ  " & Format(FARSIDATE(DATE), "####/##/##"), 255)
  SHRST.Fields("GHATEI") = 0
  SHRST.Fields("NO_S") = 0
  SHRST.Fields("OKF") = True
  SHRST.Fields("USER_NAME") = UCurrentUser()
     SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
 SHRST.update
    'تهيه ليست جمع بدهكار كنترل كالاي درجريان
         CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MBED" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MBED" & Forms![BASEKNOW]![USERCOD])
         CurrentProject.Connection.Execute ("SELECT HES_T, SUM(BED - BES) AS BEDD INTO dbo.MBED" & Forms![BASEKNOW]![USERCOD] & " FROM dbo.DEED_DTL WHERE     (HES_K =" & Forms![BASEKNOW]![CONKAL] & ") AND (HES_M = 99999) GROUP BY HES_T")
         rst.Open "MBED" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    'تهيه ليست جمع بستانكار كنترل كالاي درجريان
         CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MBES" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MBES" & Forms![BASEKNOW]![USERCOD])
         CurrentProject.Connection.Execute ("SELECT HES_T, SUM(BES - BED) AS BESS INTO DBO.MBES" & Forms![BASEKNOW]![USERCOD] & " FROM dbo.DEED_DTL WHERE (HES_K = " & Forms![BASEKNOW]![CONKAL] & ") And (HES_M <> 99999)  GROUP BY HES_T ")
    'تهيه ليست جمع بستانكار كنترل كالاي درجريان به تفكيك كالاي ساخته شده
         CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MBESK" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MBESK" & Forms![BASEKNOW]![USERCOD])
         CurrentProject.Connection.Execute ("SELECT     HES_T, SUM(BES - BED) AS MBES, HES_M INTO DBO.MBESK" & Forms![BASEKNOW]![USERCOD] & " FROM dbo.DEED_DTL WHERE     (HES_K = " & Forms![BASEKNOW]![CONKAL] & ") GROUP BY HES_T, HES_M HAVING   (HES_M <> 99999)")

   Me.Text0 = "n"
   Me.Text0.Visible = True
   Me.Box7.Visible = True
   Me.Refresh
   Me.Text0 = Me.Text0 + "n"
   Me.Refresh
   Me.Text0.Requery
   Me.Repaint
    CON = rst.RecordCount / 100
   If CON < 1 Then
      CON = 1
   End If
   i = 1
   rst4.Open "DEED_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF()
     RST2.Open "SELECT HES_T,BESS FROM MBES" & Forms![BASEKNOW]![USERCOD] & " WHERE (HES_T = " & rst.Fields("HES_T") & " AND BESS <> 0 ) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
         JBES = RST2.Fields("BESS")
     Else
         JBES = 0
     End If
     rst3.Open "SELECT   HES_T, MBES, HES_M  FROM dbo.MBESK" & Forms![BASEKNOW]![USERCOD] & " WHERE     (HES_T = " & rst.Fields("HES_T") & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JAZBK = 0
     While Not rst3.EOF()
           JAZB = Round(rst3.Fields("MBES") / JBES * rst.Fields("BEDD"))
           rst4.AddNew
           rst4.Fields("N_S") = max_ns
           rst4.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
           rst4.Fields("HES_M") = rst3.Fields("HES_M")
           rst4.Fields("HES_T") = rst3.Fields("HES_T")
           rst4.Fields("HES") = [Forms]![BASEKNOW]![CONKAL] & "-" & rst3.Fields("HES_M") & "-" & rst3.Fields("HES_T")
           rst4.Fields("SHARH") = "انتقال از  بدهكار حساب كالاي درجريان سهم " & GETKALANAME(rst3.Fields("HES_M"))
           rst4.Fields("BED") = JAZB
           rst4.update
           JAZBK = JAZBK + JAZB
           If JAZB - rst3.Fields("MBES") <> 0 Then
                rst4.AddNew
                rst4.Fields("N_S") = max_ns
                rst4.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                rst4.Fields("HES_M") = rst3.Fields("HES_M")
                rst4.Fields("HES_T") = rst3.Fields("HES_T")
                rst4.Fields("HES") = [Forms]![BASEKNOW]![CONKAL] & "-" & rst3.Fields("HES_M") & "-" & rst3.Fields("HES_T")
                rst4.Fields("SHARH") = "انتقال به حساب عملكرد " & GETKALANAME(rst3.Fields("HES_M"))
                If JAZB - rst3.Fields("MBES") > 0 Then
                     rst4.Fields("BES") = JAZB - rst3.Fields("MBES")
                Else
                     rst4.Fields("BED") = rst3.Fields("MBES") - JAZB
                End If
                rst4.update
                If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst3.Fields("HES_M")), CLng(10000 & rst3.Fields("HES_T"))) Then
                  Err.Clear
                  Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(rst3.Fields("HES_M")), CLng(10000 & rst3.Fields("HES_T")), "مغايرت وزن " & GETKALANAME(rst3.Fields("HES_T")))
                End If
                rst4.AddNew
                rst4.Fields("N_S") = max_ns
                rst4.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                rst4.Fields("HES_M") = rst3.Fields("HES_M")
                rst4.Fields("HES_T") = 10000 & rst3.Fields("HES_T")
                rst4.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-" & rst3.Fields("HES_M") & "-10000" & rst3.Fields("HES_T")
                rst4.Fields("SHARH") = "انتقال به حساب عملكرد " & GETKALANAME(rst3.Fields("HES_M"))
                If JAZB - rst3.Fields("MBES") > 0 Then
                     rst4.Fields("BED") = JAZB - rst3.Fields("MBES")
                Else
                     rst4.Fields("BES") = rst3.Fields("MBES") - JAZB
                End If
                rst4.update
           End If
           rst3.MoveNext
     Wend
     'به خاطر رند شدن اعداد سرشكن شده ممكن است يك مقدار مانده در كنترل كالا بماند كه با سند دستي اصلاح شود
     rst4.AddNew
     rst4.Fields("N_S") = max_ns
     rst4.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
     rst4.Fields("HES_M") = "99999"
     rst4.Fields("HES_T") = rst.Fields("HES_T")
     rst4.Fields("HES") = [Forms]![BASEKNOW]![CONKAL] & "-99999-" & rst.Fields("HES_T")
     rst4.Fields("SHARH") = "انتقال   به بستانكار حساب كالاي درجريان  "
     rst4.Fields("BES") = JAZBK 'RST.Fields("BEDD")
     rst4.update
     If (i Mod CON) = 0 Then
        Me.Text0 = Me.Text0 + "n"
        Me.Refresh
        Me.Text0.Requery
        Me.Repaint
    End If
    i = i + 1
    rst.MoveNext
    RST2.Close
    Set RST2 = New ADODB.Recordset
    rst3.Close
    Set rst3 = New ADODB.Recordset
  Wend

  'مغايرت نرخ به كالاي در جريان ساخت
  rst.Close
  Set rst = New ADODB.Recordset
  rst.Open "SELECT HES_T, SUM(BED - BES) AS BEDD  FROM dbo.DEED_DTL WHERE     (HES_K =" & Forms![BASEKNOW]![AMALKARD] & ") AND (HES_M = 99999) GROUP BY HES_T", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not rst.EOF()
     RST2.Open "SELECT HES_T,BESS FROM MBES" & Forms![BASEKNOW]![USERCOD] & " WHERE (HES_T = " & rst.Fields("HES_T") & " AND BESS <> 0 ) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
         JBES = RST2.Fields("BESS")
     Else
         JBES = 0
     End If
     rst3.Open "SELECT   HES_T, MBES, HES_M  FROM dbo.MBESK" & Forms![BASEKNOW]![USERCOD] & " WHERE     (HES_T = " & rst.Fields("HES_T") & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JAZBK = 0
     While Not rst3.EOF()
           JAZB = Round(rst3.Fields("MBES") / JBES * rst.Fields("BEDD"))
           rst4.AddNew
           rst4.Fields("N_S") = max_ns
           rst4.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
           rst4.Fields("HES_M") = rst3.Fields("HES_M")
           rst4.Fields("HES_T") = rst3.Fields("HES_T")
           rst4.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-" & rst3.Fields("HES_M") & "-" & rst3.Fields("HES_T")
           rst4.Fields("SHARH") = "مغايرت نرخ انتقال از مغايرت كلي نرخ ، سهم " & GETKALANAME(rst3.Fields("HES_M"))
           If JAZB > 0 Then
                rst4.Fields("BED") = JAZB
           Else
                rst4.Fields("BES") = JAZB * -1
           End If
           rst4.update
           JAZBK = JAZBK + JAZB
           rst3.MoveNext
     Wend
     'به خاطر رند شدن اعداد سرشكن شده ممكن است يك مقدار مانده در كنترل كالا بماند كه با سند دستي اصلاح شود
     rst4.AddNew
     rst4.Fields("N_S") = max_ns
     rst4.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
     rst4.Fields("HES_M") = "99999"
     rst4.Fields("HES_T") = rst.Fields("HES_T")
     rst4.Fields("HES") = [Forms]![BASEKNOW]![AMALKARD] & "-99999-" & rst.Fields("HES_T")
     rst4.Fields("SHARH") = "انتقال به  حساب كالاي درجريان  "
     If JAZBK > 0 Then
        rst4.Fields("BES") = JAZBK 'RST.Fields("BEDD")
     Else
        rst4.Fields("BED") = JAZBK * -1 'RST.Fields("BEDD")
     End If
     rst4.update
     If (i Mod CON) = 0 Then
        Me.Text0 = Me.Text0 + "n"
        Me.Refresh
        Me.Text0.Requery
        Me.Repaint
    End If
    i = i + 1
    rst.MoveNext
    RST2.Close
    Set RST2 = New ADODB.Recordset
    rst3.Close
    Set rst3 = New ADODB.Recordset
  Wend
  If Err = 0 Then
    DoCmd.Hourglass False
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات صدور سند با موفقيت انجام شد" & " سند شماره : " & max_ns
    Me.Text0.Visible = False
    Me.Box7.Visible = False
    Me.Refresh
    DoCmd.Close acForm, Me.NAME
  End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Timer()
                     On Error Resume Next
   Me.Refresh
End Sub
