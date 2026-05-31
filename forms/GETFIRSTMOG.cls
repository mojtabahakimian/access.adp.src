Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command3_Click()
    Dim rst As New ADODB.Recordset
    Dim dt As Date
    Dim cn As New ADODB.Connection
                  On Error Resume Next
    dt = Now()
    If Not IsNull(Me.YEA) Then
        DoCmd.OpenForm "BUN"
 '       ' DEBUG.PRINT "INSERT INTO dbo.TR_STUF_DEF (CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART, IDD, CMBAA, VAZN, OKF, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT     CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART,IDD , CMBAA, VAZN, OKF, " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "' FROM  dbo.STUF_DEF "
        DoCmd.RunSQL ("INSERT INTO dbo.TR_STUF_DEF (CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART, IDD, CMBAA, VAZN, OKF, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT     CODE, NAME, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, KINDK, MABL_F, DEPART,IDD , CMBAA, VAZN, OKF, " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'System" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "' FROM  dbo.STUF_DEF ")
        Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("INSERT INTO dbo.TR_STUF_FSK (CODE, ANBAR, MOGODI_A, FI_A, MABL_A, MANDAH_A, VAZ, IDD, POSITION,B_SEF,N_SEF,MIN_M,MAX_M, UP_TIME, UP_DATE) SELECT     CODE, ANBAR, MOGODI_A, FI_A, MABL_A, MANDAH_A, VAZ, IDD, POSITION,B_SEF,N_SEF,MIN_M,MAX_M," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.STUF_FSK ")
        Forms![BUN].Form.Refresh
        CurrentProject.Connection.CommandTimeout = 0
        CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'STUFFSK')   DROP TABLE STUFFSK")
        rst.Open "select * from " & Me.YEA & ".DBO.tcod_anbar", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        cn.connectionString = CurrentProject.BaseConnectionString
        cn.Open
        cn.CommandTimeout = 0
        cn.Execute ("SELECT     CODE, MAND, FII, MABLK, ANBAR INTO dbo.STUFFSK FROM  " & Me.YEA & ".DBO.AKMOGUDI_KOL_ANBAR(99999999,'" & rst.Fields("CODE") & "') AKMOGUDI_KOL_ANBAR")
        rst.MoveNext
        While Not rst.EOF()
            cn.Execute ("INSERT INTO dbo.STUFFSK (CODE, MAND, FII, MABLK, ANBAR )  SELECT     CODE, MAND, FII, MABLK, ANBAR FROM         " & Me.YEA & ".DBO.AKMOGUDI_KOL_ANBAR(99999999,'" & rst.Fields("CODE") & "') AKMOGUDI_KOL_ANBAR")
            Forms![BUN].Form.Refresh
            rst.MoveNext
        Wend
        Set rst = New ADODB.Recordset
        rst.Open "select * from STUFFSK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            DoCmd.RunSQL "UPDATE    dbo.STUF_FSK SET MOGODI_A = " & rst.Fields("MAND") & ", FI_A = " & rst.Fields("FII") & ", MABL_A = " & rst.Fields("MABLK") & " WHERE     (CODE = N'" & rst.Fields("code") & "') AND (ANBAR = " & rst.Fields("ANBAR") & ")"
            rst.MoveNext
             Forms![BUN].Form.Refresh
        Wend
        DoCmd.Close acForm, "bun"
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " موجوديها به روز رساني شد "
        DoCmd.Close acForm, Me.NAME
     End If
End Sub
Private Sub Command4_Click()
                   On Error Resume Next
 DoCmd.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, SST As New ADODB.Recordset
   rst.Open "SELECT * FROM  master.dbo.sysdatabases ", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
   While Not rst.EOF
      Err.Clear
      SST.Open "SELECT YEA,NAME,SANAT FROM " & rst.Fields("NAME") & ".DBO.SAZMAN", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
      If Err = 0 Then
        YEA.RowSource = YEA.RowSource & rst.Fields("name") & ";"
      End If
      rst.MoveNext
      Set SST = New ADODB.Recordset
   Wend
  Call SETSECURITY(Me.NAME, "NEWYEAR", 3)
End Sub

Private Sub YEA_AfterUpdate()
                  On Error Resume Next
 If Me.YEA = "+" Or Me.YEA = "-" Then
   DoCmd.OpenForm "SELECTYEAR", acNormal, , , , acDialog, 2
 End If

End Sub
