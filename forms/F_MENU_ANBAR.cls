Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CANBAR_AfterUpdate()
                      On Error Resume Next
   ANBAR = CANBAR
   Me.KALA.Requery
   Me.KALA.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF INNER JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (((STUF_FSK.ANBAR)=" & [Forms]![F_MENU_ANBAR]![ANBAR] & ")) GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME"
End Sub
Private Sub ANBAR_AfterUpdate()
                      On Error Resume Next
   CANBAR = ANBAR
   Me.KALA.Requery
   Me.KALA.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF INNER JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (((STUF_FSK.ANBAR)=" & [Forms]![F_MENU_ANBAR]![ANBAR] & ")) GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE ORDER BY STUF_DEF.NAME"
End Sub

Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If IsNull(Me.ANBAR) Then
     MANBAR = "%"
  Else
    MANBAR = Me.ANBAR
  End If
  If IsNull(Me.KALA) Then
     MKALA = "%"
  Else
    MKALA = Me.KALA
  End If
 Select Case Me.OpenArgs
    Case "R":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "R_AK_MOGUDI_ANBAR", acViewPreview
            Else
                DoCmd.OpenReport "R_AK_MOGUDI_ANBAR", acViewPreview, , "CODE = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "R2":
            If IsNull(Me.KALA) Then
                DoCmd.OpenReport "R_AK_MOGUDI_ANBAR2", acViewPreview
            Else
                DoCmd.OpenReport "R_AK_MOGUDI_ANBAR2", acViewPreview, , "CODE = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "F":
            If IsNull(Me.KALA) Then
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST", acFormDS
            Else
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST", acFormDS, , "mogudi_tafkik.CODE = '" & Me.KALA & "'", , , Me.KALA
            End If
              DoCmd.OpenForm Me.NAME, , , , , acHidden
  Case "F2":
            If IsNull(Me.KALA) Then
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST2", acFormDS
            Else
                DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST2", acFormDS, , "CODE = '" & Me.KALA & "'"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
  Case "MOGDT":
  'محاسبه تاريخ ايجاد مانده کالاهادر جدولmogdt+usercod
        DoCmd.OpenForm "BUN"
              DoCmd.OpenForm Me.NAME, , , , , acHidden
        Dim meg As Double, CODE As String
            CurrentProject.Connection.Execute ("IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[mogdt" & Forms![BASEKNOW]![USERCOD] & "]') AND type in (N'U')) BEGIN CREATE TABLE [dbo].[mogdt" & Forms![BASEKNOW]![USERCOD] & "](    [CODE] [nvarchar](15)  NOT NULL,    [ANBAR] [float] NOT NULL,    [MAND] [float] NOT NULL,    [FII] [float] NOT NULL,    [MABLK] [float] NOT NULL,    [NAMES] [NVARCHAR](50) NOT NULL,    [GRCOD] [float] NOT NULL,    [EGDATE] [bigint] NOT NULL, CONSTRAINT [PK_mogdt" & Forms![BASEKNOW]![USERCOD] & "] PRIMARY KEY CLUSTERED(    [CODE] ASC,    [ANBAR] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] End")
            CurrentProject.Connection.Execute ("IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_mogdt_EGDATE" & Forms![BASEKNOW]![USERCOD] & "]') AND type = 'D') BEGIN ALTER TABLE [dbo].[mogdt" & Forms![BASEKNOW]![USERCOD] & "] ADD  CONSTRAINT [DF_mogdt_EGDATE" & Forms![BASEKNOW]![USERCOD] & "]  DEFAULT ((0)) FOR [EGDATE] End")
            DoCmd.RunSQL "delete from dbo.mogdt" & Forms![BASEKNOW]![USERCOD]
            If IsNull(Me.KALA) Then
                 DoCmd.RunSQL "INSERT INTO dbo.mogdt" & Forms![BASEKNOW]![USERCOD] & " ([CODE],[ANBAR],[MAND], [FII] , [MABLK] , [NAMES] ,[GRCOD], [EGDATE]) SELECT  [CODE],[ANBAR],[MAND], [FII] , [MABLK] , [NAMES] ,[GRCOD], " & Forms![BASEKNOW]![YEA] & "0101 as EGDATE  FROM dbo.mogudi_tafkik(" & Me.DT2 & ",'" & MANBAR & "') WHERE     (dbo.mogudi_tafkik.mand <> 0 )"
            Else
                 DoCmd.RunSQL "INSERT INTO dbo.mogdt" & Forms![BASEKNOW]![USERCOD] & " ([CODE],[ANBAR],[MAND], [FII] , [MABLK] , [NAMES] ,[GRCOD], [EGDATE]) SELECT  [CODE],[ANBAR],[MAND], [FII] , [MABLK] , [NAMES] ,[GRCOD], " & Forms![BASEKNOW]![YEA] & "0101  as EGDATE  FROM dbo.mogudi_tafkik(" & Me.DT2 & ",'" & MANBAR & "') WHERE     (dbo.mogudi_tafkik.mand <> 0  and CODE = '" & Me.KALA & "')"
            End If
            DoCmd.OpenForm Me.NAME, , , , , acHidden
            rst.Open "mogdt" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst.EOF
                 If rst.Fields("CODE") = "2542" Then
                    DoCmd.Beep
                 End If
                 meg = rst.Fields("mand")
                 Set RST2 = New ADODB.Recordset
                 RST2.Open "SELECT RTRIM(LTRIM(CODE)) AS CODE,ANBAR,  BARGAH, DATE_N,  MEGK, MEGKM,  NUMBER, TAG FROM   dbo.KART_KALA(" & CLng(Forms![BASEKNOW]![YEA] & "0101") & ", " & rst.Fields("anbar") & ", " & Me.DT2 & ") KART_KALA where (tag = 1 or tag= 0 or tag = 6 or tag = 7 or tag = 4 or tag = 9 or tag = 17 or tag = 24 or tag = 22) and  code = '" & rst.Fields("code") & "' ORDER BY  DATE_N DESC, BARGAH DESC, NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 While Not RST2.EOF
                     meg = meg - RST2.Fields("MEGK")
                     If meg <= 0 Then
                         'If meg = 0 And Not RST2.BOF Then
                           '  RST2.MovePrevious
                         '    rst.Fields("EGDATE") = RST2.Fields("DATE_N")
                       '  Else
                          rst.Fields("EGDATE") = RST2.Fields("DATE_N")
                       '  End If
                         rst.update
                         GoTo badee
                     End If
                     RST2.MoveNext
                     Forms![BUN].Form.Refresh
                 Wend
                 If RST2.EOF Then
                    rst.Fields("EGDATE") = CLng(Forms![BASEKNOW]![YEA] & "0101")
                    rst.update
                 End If
badee:           rst.MoveNext
            Wend
            DoCmd.Close acForm, "bun"
            DoCmd.OpenForm "zmogudi_date", acFormDS
        End Select
  BringWindowToTop (Forms![BASEKNOW]![hhwin])

End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If


End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  Me.ANBAR.RowSource = "SELECT     CODE, NAMES FROM         TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE IN (SELECT  ANBCO  FROM dbo.OPANBACCESS  WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & ")))"
  Me.CANBAR.RowSource = "SELECT CODE FROM TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE IN (SELECT  ANBCO  FROM dbo.OPANBACCESS  WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & ")))"
End Sub

Private Sub KALA_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
  If Me.KALA.Text = "-" Or Me.KALA.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.KALA.Text = ""
     DoCmd.OpenForm "SERCHK2"
  Else
    Me.KALA = NewData
  End If
  Response = acDataErrContinue
End Sub
