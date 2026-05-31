Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub CODE_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "F_MENU_KART"
   Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
   DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.CODE & "'"
   DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
End Sub
Private Sub Detail_MouseMove(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
   If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub
Private Sub FII1_DblClick(CANCEL As Integer)
                   On Error Resume Next
 '  Dim rst As New ADODB.Recordset, meg As Double, CODE As String, i As Long
 '  CurrentProject.Connection.Execute ("drop table kart_tmp")
 '  CurrentProject.Connection.Execute ("CREATE TABLE [kart_tmp] ( [ANBAR] [float] NULL , [avrage] [float] NULL ,[BARGAH] [nvarchar] (50) COLLATE Arabic_CI_AS NULL , [BEDNAME] [nvarchar] (151) COLLATE Arabic_CI_AS NULL ,[CODE] [nvarchar] (15) COLLATE Arabic_CI_AS NOT NULL , [DATE_N] [bigint] NULL , [FNUMCO] [float] NULL , [mabl_a] [float] NOT NULL ,[MABLK] [float] NULL , [MABM] [float] NULL , [MEG] [float] NULL , [MEGK] [float] NULL , [MEGKM] [float] NULL , [N_FANI] [nvarchar] (30) COLLATE Arabic_CI_AS NULL , [NAME] [nvarchar] (80) COLLATE Arabic_CI_AS NULL , [NAMES] [nvarchar] (32) COLLATE Arabic_CI_AS NOT NULL , [NUMBER] [float] NULL , [TAG] [float] NULL , [id] [int] IDENTITY (1, 1) NOT NULL , CONSTRAINT [PK_kart_tmp] PRIMARY KEY  CLUSTERED  ( [ID] )  ON [PRIMARY]) ON [PRIMARY] ")
 '  DoCmd.RunSQL "INSERT INTO dbo.kart_tmp  (ANBAR, avrage, BARGAH, BEDNAME, CODE, DATE_N, FNUMCO, mabl_a, MABLK, MABM, MEG, MEGK, MEGKM, N_FANI, NAME, NAMES, NUMBER, TAG) SELECT ANBAR, avrage, BARGAH, BEDNAME, RTRIM(LTRIM(CODE)) AS Expr1, DATE_N, FNUMCO, mabl_a, MABLK, MABM, MEG, MEGK, MEGKM, N_FANI, NAME, NAMES, NUMBER, TAG FROM   dbo.usp_KART_ANBAR_TOTAL_CLEAN_V3(" & Forms![ak_mogudi_anbar_list]![ANBAR] & ") usp_KART_ANBAR_TOTAL_CLEAN_V3 ORDER BY CODE, DATE_N, BARGAH, NUMBER"
 '  meg = 0
 '  rst.Open "SELECT  TOP 100 PERCENT ANBAR, avrage, BARGAH, BEDNAME, CODE, DATE_N, FNUMCO, mabl_a, MABLK, MABM, MEG, MEGK, MEGKM, N_FANI, NAME, NAMES, NUMBER,TAG FROM dbo.kart_tmp ORDER BY CODE, DATE_N, BARGAH, NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  DoCmd.OpenForm "BUN"
 '  While Not rst.EOF
 '       CODE = rst.Fields("code")
  '      meg = rst.Fields("meg")
 ''       rst.MoveNext
  '      i = i + 1
  '      If i Mod 100 = 0 Then
  '          Forms![BUN].Form.Refresh
  '      End If
  '      If Not rst.EOF Then
  '          While CODE = rst.Fields("code")
  '              meg = meg + rst.Fields("meg")
  '              rst.Fields("meg") = meg
  '              rst.update
  '              rst.MoveNext
  '              i = i + 1
  '              If rst.EOF Then
  '                   GoTo endf
  '              End If
  '          Wend
  '      End If
  ' Wend
'endf:

 DoCmd.OpenForm "KART_KALA_LIST", acFormDS
 DoCmd.Close acForm, "bun"
End Sub
Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub
Private Sub Form_Current()
                       On Error Resume Next
  If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_MouseMove(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
                       On Error Resume Next
  If Forms![BASEKNOW]![UGRP] = 4 And (Me.MABLK.ColumnHidden = False Or Me.FII1.ColumnHidden = False) Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
End Sub
Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
                     DoCmd.Maximize
If [Forms]![F_MENU_ANBAR]![MANBAR] = "'" Then
    Me.RecordSource = "EXEC [dbo].[sp_Mogudi_Tafkik_Optimized] " & [Forms]![F_MENU_ANBAR]![DT2] & ",'" & [Forms]![F_MENU_ANBAR]![MANBAR] & "'"
Else
    Me.RecordSource = "EXEC [dbo].[sp_Mogudi_Tafkik_Optimized] " & [Forms]![F_MENU_ANBAR]![DT2] & ",'" & [Forms]![F_MENU_ANBAR]![MANBAR] & "'"
End If
'Me.Requery
  Me.MAND.DecimalPlaces = Forms![BASEKNOW]![DIG]
   If Forms![BASEKNOW]![UGRP] = 4 Then
      Me.MABLK.ColumnHidden = True
      Me.FII1.ColumnHidden = True
   End If
   DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.Filter = Me.ServerFilter
   Me.FilterOn = True
 On Error Resume Next
  '  For Each itemm In Me.Form
  '      Me(itemm.NAME).ColumnWidth = -2
  '  Next
    Err.Clear
End Sub
Private Sub MABLK_DblClick(CANCEL As Integer)
           On Error Resume Next
   DoCmd.OpenForm "MOGHAYERAT_ANABR", acFormDS
End Sub
Private Sub N_FANI_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             FR = FR + rst.Fields("sizef")
        Next i
   End If
      '   ' DEBUG.PRINT sqlstr
     '   ' DEBUG.PRINT "ALTER FUNCTION dbo.AKMOGUDI_KOL_ANBAR_N_FANI (@Forms___F_MENU_ANBAR___DT2 BIGINT,@Forms___F_MENU_ANBAR___MANBAR nvarchar(10))RETURNS Table AS RETURN ( SELECT     CODE, MAND, MEGF, SMEGH, FII, MABLK, NAME, NAMES, KK, ANBAR, ANBARN, VCOD, GHEMAT, GRCOD, GRNAME, N_FANI, TOZIH, NESBAT, MANDF, " & sqlstr & " FROM         dbo.AKMOGUDI_KOL_ANBAR(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AKMOGUDI_KOL_ANBAR )"
        CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.AKMOGUDI_KOL_ANBAR_N_FANI (@Forms___F_MENU_ANBAR___DT2 BIGINT,@Forms___F_MENU_ANBAR___MANBAR nvarchar(10))RETURNS Table AS RETURN ( SELECT     CODE, MAND, MEGF, SMEGH, FII, MABLK, NAME, NAMES, KK, ANBAR, ANBARN, VCOD, GHEMAT, GRCOD, GRNAME, N_FANI, TOZIH, NESBAT, MANDF, " & sqlstr & " FROM         dbo.AKMOGUDI_KOL_ANBAR(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AKMOGUDI_KOL_ANBAR )")
   DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST_MP1", acFormDS
End Sub
