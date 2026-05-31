Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CODE_DblClick(CANCEL As Integer)

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
    '     DoCmd.RunSQL "DROP function  KHAREED_MP"


    CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.khareed_MP (@Forms___F_MENU_ANBAR___DT1 BIGINT,@Forms___F_MENU_ANBAR___DT2 BIGINT)RETURNS Table AS RETURN ( SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.INVO_LST.MEGHk, dbo.STUF_DEF.N_FANI, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, " & sqlstr & ", dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.NUMBER FROM         dbo.HEAD_LST INNER JOIN    dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE " & _
                                           "WHERE     (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N BETWEEN 1 AND 99999999))")
   DoCmd.OpenForm "khareed_MP1", acFormDS


End Sub

Private Sub Form_Close()
                   On Error Resume Next
  DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub TEDAD_Click()
   DoCmd.OpenForm "khareed_MP1", acFormDS
End Sub
