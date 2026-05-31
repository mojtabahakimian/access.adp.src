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

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   Me.MAND.DecimalPlaces = Forms![BASEKNOW]![DIG]

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
