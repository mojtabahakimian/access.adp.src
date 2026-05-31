Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Close()
   DoCmd.Close acForm, "F_MENU_AJNAS"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    If Forms![F_MENU_AJNAS]![cdd] = False Then
       Me.RecordSource = " SELECT STUF_DEF.NAME AS NAM, STUF_DEF.MABL_F, 1 AS Expr1, STUF_DEF.B_SEF, MOGUDI_KOL_ANBARHA.MANDAH FROM STUF_DEF INNER JOIN MOGUDI_KOL_ANBARHA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE"
    End If
    If Forms![F_MENU_AJNAS]![CHAP] = 2 Then

      Me.MABL_F.Visible = False
    Else
      Me.MABL_F.Visible = True
    End If
    If Forms![F_MENU_AJNAS]![CHAP] = 3 Then
      Me.MABL_F.ControlSource = "B_SEF"
    End If
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
