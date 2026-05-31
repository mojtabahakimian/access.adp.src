Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Close()
                   On Error Resume Next
 If Not (IsLoaded("TARAZ_4") Or IsLoaded("TARAZ_4_MOIN") Or IsLoaded("TARAZ_TAF_DIRECT")) Then
       DoCmd.Close acForm, "FMENU_TARAZ_4"
    End If
    Me.ServerFilter = ""
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Me.Filter = Me.ServerFilter
   Me.FilterOn = True
If IsLoaded("TARAZ_4") Then
Else
   If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.Filter = "BED <> 0 OR BES <> 0"
        Me.FilterOn = True
   End If
End If
DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then

   DoCmd.OpenForm "DEED_SEARCH", acFormDS, , "HES_K = " & Me.حساب_كل & " AND HES_M = " & Me.معين & " AND TNUMBER = " & Me.تفصيلي & " AND N_S Between " & [Forms]![FMENU_TARAZ_4]![SNDNUM1] & " And " & [Forms]![FMENU_TARAZ_4]![SNDNUM2] & " AND DATE_S Between '" & Format([Forms]![FMENU_TARAZ_4]![DT1], "####/##/##") & "' And '" & Format([Forms]![FMENU_TARAZ_4]![DT2], "####/##/##") & "'"
   End If
End Sub

Private Sub تفصيلي_DblClick(CANCEL As Integer)
                     On Error Resume Next
     If BLOCKED(Me.حساب_كل, Me.معين, Me.تفصيلي) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
       DoCmd.OpenForm "DEED_SEARCH", acFormDS, , "HES_K = " & Me.حساب_كل & " AND HES_M = " & Me.معين & " AND TNUMBER = " & Me.تفصيلي & " AND N_S Between " & [Forms]![FMENU_TARAZ_4]![SNDNUM1] & " And " & [Forms]![FMENU_TARAZ_4]![SNDNUM2] & " AND DATE_S Between " & [Forms]![FMENU_TARAZ_4]![DT1] & " And " & [Forms]![FMENU_TARAZ_4]![DT2] & ""
    End If
End Sub

Private Sub حساب_كل_DblClick(CANCEL As Integer)
  Dim MAN As Double, rst As New ADODB.Recordset
     If BLOCKED(Me.حساب_كل, Me.معين, Me.تفصيلي) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
        If IsLoaded("R_DAFTAR_MOIN_LIST") Then
              DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
        End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.حساب_كل & "-" & Me.معين & "-" & Me.تفصيلي & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
    End If
End Sub

Private Sub معين_DblClick(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.OpenForm "DEED_SEARCH", acFormDS, , "HES_K = " & Me.حساب_كل & " AND HES_M = " & Me.معين & " AND N_S Between " & [Forms]![FMENU_TARAZ_4]![SNDNUM1] & " And " & [Forms]![FMENU_TARAZ_4]![SNDNUM2] & " AND DATE_S Between " & [Forms]![FMENU_TARAZ_4]![DT1] & " And " & [Forms]![FMENU_TARAZ_4]![DT2]
End Sub

Private Sub نام_حساب_تفصيلي_DblClick(CANCEL As Integer)
                 On Error Resume Next
    DoCmd.OpenReport "TARAZ4_TAFZIL2", acViewPreview
End Sub
