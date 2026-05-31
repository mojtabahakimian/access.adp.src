Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub mab_DblClick(CANCEL As Integer)

                 On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset

    If BLOCKEDMK(HES) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
       Exit Sub
    End If
      If IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
      End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.HES & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With

End Sub

Private Sub MABLK_DblClick(CANCEL As Integer)
                    On Error Resume Next
    If LETSGO("KARTR") Then
       DoCmd.OpenForm "F_MENU_KART"
       Forms![F_MENU_KART]![ANBAR] = Forms![ak_mogudi_anbar_list]![ANBAR]
       DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.CODE & "'"
       DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
    Else
       DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
    End If
End Sub
