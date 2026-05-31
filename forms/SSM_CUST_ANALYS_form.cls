Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CUST_NO_DblClick(CANCEL As Integer)
          On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset
  If Not IsNull(CUST_NO) Then
    If BLOCKEDMK(CUST_NO) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
       Exit Sub
    End If
    If IsLoaded("R_DAFTAR_MOIN_LIST") Then
          DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
    End If
    DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
    With Forms!R_DAFTAR_MOIN_LIST
        .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.CUST_NO & "'"
        .Requery                  ' بارگذاري مجدد داده
        .Visible = True           ' نمايش فرم
    End With
  End If
End Sub
