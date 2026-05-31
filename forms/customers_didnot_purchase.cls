Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Open(CANCEL As Integer)
Forms![BASEKNOW]![hhwin] = Me.hWnd
    Me.RecordSource = "SELECT customers_whodid_purchase.CUST_NO, CUST_HESAB.* FROM CUST_HESAB LEFT OUTER JOIN customers_whodid_purchase(" & Forms![F_MENU_DATE]![DT1] & ", " & Forms![F_MENU_DATE]![DT2] & ") customers_whodid_purchase ON CUST_HESAB.hes = customers_whodid_purchase.CUST_NO WHERE (CUST_HESAB.hes LIKE N'" & Forms![BASEKNOW]![BEDEHKAR] & "-%') AND (customers_whodid_purchase.CUST_NO IS NULL)"
End Sub

Private Sub NAME_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "sms_vip", , , , , acDialog, "Q_FROOSH_CUSTOMER"
End Sub
Private Sub SumOfTAKHFIF_DblClick(CANCEL As Integer)
  DoCmd.OpenReport "R_FROOSH_DAYLY_person_mba2", acViewPreview, , "sader = 1 "
End Sub
