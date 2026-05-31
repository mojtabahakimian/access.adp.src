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
    Me.RecordSource = "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, CUST_HESAB.NAME, HEAD_LST.DATE_N, HEAD_LST.USER_NAME, TAGCOD.BARGAH, CUST_HESAB.hes FROM HEAD_LST INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes INNER JOIN TAGCOD ON HEAD_LST.TAG = TAGCOD.CODE LEFT OUTER JOIN OTHER_DTL ON HEAD_LST.NUMBER = OTHER_DTL.NUMBER AND HEAD_LST.TAG = OTHER_DTL.TAG WHERE   (dbo.HEAD_LST.TAG = 1 OR  dbo.HEAD_LST.TAG = 2) and (OTHER_DTL.NUMBER IS NULL) AND (HEAD_LST.DATE_N BETWEEN " & Forms![F_MENU_DATE]![DT1] & " AND " & Forms![F_MENU_DATE]![DT2] & ")"
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    Select Case Me!TAG
        Case 2:
            DoCmd.OpenForm "head_lst_froosh22", , , , , , "NUMBER = " & Me.NUMBER
        Case 1:
            DoCmd.OpenForm "head_lst_rasid", , , , , , "NUMBER = " & Me.NUMBER
    End Select
End Sub
