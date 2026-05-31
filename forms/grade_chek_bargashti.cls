Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
    If IsLoaded("NABZ_MOSHTARI") Then
        If IsNull(Forms![NABZ_MOSHTARI].Form.OpenArgs) Then
            Me.RecordSource = "SELECT N_SERI, vaz,BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, HES1, HES2, HES3, VAZ FROM PAY_GETD WHERE (VAZ = 5 OR VAZ = 6) AND (CUST_NO = N'" & Forms![HEAD_LST_PISHFROOSH2].Form.CUST_NO & "')"
        Else
            Me.RecordSource = "SELECT N_SERI, vaz,BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, HES1, HES2, HES3, VAZ FROM PAY_GETD WHERE (VAZ = 5 OR VAZ = 6) AND (CUST_NO = N'" & Forms![NABZ_MOSHTARI].Form.OpenArgs & "')"
        End If
    Else
            Me.RecordSource = "SELECT N_SERI, vaz,BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, HES1, HES2, HES3, VAZ FROM PAY_GETD WHERE (VAZ = 5 OR VAZ = 6) AND (CUST_NO = N'" & Forms![HEAD_LST_PISHFROOSH2].Form.CUST_NO & "')"
    End If

End Sub
