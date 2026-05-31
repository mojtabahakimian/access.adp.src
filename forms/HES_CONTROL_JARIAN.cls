Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  Me.RecordSource = "SELECT DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T, Sum([BED]-[bes]) AS MOGH FROM DEED_HED INNER JOIN DEED_DTL ON (DEED_HED.N_S = DEED_DTL.N_S) AND (DEED_HED.N_S = DEED_DTL.N_S) WHERE (((DEED_HED.DATE_S) Between " & [Forms]![F_MENU_SANAD_MOGHA]![DT1] & " And " & [Forms]![F_MENU_SANAD_MOGHA]![DT2] & "))GROUP BY DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T HAVING (((DEED_DTL.HES_K)=" & [Forms]![BASEKNOW]![CONKAL] & ") AND ((DEED_DTL.HES_M) Like '" & [Forms]![F_MENU_SANAD_MOGHA]![MKALA] & "') AND ((DEED_DTL.HES_T)<99999996) AND ((Sum([BED]-[bes]))<>0))"
End Sub
