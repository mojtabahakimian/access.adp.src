Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
Dim PSK1, PSM1, PST1, PST2, PST3, PST4
                     On Error Resume Next
     Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
    If IsNull(PST2) Then
        Me.RecordSource = "SELECT  HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR FROM dbo.SALARY_BEDS(" & UMonth(Forms![WORKHEAD]![WDATE]) & ", " & PSK1 & "," & Forms![BASEKNOW]![YEA] & ") SALARY_BEDS WHERE  HES_T = " & Forms![WORKHEAD]![WORKING_SUB2].Form![PCODE] & " ORDER BY BES DESC, BED DESC"
    Else
        Me.RecordSource = "SELECT  HES_K, HES_M, HES_T, HES_T2, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR FROM dbo.SALARY_BEDS2(" & UMonth(Forms![WORKHEAD]![WDATE]) & ", " & PSK1 & "," & PSM1 & "," & PST1 & "," & Forms![BASEKNOW]![YEA] & ") SALARY_BEDS2 WHERE  HES_T2 = " & Forms![WORKHEAD]![WORKING_SUB2].Form![PCODE] & " ORDER BY BES DESC, BED DESC"
    End If
End Sub
