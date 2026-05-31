Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim CUST_NO As String


Private Sub Form_Close()
    DoCmd.Close acForm, "F_MENU_KOL_MOIN_TAFZIL"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
       Forms![BASEKNOW]![hhwin] = Me.hWnd
    Dim GRD As Integer, GRDV As Integer, i, ETB As Integer, OKETEB As Boolean, OKETEBf As Boolean, FC As Integer, si As Integer
    Dim rst As New ADODB.Recordset, RSTV As New ADODB.Recordset
    Dim FUNC As String
    Dim PAYAM As String
    Dim vhes As String
    CUST_NO = Me.OpenArgs
    Me.MOSHTARI = CUST_NO & " - " & GETHESNAME(CUST_NO)
    vhes = GetVisitHES(CUST_NO)
    GRDV = GetGrade(vhes)
    GRD = GetGrade(CUST_NO)
    PAYAM = "  گريد کارشناس : " & GetGradeName(GRDV) & Chr(13) & PAYAM & Chr(13) & "مانده حساب شخصي: " & Format(GetMhesab(CUST_NO), "#,###") & Chr(13)
    PAYAM = PAYAM & Chr(10) & Chr(13) & "مانده حساب کارشناس : " & Format(GetMhesabF(vhes), "#,###") & Chr(13) & "مانده پيش فاکتورهاي تاييد شده : " & Format(GetMPishfF(vhes), "#,###") & Chr(13) & "جمع حواله هاي بارگيري نشده : " & Format(GetMHavlF(vhes), "#,###") & Chr(13) & "مانده اسناد پاس نشده: " & Format(GetMAsnadF(vhes), "#,###") & Chr(10) & "Mr.Correct..."
    PAYAM = PAYAM & " OK"
    Me.MSG.ControlSource = "=" & Chr(34) & PAYAM & Chr(34)
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
