Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim CUST_NO As String


Private Sub Command36_Click()
    DoCmd.OpenForm "AZAE", , , "HES = '" & Me.OpenArgs & "'", , acDialog
End Sub

Private Sub Command53_Click()
    DoCmd.OpenForm "NABZ_MOSHTARI_PAYMENT", , , "cust_no = '" & Me.OpenArgs & "'", , acDialog
End Sub

Private Sub Command54_Click()
    DoCmd.OpenForm "tasks_company", acFormDS, , , , acDialog, Me.OpenArgs
End Sub

Private Sub Command55_Click()
    DoCmd.OpenForm "froosh_naraftah_PERS", acFormDS, , , , acDialog, Me.OpenArgs
End Sub

Private Sub Command58_Click()
    DoCmd.RunMacro "macro24"
End Sub

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
    Me.MNAM = GETHESNAME(CUST_NO)
    Me.MOSHTARI = CUST_NO & " - " & GETHESNAME(CUST_NO)
    GRD = GetGrade(CUST_NO)
    grdn = GetGradeName(GRD)
    mhesab = GetMhesab(CUST_NO)
    If mhesab = 0 Then
        Me.mhesab.ControlSource = "=" & Chr(34) & "   0 ريال" & Chr(34)
    Else
        Me.mhesab.ControlSource = "=" & Chr(34) & IIf(mhesab > 0, Format(mhesab, "##,# ريال بدهكار"), Format(mhesab * -1, "##,# ريال بستانكار")) & Chr(34)
    End If

    mpish = Format(GetMPishf(CUST_NO), "#,###")
    mhav = Format(GetMHavl(CUST_NO), "#,###")
    masnad = Format(GetMAsnad(CUST_NO), "#,###")
    DoCmd.restore

End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
