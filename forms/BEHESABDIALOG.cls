Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Command1_Click()
                    On Error Resume Next
 If IsNull(Me.HHMOIN) Or Me.HHMOIN = "" Then
  DoCmd.OpenForm "mesag", , , , , acDialog, "حسابي كه چك بايد به آن واگذار شود مشخص نگرديده است...!"
 Else
  Forms![BASEKNOW]![Text44] = True
  DoCmd.OpenForm Me.NAME, , , , , acHidden
 End If
End Sub


Private Sub Command17_Click()
                      On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DTS_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DTS, False)
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  If Not IsNull(Me.OpenArgs) Then
     Me.HHMOIN = Me.OpenArgs
  End If
  Me.HHMOIN.RowSource = " SELECT     hes, NAME,hes FROM CUST_HESAB WHERE (dbo.GETKOL(HES) = " & [Forms]![BASEKNOW]![BANKHA] & ")"
'  Me.HHMOIN.RowSource = "SELECT CAST(N_KOL AS nvarchar) + N'-' + CAST(NUMBER AS nvarchar) + N'-' + CAST(TNUMBER AS nvarchar) AS hes, NAME, TNUMBER AS Expr1 FROM TDETA_HES WHERE     ( ((TDETA_HES.N_KOL) =" & [Forms]![baseknow]![BANKHA] & "))"
  Me.HHMOIN.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub
