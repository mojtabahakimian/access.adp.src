Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Command1_Click()
                     On Error Resume Next
  Forms![BASEKNOW]![Text44] = True
  DoCmd.Close acForm, "msgdialog"
End Sub


Private Sub Command17_Click()
                     On Error Resume Next
  Forms![BASEKNOW]![Text44] = False
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.restore
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next

If Me.OpenArgs = "با انتخاب اين گزينه تمام اجازه هاي دسترسي براي كاربر غير فعال مي شود آيا مايليد اين كار انجام شود؟" Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If
If Me.OpenArgs = "با انتخاب اين گزينه تمام اجازه هاي دسترسي براي كاربر فعال مي شود آيا مايليد اين كار انجام شود؟" Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If
If Me.OpenArgs = "با اين شماره سريال قبلاً پذيرش انجام شده است آيا اطلاعات پذيرش قبلي بازيابي شود؟" Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If
If Me.OpenArgs = "براي اين حواله فاكتور خدمات صادر نشده است فاكتور جديد ايجاد شود؟" Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If
If Me.OpenArgs = "براي اين پذيرش حواله صادر نشده حواله جديد ايجاد شود؟ " Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If

If Me.OpenArgs = "تغيرات داده شده ثبت شود؟ در صورتيكه مايليد تغييرات ثبت نشود بعداز بستن اين پنجره كليد  اسكيپ را فشار دهيد." Then
  Me.Detail.BackColor = 125
  Me.MSG.BackColor = 125
  Me.MSG.ForeColor = 16777215
  Me.Command17.SetFocus
  Me.Repaint
End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub
