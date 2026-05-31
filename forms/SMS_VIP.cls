Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim i, ST, le As Long, frm As Form, CTL As Control, JAM As Double
Private Sub Command1_Click()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   If Me.OpenArgs = "1" Then
    Forms![BASEKNOW]![matnsms] = Me.MATN
    DoCmd.Close acForm, Me.NAME
   Else
        Set rst = frm.Form.RecordsetClone
         DoCmd.OpenForm "bun"
        DoCmd.GoToRecord acDataForm, frm.NAME, acGoTo, ST
        JAM = 0
        For i = 1 To le
             Forms![BUN].[num].CAPTION = frm.Form("COMPANY_NAME")
             Forms![BUN].Form.Refresh
             If frm.NAME = "copmanes_form" Then
                  Call ersal_sms2("مشتري محترم :" & frm.Form("COMPANY_NAME") & Chr(10) & Chr(13) & Me.MATN, frm.Form("mobile"), True)
             Else
                  Call ersal_sms(frm.Form("HES"), "مشتري محترم :" & GETHESNAME(frm.Form("HES")) & Chr(10) & Chr(13) & Me.MATN, 0, 0, True)
             End If
              rst.MoveNext
           DoCmd.GoToRecord acDataForm, frm.NAME, acNext
        Next i
        DoCmd.GoToRecord acDataForm, frm.NAME, acGoTo, ST
        DoCmd.Close acForm, "bun"
    End If
End Sub

Private Sub Command19_Click()
                     On Error Resume Next
  DoCmd.Close acForm, Me.NAME



End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                     On Error Resume Next
    If Me.OpenArgs = "1" Then
    Else
        Set frm = Screen.ActiveDatasheet
        Set CTL = Screen.ActiveControl
        ST = frm.SelTop
        le = frm.SelHeight
        If le = 0 Then
           DoCmd.OpenForm "mesageform", , , , , acDialog, "هيچ ستوني انتخاب نشده است براي ارسال اس ام اس ، ستون شماره موبايل يا كد حساب يا نام اشخاص را انتخاب كنيد"
           DoCmd.Close acForm, Me.NAME
        End If
    End If
End Sub
