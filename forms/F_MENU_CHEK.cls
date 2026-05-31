Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = Forms![BASEKNOW]![YEA] & "0101"
  End If
  If IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If IsNull(Me.HHMOIN) Then
     MMOIN = "%"
  Else
    MMOIN = Me.HHMOIN
  End If
  Select Case Me.OpenArgs
    Case "PCHS":
               DoCmd.OpenForm "CHEK_PLISTS", , , , , , 1
               Forms![CHEK_PLISTS]![CHEK_PLIST].Form.HelpContextId = 800601
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "DCHS":
               DoCmd.OpenForm "CHEK_DLISTS", , , , , , 1
               Forms![chek_DLISTS]![CHKE_DLIST].Form.HelpContextId = 800501
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "PCHSS":
               DoCmd.OpenForm "CHEK_PLISTS", , , , , , 2
               Forms![CHEK_PLISTS]![CHEK_PLIST].Form.HelpContextId = 800602
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "DCHSS":
               DoCmd.OpenForm "CHEK_DLISTS", , , , , , 2
               Forms![chek_DLISTS]![CHKE_DLIST].Form.HelpContextId = 800502
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RCHEKD":
               DoCmd.OpenReport "CHEK_DLISTS", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RCHEKP":
               DoCmd.OpenReport "CHEK_PLIST", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "CHKM":
               DoCmd.OpenForm "CHEK_MOGUD", acFormDS, , , , , 2
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "CHKB":
               DoCmd.OpenForm "CHEK_BARGASHTI_MAIN", , , , , , 2
               Forms![CHEK_BARGASHTI_MAIN]![CHEK_BARGASHTI].Form.HelpContextId = 8008
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "CHKV":
               DoCmd.OpenForm "CHEK_VOSUL_LES", , , , , , 1
               Forms![CHEK_VOSUL_LES]![CHEK_VLIST].Form.HelpContextId = 800503
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "CHKVA":
               DoCmd.OpenForm "CHEK_VLISTALL", acFormDS
               Forms![CHEK_VLISTall].Form.HelpContextId = 800503
               DoCmd.OpenForm Me.NAME, , , , , acHidden

    Case "CHKP":
               DoCmd.OpenForm "CHK_V_PRINT", , , , , , 1
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "CVP": Calendar = vbCalGreg
               DoCmd.OpenForm "CHECK_PVLIST", acFormDS
               DoCmd.OpenForm Me.NAME, , , , , acHidden

  End Select
End Sub

Private Sub Command6_Click()
                    On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                 On Error Resume Next
 Select Case Me.OpenArgs
    Case "pchs": Me.HelpContextId = 800601

    Case "dchs": Me.HelpContextId = 800501

    Case "pchss": Me.HelpContextId = 800602

    Case "dchss": Me.HelpContextId = 800502

    Case "chkm": Me.HelpContextId = 8009

    Case "chkb": Me.HelpContextId = 8008

    Case "chkv": Me.HelpContextId = 800503

    Case "chkp": Me.HelpContextId = 8016
  End Select
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
