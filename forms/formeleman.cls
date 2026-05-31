Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Compare Database
Option Explicit
 Dim frm As Form
 Dim CTL As Control
 Dim formn As String
 Dim fil As String
Sub Command12_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "form"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Public Sub Command13_Click()
      Dim rs As New ADODB.Recordset, intI As Integer, fld As Field
                        On Error Resume Next
'   DoCmd.Close Forms(formn).Name
'   DoCmd.OpenForm Forms(formn).Name
'    Set rs = Forms(formn).RecordsetClone
'    For intI = 0 To rs.Fields.Count - 1
'          If rs.Fields(intI).Name = fil Then
'             fld = rs.Fields(intI)
'          End If
'     Next intI
      frm.Detail.BackColor = Text10.BackColor
      frm.FormHeader.BackColor = Text10.BackColor
   '   fld = Me(Forms(formn))
      CTL.FontSize = Label24.FontSize
      CTL.FontName = Label24.FontName
      CTL.FontItalic = Label24.FontItalic
      CTL.FontUnderline = Label24.FontUnderline
      CTL.ForeColor = Text14.BackColor
      CTL.BackColor = Text12.BackColor
      CTL.SpecialEffect = Label24.SpecialEffect
      CTL.Top = Text45 * 1000
      CTL.Left = Text47 * 1000
      CTL.Width = Text49 * 1000
      CTL.Height = Text51 * 1000
      DoCmd.Save acForm, frm.NAME
End Sub

Private Sub Command14_Click()
   DoCmd.Close acForm, "formeleman"
End Sub

Private Sub Command17_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "zamin"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click

End Sub

Private Sub Command30_Click()
End Sub

Private Sub Command33_Click()
 On Error GoTo Err_Command12_Click
   DoCmd.OpenForm "bordcolor"
   [Forms]![bordcolor].[Text210] = "matn"
Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                       On Error Resume Next
    Set frm = Screen.ActiveForm
    Set CTL = Screen.ActiveControl
    Text10.BackColor = frm.Detail.BackColor
    formn = frm.NAME
    fil = CTL.NAME
    Text10.ForeColor = frm.Detail.BackColor
    Text12.BackColor = CTL.BackColor
    Text12.ForeColor = CTL.BackColor
    Text14.ForeColor = CTL.ForeColor
    Text14.BackColor = CTL.ForeColor
    Text18 = CTL.FontName
    Text21 = CTL.FontSize
    Text45 = CTL.Top / 1000
    Text47 = CTL.Left / 1000
    Text49 = CTL.Width / 1000
    Text51 = CTL.Height / 1000
    If CTL.FontWeight >= 650 Then
       Toggle27 = True
    Else
      Toggle27 = False
    End If
    Toggle28 = CTL.FontItalic
    Toggle29 = CTL.FontUnderline
    Label24.SpecialEffect = CTL.SpecialEffect
    Label24.FontWeight = CTL.FontWeight
    Label24.FontSize = CTL.FontSize
    Label24.FontName = CTL.FontName
    Label24.FontItalic = CTL.FontItalic
    Label24.ForeColor = CTL.ForeColor
    Label24.BackColor = CTL.BackColor
    Label24.FontUnderline = CTL.FontUnderline
    Select Case Label24.SpecialEffect
      Case 0: Toggle42 = True
      Case 1: Toggle41 = True
      Case 2: Toggle40 = True
      Case 3: Toggle44 = True
      Case 4: Toggle43 = True
    End Select
   '   DoCmd.OpenForm frm.Name, acDesign
End Sub



Private Sub Text12_AfterUpdate()
                        On Error Resume Next
     Label24.BackColor = Text12.BackColor

End Sub

Private Sub text14_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
    Label24.ForeColor = Text14.ForeColor

End Sub


Private Sub Text18_AfterUpdate()
                     On Error Resume Next
  Label24.FontName = Text18
  Label24.FontSize = Text21
End Sub

Private Sub Text21_AfterUpdate()
                     On Error Resume Next
  Label24.FontName = Text18
  If Text21 < 128 And Text21 > 7 Then
    Label24.FontSize = Text21
  Else
    Label24.FontSize = 127
    Text21 = 127
  End If
End Sub

Private Sub Toggle27_Click()
                        On Error Resume Next
 If Toggle27 Then
     Label24.FontWeight = 650
  Else
      Label24.FontWeight = 100
  End If
End Sub


Private Sub Toggle28_Click()
                       On Error Resume Next
    Label24.FontItalic = Toggle28
End Sub


Private Sub Toggle29_Click()
                       On Error Resume Next
   Label24.FontUnderline = Toggle29
End Sub


Private Sub Toggle40_Click()
                    On Error Resume Next
 Label24.SpecialEffect = 2
  Toggle40 = True
  Toggle41 = False
  Toggle42 = False
  Toggle43 = False
  Toggle44 = False

End Sub


Private Sub Toggle41_Click()
                     On Error Resume Next
  Label24.SpecialEffect = 1
  Toggle40 = False
  Toggle41 = True
  Toggle42 = False
  Toggle43 = False
  Toggle44 = False
End Sub


Private Sub Toggle42_Click()
                     On Error Resume Next
  Label24.SpecialEffect = 0
  Toggle40 = False
  Toggle41 = False
  Toggle42 = True
  Toggle43 = False
  Toggle44 = False
End Sub


Private Sub Toggle43_Click()
                     On Error Resume Next
  Label24.SpecialEffect = 4
  Label24.BorderWidth = 4
  CTL.BorderWidth = 4
  Toggle40 = False
  Toggle41 = False
  Toggle42 = False
  Toggle43 = True
  Toggle44 = False
End Sub


Private Sub Toggle44_Click()
                     On Error Resume Next
  Label24.SpecialEffect = 3
  Toggle40 = False
  Toggle41 = False
  Toggle42 = False
  Toggle43 = False
  Toggle44 = True
End Sub
