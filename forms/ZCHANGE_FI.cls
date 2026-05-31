Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                    On Error Resume Next
  Dim CC   As String, rst As New ADODB.Recordset, prs As Long
  If IsNull(Me.SNDNUM1) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Forms![BASEKNOW]![Text44] = False
  DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا براي تغيير نرخها  اطمينان داريد."
  If Forms![BASEKNOW]![Text44] Then
    DoCmd.OpenForm "BUN"
    Forms![BUN].[num].CAPTION = "تغيير نرخ گروهي"
     rst.Open "SELECT     *  FROM dbo.TAKHPERS WHERE     PUT = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not rst.EOF
         Forms![BUN].[num].CAPTION = GETKALANAME(rst.Fields("TAKH_COD"))
         Forms![BUN].Form.Refresh
         If Me.DIG > 0 Then
            prs = Round((Me.SNDNUM1 + rst.Fields("PERS") / 100 * Me.SNDNUM1 + rst.Fields("BLNS")) / 10 ^ Me.DIG, 0) * (10 ^ Me.DIG)
         Else
            prs = Me.SNDNUM1 + rst.Fields("PERS") / 100 * Me.SNDNUM1 + rst.Fields("BLNS")
         End If
         If prs >= 0 Then
            rst.Fields("PRICE_M") = prs
            rst.update
         Else
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "حاصل تغيير قيمت براي اين كد منفي شد اعمال نگرديد لطفا پارامترهاي درصد و بالانس را كنترل كنيد  نام كالا : " & GETKALANAME(rst.Fields("TAKH_COD")) & "   كد گروه : " & rst.Fields("CUST_CO")
         End If
         rst.MoveNext
     Wend
      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " تغيير نرخها با موفقيت انجام گرديد"
  End If
  DoCmd.Close acForm, "BUN"
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
   If Not IsLoaded("STUF_DEF") Then
      DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub SNDNUM1_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub
