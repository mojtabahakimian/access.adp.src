Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command206_Click()
                     On Error Resume Next
  Select Case Text210
  Case "form":   Forms![formeleman]![Text10].BackColor = Text208.BackColor
  Case "matn": Forms![formeleman]![Text14].BackColor = Text208.BackColor
                      Forms![formeleman]![Text14].ForeColor = Text208.BackColor
                      Forms![formeleman]![Label24].ForeColor = Text208.BackColor

   Case "zamin": Forms![formeleman]![Text12].BackColor = Text208.BackColor
                         Forms![formeleman]![Text12].ForeColor = Text208.BackColor
                         Forms![formeleman]![Label24].BackColor = Text208.BackColor
   Case "CHK_DATE": Forms![Format]![DATE_COLOR].BackColor = Text208.BackColor
                         Forms![Format]![DATE_COLOR] = Text208.BackColor
   Case "CHK_ALPHA": Forms![Format]![ALPHA_COLOR].BackColor = Text208.BackColor
                         Forms![Format]![ALPHA_COLOR] = Text208.BackColor
   Case "CHK_VAJH": Forms![Format]![VAJH_COLOR].BackColor = Text208.BackColor
                         Forms![Format]![VAJH_COLOR] = Text208.BackColor
   Case "CHK_NUM": Forms![Format]![NUM_COLOR].BackColor = Text208.BackColor
                         Forms![Format]![NUM_COLOR] = Text208.BackColor

  End Select
  DoCmd.Close
End Sub


Private Sub Command207_Click()
                      On Error Resume Next
   DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Option100_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label101.BackColor
End Sub


Private Sub Option102_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label103.BackColor

End Sub

Private Sub Option104_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label105.BackColor
End Sub


Private Sub Option106_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label107.BackColor
End Sub


Private Sub Option108_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label109.BackColor
End Sub


Private Sub Option110_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label111.BackColor
End Sub


Private Sub Option112_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label113.BackColor
End Sub


Private Sub Option114_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label115.BackColor
End Sub


Private Sub Option116_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label117.BackColor
End Sub


Private Sub Option118_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label119.BackColor
End Sub


Private Sub Option120_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label121.BackColor
End Sub


Private Sub Option122_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label123.BackColor
End Sub


Private Sub Option124_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label125.BackColor
End Sub


Private Sub Option126_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label127.BackColor
End Sub


Private Sub Option128_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label129.BackColor
End Sub


Private Sub Option130_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label131.BackColor
End Sub


Private Sub Option132_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label133.BackColor
End Sub


Private Sub Option134_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label135.BackColor
End Sub


Private Sub Option136_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label137.BackColor
End Sub


Private Sub Option138_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label139.BackColor
End Sub


Private Sub Option140_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label141.BackColor
End Sub


Private Sub Option142_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label143.BackColor
End Sub


Private Sub Option144_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label145.BackColor
End Sub


Private Sub Option146_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label147.BackColor
End Sub


Private Sub Option148_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label149.BackColor
End Sub


Private Sub Option150_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label151.BackColor
End Sub


Private Sub Option152_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label153.BackColor
End Sub


Private Sub Option154_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label155.BackColor
End Sub


Private Sub Option156_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label157.BackColor
End Sub


Private Sub Option158_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label159.BackColor
End Sub


Private Sub Option160_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label161.BackColor
End Sub


Private Sub Option162_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label163.BackColor
End Sub


Private Sub Option164_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label165.BackColor
End Sub


Private Sub Option166_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label167.BackColor

End Sub


Private Sub Option168_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label169.BackColor
End Sub


Private Sub Option170_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label171.BackColor
End Sub


Private Sub Option172_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label173.BackColor

End Sub


Private Sub Option174_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label175.BackColor
End Sub


Private Sub Option176_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label177.BackColor
End Sub


Private Sub Option178_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label179.BackColor
End Sub


Private Sub Option180_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label181.BackColor
End Sub


Private Sub Option182_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label183.BackColor
End Sub


Private Sub Option184_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label185.BackColor
End Sub


Private Sub Option186_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label187.BackColor
End Sub


Private Sub Option188_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label189.BackColor
End Sub


Private Sub Option190_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label191.BackColor
End Sub


Private Sub Option192_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label193.BackColor
End Sub


Private Sub Option194_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label195.BackColor
End Sub


Private Sub Option196_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label197.BackColor
End Sub


Private Sub Option198_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label199.BackColor
End Sub


Private Sub Option200_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label201.BackColor
End Sub


Private Sub Option202_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label203.BackColor
End Sub


Private Sub Option204_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label205.BackColor
End Sub


Private Sub Option93_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label94.BackColor
End Sub


Private Sub Option96_GotFocus()
                      On Error Resume Next
   Text208.BackColor = Label97.BackColor
End Sub


Private Sub Option98_GotFocus()
                     On Error Resume Next
  Text208.BackColor = Label99.BackColor
End Sub
