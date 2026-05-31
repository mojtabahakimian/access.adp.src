Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command5_Click()
                    On Error Resume Next
    If Not IsNull(HF1) And Not IsNull(HF2) Then
        Select Case Me.OpenArgs
          Case "GRFRO":   Call GENSANADFROOSH(Me.HF1, Me.HF2)
          Case "GRKHO":   Call SANADKHORUGMAVAD(Me.HF1, Me.HF2)
          Case "GRVRO":   Call SANADVORUDSAKHT(Me.HF1, Me.HF2)
          Case "GRKHS":   Call SANADKHORUGSAYER(Me.HF1, Me.HF2)
          Case "GRENT":   Call SANADENTEGHAL(Me.HF1, Me.HF2)
          Case "GRKHA":   Call SANADKHAD(Me.HF1, Me.HF2)
          Case "GRKHZ":   Call GENSANADKHAZ(Me.HF1, Me.HF2)
          Case "GRVDS":   Call GENSANADVD(Me.HF1, Me.HF2)
          Case "GRKHR":   Call GENSANADKHAREED(Me.HF1, Me.HF2)
          Case "GRANB":   Call GENSANADANBARGARD(Me.HF1, Me.HF2)
          Case "GENFRBR": Call gensanadbargashfroosh(Me.HF1, Me.HF2)
        End Select
    End If
End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub
