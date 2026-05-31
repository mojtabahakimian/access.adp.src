Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 11
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(CODE1) / 567 > 6.3 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(CODE1) / 567 > 6.3 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(CODE1) / 567 > 6.3 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(CODE1) / 567 > 6.3 Then
                    Me.CODE1.FontSize = Me.FontSize - 1
               Else
                     Me.CODE1.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.CODE1.FontSize = Me.FontSize + 1
            End If
        Else
            Me.CODE1.FontSize = Me.FontSize + 1
        End If
    Else
        Me.CODE1.FontSize = Me.FontSize + 1
    End If
Me.FontSize = 8
Me.FontName = Me.mabkn.FontName
Me.FontBold = Me.mabkn.FontBold
    If Me.TextWidth(mabkn) / 567 > 2.476 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(mabkn) / 567 > 2.476 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(mabkn) / 567 > 2.476 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(mabkn) / 567 > 2.476 Then
                    Me.mabkn.FontSize = Me.FontSize - 1
               Else
                     Me.mabkn.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.mabkn.FontSize = Me.FontSize + 1
            End If
        Else
            Me.mabkn.FontSize = Me.FontSize + 1
        End If
    Else
        Me.mabkn.FontSize = Me.FontSize + 1
    End If
Me.FontSize = 8
Me.FontName = Me.Text279.FontName
Me.FontBold = Me.Text279.FontBold
    If Me.TextWidth(Text279) / 567 > 2.476 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(Text279) / 567 > 2.476 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(Text279) / 567 > 2.476 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(Text279) / 567 > 2.476 Then
                    Me.Text279.FontSize = Me.FontSize - 1
               Else
                     Me.Text279.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.Text279.FontSize = Me.FontSize + 1
            End If
        Else
            Me.Text279.FontSize = Me.FontSize + 1
        End If
    Else
        Me.Text279.FontSize = Me.FontSize + 1
    End If
Me.FontSize = 8
Me.FontName = Me.mabkbt.FontName
Me.FontBold = Me.mabkbt.FontBold
    If Me.TextWidth(mabkbt) / 567 > 2.582 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(mabkbt) / 567 > 2.582 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(mabkbt) / 567 > 2.582 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(mabkbt) / 567 > 2.582 Then
                    Me.mabkbt.FontSize = Me.FontSize - 1
               Else
                     Me.mabkbt.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.mabkbt.FontSize = Me.FontSize + 1
            End If
        Else
            Me.mabkbt.FontSize = Me.FontSize + 1
        End If
    Else
        Me.mabkbt.FontSize = Me.FontSize + 1
    End If
 Me.FontSize = 8
Me.FontName = Me.Text275.FontName
Me.FontBold = Me.Text275.FontBold
    If Me.TextWidth(Text275) / 567 > 2.582 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(Text275) / 567 > 2.582 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(Text275) / 567 > 2.582 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(Text275) / 567 > 2.582 Then
                    Me.Text275.FontSize = Me.FontSize - 1
               Else
                     Me.Text275.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.Text275.FontSize = Me.FontSize + 1
            End If
        Else
            Me.Text275.FontSize = Me.FontSize + 1
        End If
    Else
        Me.Text275.FontSize = Me.FontSize + 1
    End If

  Me.FontSize = 8
Me.FontName = Me.MABL_K.FontName
Me.FontBold = Me.MABL_K.FontBold
    If Me.TextWidth(MABL_K) / 567 > 2.582 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(MABL_K) / 567 > 2.582 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(MABL_K) / 567 > 2.582 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(MABL_K) / 567 > 2.582 Then
                    Me.MABL_K.FontSize = Me.FontSize - 1
               Else
                     Me.MABL_K.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.MABL_K.FontSize = Me.FontSize + 1
            End If
        Else
            Me.MABL_K.FontSize = Me.FontSize + 1
        End If
    Else
        Me.MABL_K.FontSize = Me.FontSize + 1
    End If


  Me.FontSize = 8
Me.FontName = Me.Text271.FontName
Me.FontBold = Me.Text271.FontBold
    If Me.TextWidth(Text271) / 567 > 2.582 Then
        Me.FontSize = Me.FontSize - 1
        If Me.TextWidth(Text271) / 567 > 2.582 Then
            Me.FontSize = Me.FontSize - 1
            If Me.TextWidth(Text271) / 567 > 2.582 Then
               Me.FontSize = Me.FontSize - 1
               If Me.TextWidth(Text271) / 567 > 2.582 Then
                    Me.Text271.FontSize = Me.FontSize - 1
               Else
                     Me.Text271.FontSize = Me.FontSize + 1
               End If
            Else
                 Me.Text271.FontSize = Me.FontSize + 1
            End If
        Else
            Me.Text271.FontSize = Me.FontSize + 1
        End If
    Else
        Me.Text271.FontSize = Me.FontSize + 1
    End If
End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
  On Error Resume Next
 Me.HR.CAPTION = ALPHANUM(Me.Text279) & " " & "ريال"
       If Me.DEPART = "" Or IsNull(Me.DEPART) Then
            Me.DEPART.Visible = False
            Me.DEPNAME.Visible = False
       End If
  Call Showemza(Me, "FFRB_FROOSHTX", "FFRB_ANBTX", "FFRB_HESABTX")


End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
