Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        On Error Resume Next
        Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
      '  RST.Open "TALARS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        i = 1
      '  While Not RST.EOF()
           Me("T" & i).CAPTION = "جلسات همه" 'RST.Fields("TAL_NAME")
           Me("T2").CAPTION = "جلسات من"  'RST.Fields("TAL_NAME")
           i = i + 1
        ' RST.MoveNext
      '  Wend
        Me.Frame13 = UMonth(FARSIDATE(DATE))
         'نمايش اتاقهاي داراي شماره
        For i = 1 To 4
           If Me("T" & i).CAPTION = "0" Then
                Me("T" & i).Visible = False
           Else
              Me("T" & i).Visible = True
           End If
        Next i
        Dim DT1 As Long, DT2 As Long
        DT1 = Left(FARSIDATE(DATE), 6) * 100
        DT2 = Left(FARSIDATE(DATE), 6) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])
  DoCmd.Maximize
End Sub

Private Sub Frame13_Click()
       On Error Resume Next
        Dim DT1 As Long, DT2 As Long
        DT1 = (Left(FARSIDATE(DATE), 4) * 100 + Me.Frame13) * 100
        DT2 = (Left(FARSIDATE(DATE), 4) * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])
End Sub


Public Sub SHOWRAKM(DT1 As Long, DT2 As Long, UID As Integer)
       On Error Resume Next
   Dim rst As New ADODB.Recordset, dm As Integer, DMS As String, tti As Integer
   Dim i, j As Integer
        For i = 1 To 4
           If Me("T" & i).CAPTION = "0" Then
                Me("T" & i).Visible = False
           Else
              Me("T" & i).Visible = True
           End If
        Next i
   For i = 1 To 31
     For j = 1 To 4
        Me("T" & RIGHT("0" & i, 2) & j).BackColor = 12632256
        Me("T" & RIGHT("0" & i, 2) & j).SpecialEffect = 1
        Me("T" & RIGHT("0" & i, 2) & j).Width = 1029
        Me("T" & RIGHT("0" & i, 2) & j).Visible = True
        Select Case j
            Case 1: Me("T" & RIGHT("0" & i, 2) & j).CAPTION = "شب"
            Case 2: Me("T" & RIGHT("0" & i, 2) & j).CAPTION = "عصر"
            Case 3: Me("T" & RIGHT("0" & i, 2) & j).CAPTION = "ظهر"
            Case 4: Me("T" & RIGHT("0" & i, 2) & j).CAPTION = "صبح"
        End Select
     Next j
   Next i
        If Me.TOPT = 0 Then
            rst.Open "SELECT * FROM crmevents WHERE  miting = -1 and  NEXT_DATE between  " & DT1 & " and " & DT2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Else
            rst.Open "SELECT * FROM crmevents WHERE  miting = -1 and  USERID = " & UID & " AND NEXT_DATE between  " & DT1 & " and " & DT2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        i = 1
        While Not rst.EOF()
            If rst.Fields("NEXT_time") > 1 And rst.Fields("NEXT_time") <= 1100 Then
                tti = 4
            Else
                If rst.Fields("NEXT_time") > 1100 And rst.Fields("NEXT_time") <= 1500 Then
                    tti = 3
                Else
                    If rst.Fields("NEXT_time") > 1500 And rst.Fields("NEXT_time") < 1900 Then
                        tti = 2
                    Else
                        tti = 1
                    End If
                End If
            End If

            DMS = RIGHT("0" & UDay(rst.Fields("NEXT_DATE")), 2)
            dm = DMS
            If Len(Me("T" & DMS & tti).CAPTION) < 4 Then
                Me("T" & DMS & tti).CAPTION = ""
            End If
            Me("T" & DMS & tti).ControlTipText = Left(GETCRMCOMPNAME(rst.Fields("idc")) & Chr(13) & Chr(10) & GETCRMmodir(rst.Fields("idc")) & Chr(13) & Chr(10) & "(" & rst.Fields("NEXT_time") & ")" & Chr(13) & Chr(10) & IIf(IsNull(rst.Fields("COMMENT")), " ", rst.Fields("COMMENT")) & Chr(13) & Chr(10) & IIf(IsNull(rst.Fields("SALER")), " ", rst.Fields("SALER")) & Chr(13) & Chr(10) & IIf(IsNull(rst.Fields("BUYER")), " ", rst.Fields("BUYER")) & Chr(13) & Chr(10) & Me("T" & DMS & tti).ControlTipText, 254)
            Me("T" & DMS & tti).CAPTION = GETCRMCOMPNAME(rst.Fields("idc")) & Chr(13) & Chr(10) & GETCRMmodir(rst.Fields("idc")) & Chr(13) & Chr(10) & "(" & rst.Fields("NEXT_time") & ")" & Chr(13) & Chr(10) & Me("T" & DMS & tti).CAPTION
            Select Case tti
               Case 1: Me("T" & DMS & tti).BackColor = 13408767
               Case 2: Me("T" & DMS & tti).BackColor = 52377
               Case 3: Me("T" & DMS & tti).BackColor = 65535
               Case 4: Me("T" & DMS & tti).BackColor = 13421619
            End Select
           i = i + 1
           rst.MoveNext
        Wend
End Sub

Public Function ELLAT(EI As Integer) As String
          On Error Resume Next
 Select Case EI
    Case 1: ELLAT = "عروسي"
    Case 2: ELLAT = "مهماني"
    Case 3: ELLAT = "سمينار"
    Case 4: ELLAT = "بزرگداشت"
    Case 5: ELLAT = "عزاداري"
    Case 6: ELLAT = "ساير"
  End Select
End Function


Private Sub T1_Click()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])

End Sub

Private Sub T2_Click()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])

End Sub

Private Sub T3_Click()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])

End Sub

Private Sub T4_Click()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])

End Sub

Private Sub Text35_AfterUpdate()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])
End Sub
Private Function GETCRMCOMPNAME(idc As Double) As String
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
       Dim rst As New ADODB.Recordset
       rst.Open "select COMPANY_NAME from  COPMANES where  id = " & idc, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETCRMCOMPNAME = rst.Fields("COMPANY_NAME")
       Else
            GETCRMCOMPNAME = ""
       End If
End Function
Private Function GETCRMmodir(idc As Double)
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
       Dim rst As New ADODB.Recordset
       rst.Open "select MANAGER from  COPMANES where  id = " & idc, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETCRMmodir = rst.Fields("MANAGER")
       Else
            GETCRMmodir = ""
       End If
End Function




Private Sub TOPT_Change()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])

End Sub

Private Sub TOPT_Click()
          On Error Resume Next
       Dim DT1 As Long, DT2 As Long
        DT1 = (Text35 * 100 + Me.Frame13) * 100
        DT2 = (Text35 * 100 + Me.Frame13) * 100 + 32
        Call SHOWRAKM(DT1, DT2, Forms![BASEKNOW]![USERCOD])
End Sub
