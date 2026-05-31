Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim BGIDD As Long
Private Sub Command3_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command4_Click()
                On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.mah) Or IsNull(Me.SAL) Then
       DoCmd.OpenForm "mesageform", , , , , acDialog, "پارامتر ها کامل نيست ماه و سال را وارد کنيد"
    Else
        rst.Open "SELECT BGID FROM BUGET_MAIN WHERE BGMAH = " & Me.mah & " AND BGSAL = " & Me.SAL & " ORDER BY BGID DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "براي اين ماه و سال قبلا بودجه تعريف شده است همان را اصلاح ميکنيد؟"
            If Forms![BASEKNOW]![Text44] Then
                BGIDD = rst.Fields("BGID")
            Else
                GoTo NOO
            End If
        Else
NOO:        If BGIDD = 0 Then
                BGIDD = GetLIDD("BUGET_MAIN", "BGID")
               CurrentProject.Connection.Execute ("INSERT INTO BUGET_MAIN VALUES(" & BGIDD & "," & Me.bgdate & "," & Me.mah & "," & Me.SAL & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & IIf(IsNull(Forms![BASEKNOW]![hesnaghd]), " ", Forms![BASEKNOW]![hesnaghd]) & "'," & Forms![BASEKNOW]![USERCOD] & ")")
            Else
                CurrentProject.Connection.Execute ("UPDATE BUGET_MAIN SET BGCDATE =  " & Me.bgdate & ",  BGMAH = " & Me.mah & ", BGSAL = " & Me.SAL & " WHERE BGID = " & BGIDD)
            End If
        End If
        Dim DTT As Long
        DTT = Me.SAL * 100 + Me.mah
        If DTT <= Val(Left(bgdate, 6)) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "بودجه را صرفا براي ماه هاي بعد ميتوان تعريف کرد"
        Else
            DoCmd.OpenForm "buget_defaul", , , , , acDialog, BGIDD
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"


End Sub

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    'Me.WebBrowser0.OBJECT.Navigate CurrentProject.PATH & "\tink.gif"
   ' Me.AniGIF3.Play

    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    NESW = Me.WindowWidth / 15390
    NESH = Me.WindowHeight / 7140
    Me.Command4.Left = Me.Command4.Left * NESW
    Me.Detail.Height = Me.Detail.Height * NESH
    Me.Command4.Top = Me.Command4.Top * NESH
    Me.Command3.Left = Me.Command3.Left * NESW
    Me.Command3.Top = Me.Command3.Top * NESH
    Me.SAL.Top = Me.SAL.Top * NESH
    Me.SAL.Left = Me.SAL.Left * NESW
    Me.mah.Top = Me.mah.Top * NESH
    Me.mah.Left = Me.mah.Left * NESW
    Me.Label10.Top = Me.Label10.Top * NESH
    Me.Label10.Left = Me.Label10.Left * NESW
    Me.Label1.Top = Me.Label1.Top * NESH
    Me.Label1.Left = Me.Label1.Left * NESW
    Me.bgdate.Left = Me.bgdate.Left * NESW
    Me.Label6.Left = Me.Label6.Left * NESW
    Me.Label8.Left = (Me.WindowWidth / 2) - (Me.Label8.Width / 2)
   ' Me.AniGIF3.Play
End Sub
