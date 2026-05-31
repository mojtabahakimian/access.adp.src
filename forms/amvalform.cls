Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
                 On Error Resume Next
    Me.PIC.Picture = ""
    Me.picl.Picture = ""
    If Not Me.NewRecord Then
       If Dir(Forms![BASEKNOW]![BACKPATH] & "\amval\" & barchasb & ".JPG") <> "" Then
           Me.PIC.Picture = Forms![BASEKNOW]![BACKPATH] & "\amval\" & barchasb & ".JPG"
           Me.picl.Picture = Forms![BASEKNOW]![BACKPATH] & "\amval\" & barchasb & ".JPG"
       Else
         Me.PIC.Picture = ""
         Me.picl.Picture = ""
       End If
    End If
    If Not IsNull(Me.barchasb) Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
   End If
End Sub
Private Sub ESLAH_Click()
Dim dt As Date
                 On Error Resume Next
  If Not IsNull(Me.barchasb) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_AMVAL   (barchasb, aname, calmethod, rate, hesab, khdate, price, place, arzesh, astatus, Description, USER_NAME , UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT  barchasb, aname, calmethod, rate, hesab, khdate, price, place, arzesh, astatus, Description, USER_NAME ," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.AMVAL WHERE (barchasb = " & Me.barchasb & " ) ")
    End If
    Me.AllowDeletions = True
    Me.AllowEdits = True
    If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
        Me.USER_NAME = UCurrentUser()
        DoCmd.RunCommand acCmdSaveRecord
    End If
   End If
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub



Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
