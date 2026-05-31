Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
      Dim rst As New ADODB.Recordset
    If KeyAscii = 27 Then
       DoCmd.Close
    End If
    If KeyAscii = 13 Then
        If PORTP = "" Then
        Else
                Dim BB(83) As Byte, ST As String, i As Integer
                Open Me.PORTP & ":19200,N,8,1" For Random As #1 Len = 84
                ST = "RQ075PR006000000AM010" & RIGHT("0000000000" & Me.Text0, 10) & "CU003364R1000R2000T1000T2000SV000SG000AD000PD0011"
                For i = 1 To Len(ST)
                  BB(i + 3) = Asc(MID(ST, i, 1))
                Next i
                BB(0) = 48
                BB(1) = 48
                BB(2) = 56
                BB(3) = 48
                Put #1, , BB
                Close #1
       End If
       rst.Open "SELECT     TOP 100 PERCENT MOIN_VAR FROM dbo.HEAD_LST WHERE (TAG = 13) AND (MABL_VAR > 0) And (DEPATMAN = " & Forms![HEAD_LST_FROOSH22]![DEPATMAN] & ") And (NUMBER <> " & Forms![HEAD_LST_FROOSH22]![NUMBER] & ") And (Not (MOIN_VAR Is Null)) ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            Forms![HEAD_LST_FROOSH22]![MOIN_VAR] = rst.Fields("MOIN_VAR")
       End If
       Forms![HEAD_LST_FROOSH22]![MABL_VAR] = Me.Text0
       Forms![HEAD_LST_FROOSH22]![CUST_NO].SetFocus
       Call APLAYTAKH(Forms![HEAD_LST_FROOSH22]![NUMBER], 2, Forms![HEAD_LST_FROOSH22]![M_NAGHD], Forms![HEAD_LST_FROOSH22]![MABL_VAR], Forms![HEAD_LST_FROOSH22]![MABL_HAV], Forms![HEAD_LST_FROOSH22]![TICMBAA])
       DoCmd.Close acForm, Me.NAME
    End If
End Sub
Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
Dim infileopen As Boolean
Dim incmd$, batch$, Q$
   If IsLoaded("HEAD_LST_FROOSH22") Then
        If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
            DoCmd.Close acForm, Me.NAME
        End If
    End If
    If IsLoaded("HEAD_LST_FROOSH2") Then
        If Not Forms![HEAD_LST_FROOSH2].Form.AllowEdits Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
            DoCmd.Close acForm, Me.NAME
        End If
    End If
    'خواندن پورت پيش فرض پوز
    PORTP = ""
    Open "c:\POSCONF.CFG" For Input As #1
    If Err = 0 Then
        infileopen = True
        Do While Not EOF(1)
            Line Input #1, incmd$
            If incmd$ = UCurrentUser() Then
                Line Input #1, incmd$
                Me.PORTP.DefaultValue = """" & incmd$ & """"
                Close #1
                Exit Sub
            End If
       Loop
    End If
    Close #1

End Sub

Private Sub Text0_KeyPress(KeyAscii As Integer)
        If KeyAscii = 61 Then
            KeyAscii = 0
            Me.Text0.Text = "-"
        Else
         sendkeyu (KeyAscii)
        End If
End Sub
