Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CODE_AfterUpdate()
  Dim rst As New ADODB.Recordset, VAMHES As String
  Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
                  On Error Resume Next
  If Not IsNull(Me.CODE) Then
    rst.Open "SELECT MAX(VAM_ID) AS MV  FROM PVAM WHERE CODE = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MV")) Then
      Me.VAM_ID = rst.Fields("MV") + 1
    Else
      Me.VAM_ID = 1
    End If
   Call GETTAF3(Forms![BASEKNOW]![PERVAM], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
   If IsNull(CTAF2) Then
        CTAF = CLng(CODE)
        VAMHES = CKOL & "-" & CMOIN & "-" & CTAF
    Else
        If IsNull(CTAF3) Then
            CTAF2 = CLng(CODE)
            VAMHES = CKOL & "-" & CMOIN & "-" & CTAF & "-" & CTAF2
        Else
            If IsNull(CTAF4) Then
                CTAF3 = CLng(CODE)
                VAMHES = CKOL & "-" & CMOIN & "-" & CTAF & "-" & CTAF2 & "-" & CTAF3
            Else
                CTAF4 = CLng(CODE)
                VAMHES = CKOL & "-" & CMOIN & "-" & CTAF & "-" & CTAF2 & "-" & CTAF3 & "-" & CTAF4
            End If
        End If
   End If
   If Not ISHESAB3(VAMHES) Then
       Err.Clear
       Call CREATHESCO(VAMHES, " وام " & CODE.Text)
       If Err <> 0 Then
          DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & GETTAFNAME(Forms![BASEKNOW]![PERSONEL] & "-1-" & CODE)
       Else
          DoCmd.OpenForm "MESAG", , , , , acDialog, "حساب وام ايجاد شد   " & VAMHES
       End If
    End If
  End If

End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like N'%" & Me.CODE.Text & "%'"
  Else
    Me.CODE = Me.CODE.Text
      Call CODE_AfterUpdate
  End If
  Response = acDataErrContinue

End Sub
Private Sub DATE_BP_Exit(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DATE_BP, False)
End Sub

Private Sub ESLAH_Click()
    Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.VAM_ID) Then
    dt = Now()
    Call TR("PVAM", "(VAM_ID = " & Me.VAM_ID & ")", dt, 1)
    Call TR("PVAM_BAZ", "(VAM_ID = " & Me.VAM_ID & ") AND (CODE = " & Me.CODE & ")", dt, 1)
  End If
    Me.AllowEdits = True
    Me.AllowDeletions = True
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
  Dim rst As New ADODB.Recordset
  If Not IsNull(Me.CODE) And IsNull(Me.VAM_ID) Then
    rst.Open "SELECT MAX(VAM_ID) AS MV  FROM PVAM WHERE CODE = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MV")) Then
      Me.VAM_ID = rst.Fields("MV") + 1
    Else
      Me.VAM_ID = 1
    End If
  End If

End Sub

Private Sub Form_Current()
                  On Error Resume Next
    If Me.NewRecord Or IsNull(Me.VAM_ID) Then
        Me.AllowEdits = True
        Me.AllowDeletions = True
    Else
      Me.AllowEdits = False
      Me.AllowDeletions = False
    End If
   Call SETSECURITY(Me.NAME, "VAM", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "VAM", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub


Private Sub MABLBZ_BeforeUpdate(CANCEL As Integer)
    If Me.MABLBZ > 0 Then
        If Me.mabl / Me.MABLBZ <> Round(Me.mabl / Me.MABLBZ) Then
          CANCEL = True
        Else
           Me.num = Me.mabl / Me.MABLBZ
        End If
    Else
      CANCEL = True
    End If

End Sub

Private Sub MABLBZ_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub NUM_BeforeUpdate(CANCEL As Integer)
    If Me.num > 0 Then
        If Me.mabl / Me.num <> Round(Me.mabl / Me.num) Then
          CANCEL = True
        Else
           Me.MABLBZ = Me.mabl / Me.num
        End If
    Else
      CANCEL = True
    End If
End Sub
