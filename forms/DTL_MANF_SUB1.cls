Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Dim chek As Boolean

Private Sub ANBAR_AfterUpdate()
                       On Error Resume Next
   Me.CODE.Requery
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                        On Error Resume Next
    Response = acDataErrContinue
    Me.Requery
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
   Dim rst As New ADODB.Recordset
   If IsNull(Me.ANBAR) Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص است انبار و كالا نمي تواند داراي مقدار خالي باشد."
       CANCEL = True
   Else
       If IsNull(Me.CODE) Then
       Else
           rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount = 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
              CANCEL = True
           End If
       End If
   End If
End Sub

Private Sub Form_Current()
                       On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
 ' Call SETSECURITYSUB("HEAD_MANF", Me.NAME, "FORMOL", 3)
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
    Dim ST As String
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
    Else
     CANCEL = True
    End If

End Sub


Private Sub CODE_Enter()
                        On Error Resume Next
    Me.CODE.RowSource = "SELECT     TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM dbo.STUF_DEF INNER JOIN dbo.STUF_FSK INNER JOIN dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")ORDER BY dbo.STUF_DEF.NAME"
    Me.CODE.Requery
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
                        On Error Resume Next
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
End Sub
Private Sub CODE_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
    CODE = NewData
  End If
Response = acDataErrContinue
End Sub
Private Sub CODE_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.VAHED_K.Requery
   rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
   Else
       Me.VAHED_K = rst.Fields("VAHED")
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![HAZ_TOL] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & Me.CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
        rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
        rst.Fields("TNUMBER") = Me.CODE
        rst.Fields("NAME") = GETKALANAME(Me.CODE)
        rst.Fields("BED_BES") = -1
        rst.update
   Else
      If rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = Me.CODE Then
            rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
            rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
            rst.Fields("TNUMBER") = Me.CODE
            rst.Fields("NAME") = GETKALANAME(Me.CODE)
            rst.Fields("BED_BES") = -1
            rst.update
      End If
   End If

   'كنترل كالاي درجريان ساخت
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![CONKAL] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & Me.CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
        rst.Fields("TNUMBER") = Me.CODE
        rst.Fields("NAME") = GETKALANAME(Me.CODE)
        rst.Fields("BED_BES") = -1
        rst.update
   Else
        If rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = Me.CODE Then
            rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
            rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
            rst.Fields("TNUMBER") = Me.CODE
            rst.Fields("NAME") = GETKALANAME(Me.CODE)
            rst.Fields("BED_BES") = -1
            rst.update
        End If
   End If
   'عملكرد ماده
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![AMALKARD] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & Me.CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
        rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
        rst.Fields("TNUMBER") = Me.CODE
        rst.Fields("NAME") = "نرخ " & GETKALANAME(Me.CODE)
        rst.Fields("BED_BES") = -1
        rst.update
   Else
        If rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = Me.CODE Then
            rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
            rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
            rst.Fields("TNUMBER") = Me.CODE
            rst.Fields("NAME") = "نرخ " & GETKALANAME(Me.CODE)
            rst.Fields("BED_BES") = -1
            rst.update
        End If
   End If
  rst.Close
End Sub

Private Sub Form_Load()
                  On Error Resume Next
 Call SETSECURITYSUB("HEAD_MANF", Me.NAME, "FORMOL", 3)
End Sub

Private Sub MEGH_AfterUpdate()
                       On Error Resume Next
  Dim rst As New ADODB.Recordset, min As Double
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
  End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   End If
   rst.Close
   Me.MABLK = (Me.PERT + Me.MEGHk) * Me.smabl
End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub PERT_AfterUpdate()
                       On Error Resume Next
   Me.MABLK = (Me.PERT + Me.MEGHk) * Me.smabl
End Sub

Private Sub SMABL_AfterUpdate()
                       On Error Resume Next
   Me.MABLK = (Me.PERT + Me.MEGHk) * Me.smabl
End Sub

Private Sub SMABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub VAHED_K_AfterUpdate()
                       On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
   End If
   rst.Close
   Me.MABLK = (Me.PERT + Me.MEGHk) * Me.smabl
End Sub

Private Sub VAHED_K_Enter()
                     On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub VAHED_K_Exit(CANCEL As Integer)
                     On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
