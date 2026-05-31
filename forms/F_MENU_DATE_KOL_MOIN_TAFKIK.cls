Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Command5_Click()
                     On Error Resume Next
  Dim sql, COND  As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
    If IsNull(Me.DT1) Then
       DT1 = Forms![BASEKNOW]![YEA] & "0101"
     End If

  If IsNull(Me.HKOL) Or IsNull(Me.HMOIN1) Or IsNull(Me.TAFZ1) Or IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If Me.HHES = 1 Then
        COND = "(BEDBES <> 0 )"
  Else
    If Me.HHES = 2 Then
        COND = "(BEDBES = 0 )"
    Else
       COND = ""
    End If
  End If
  If (Me.HHES = 1 Or Me.HHES = 3) Then
    If BEDBES = 1 Then
        If Len(COND) > 0 Then
            COND = COND & " AND (BEDBES > 0)"
        Else
            COND = "(BEDBES > 0)"
        End If
    Else
        If BEDBES = 2 Then
            If Len(COND) > 0 Then
               COND = COND & " AND  (BEDBES < 0)"
            Else
               COND = "(BEDBES < 0)"
            End If
        End If
    End If
  End If
  DoCmd.OpenReport "R_DAFTAR_MOIN_TFKIK", acViewPreview, , COND
  DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Command6_Click()
                      On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub HKOL_AfterUpdate()
                      On Error Resume Next
    Me.HMOIN1.RowSource = "SELECT  NUMBER, NAME, NUMBER AS Expr1 FROM DETA_HES WHERE (N_KOL = " & Me.HKOL & ") ORDER BY NAME"
    If [Forms]![BASEKNOW]![PERSON] = 2 Then
        Me.TAFZ1.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & "AND (CODE_E = N'0') OR (CODE_E IS NULL)) ORDER BY TDETA_HES.TNUMBER, TDETA_HES.NAME "
        Me.TAFZ2.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & "AND (CODE_E = N'0') OR (CODE_E IS NULL)) ORDER BY TDETA_HES.TNUMBER DESC, TDETA_HES.NAME "
    Else
        Me.TAFZ1.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & " ) ORDER BY TDETA_HES.TNUMBER, TDETA_HES.NAME "
        Me.TAFZ2.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & " ) ORDER BY TDETA_HES.TNUMBER DESC, TDETA_HES.NAME "
    End If
    Me.HMOIN1.Requery
    Me.TAFZ1.Requery
    Me.TAFZ2.Requery
End Sub

Private Sub HKOL_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  Me.HKOL = NewData
  Call HKOL_AfterUpdate
  Response = acDataErrContinue
End Sub

Private Sub HMOIN1_AfterUpdate()
                     On Error Resume Next
    If [Forms]![BASEKNOW]![PERSON] = 2 Then
        Me.TAFZ1.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & "AND (CODE_E = N'0') OR (CODE_E IS NULL))) ORDER BY TDETA_HES.TNUMBER, TDETA_HES.NAME "
        Me.TAFZ2.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & "AND (CODE_E = N'0') OR (CODE_E IS NULL))) ORDER BY TDETA_HES.TNUMBER DESC, TDETA_HES.NAME "
    Else
        Me.TAFZ1.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & " )) ORDER BY TDETA_HES.TNUMBER, TDETA_HES.NAME "
        Me.TAFZ2.RowSource = "SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TNUMBER FROM TDETA_HES WHERE     (((TDETA_HES.NUMBER) = " & Me.HMOIN1 & " ) AND ((TDETA_HES.N_KOL)= " & Me.HKOL & " )) ORDER BY TDETA_HES.TNUMBER DESC, TDETA_HES.NAME "
    End If
   Me.TAFZ1.Requery
   Me.TAFZ2.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub HMOIN1_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  Me.HMOIN1 = NewData
  Call HMOIN1_AfterUpdate
  Response = acDataErrContinue
End Sub

Private Sub TAFZ1_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  Me.TAFZ1 = NewData
  Response = acDataErrContinue
End Sub

Private Sub TAFZ2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  Me.TAFZ2 = NewData
  Response = acDataErrContinue
End Sub
