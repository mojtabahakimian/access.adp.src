Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
Dim chek As Boolean
Private Sub Form_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double
   Dim ST As String
   rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_KH_BACK_azad]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.Fields(0) > 0 Then
       Forms![HEAD_LST_KH_BACK_azad]![Command100].Enabled = True
   Else
       Forms![HEAD_LST_KH_BACK_azad]![Command100].Enabled = True
   End If
   Forms![HEAD_LST_KH_BACK_azad]![LETSANAD] = True
End Sub

Private Sub Form_Current()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   If Not IsNull(ANBAR) And Not IsNull(CODE) Then
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            Forms![HEAD_LST_KH_BACK_azad]![MOGU] = Null
        Else
            Forms![HEAD_LST_KH_BACK_azad]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
        End If
        Set rst = New ADODB.Recordset
   End If
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
        Me.MEGHKG = Me.MEGH * rst.Fields(2)
   End If
   rst.Close
   Call SETSECURITYSUB("HEAD_LST_KH_BACK_azad", Me.NAME, "FACTBKH", 3)
End Sub

Private Sub MABL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, Temp As Long
  If mabl = 0 Then
    MABL_K.TabStop = True
    MABL_K = Round(Me.mabl * Me.MEGHk)
  Else
    MABL_K.TabStop = False
    MABL_K = Round(Me.mabl * Me.MEGHk)
  End If
  Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_KH_BACK_azad]![DATE_N])
End Sub

Private Sub MABL_K_AfterUpdate()
     Dim rst As New ADODB.Recordset
                      On Error Resume Next
     If MEGHk = 0 Then
        MABL_K = 0
     Else
        mabl = MABL_K / MEGHk
        MABL_K.TabStop = False
     End If
End Sub

Private Sub MABL_K_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub MEGH_MAR_AfterUpdate()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_KH_BACK_azad]![MOGU] = Null
    Else
        Forms![HEAD_LST_KH_BACK_azad]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Forms![HEAD_LST_KH_BACK_azad]![MOGU].Refresh
    rst.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
