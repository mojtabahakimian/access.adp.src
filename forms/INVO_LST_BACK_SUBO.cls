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
    rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH_BACKO]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields(0) > 0 Then
        Forms![HEAD_LST_FROOSH_BACKO]![Command100].Enabled = True
    Else
        Forms![HEAD_LST_FROOSH_BACKO]![Command100].Enabled = True
    End If
    rst.Close
    Set rst = New ADODB.Recordset
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
   If rst.RecordCount = 0 Then
   Else
      If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGH_MAR.TAG - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
         Me.MEGH_MAR = Me.MEGH_MAR.TAG
      Else

        rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGH_MAR.TAG - Me.MEGH_MAR)
        rst.update
     End If
   End If
   rst.Close
   Forms![HEAD_LST_FROOSH_BACKO]![LETSANAD] = True
End Sub


Private Sub Form_Current()
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
   Me.MEGH_MAR.TAG = Me.MEGH_MAR
   If Not IsNull(ANBAR) And Not IsNull(CODE) Then
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            Forms![HEAD_LST_FROOSH_BACKO]![MOGU] = Null
        Else
            Forms![HEAD_LST_FROOSH_BACKO]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
        End If
        rst.Close
        Set rst = New ADODB.Recordset
    End If
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHKG = Me.MEGH * rst.Fields(2)
   End If
   rst.Close
   Call SETSECURITYSUB("HEAD_LST_FROOSH_BACKO", Me.NAME, "FACTBFR", 3)
End Sub


Private Sub MEGH_MAR_AfterUpdate()
                      On Error Resume Next
 Dim rst As New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH_BACKO]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH_BACKO]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Forms![HEAD_LST_FROOSH_BACKO]![MOGU].Refresh
    rst.Close
End Sub

Private Sub MEGH_MAR_BeforeUpdate(CANCEL As Integer)
                        On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double
                     min = Getmin(Me.ANBAR, Me.CODE)
   Set rst = New ADODB.Recordset
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
      If Round(Me.MEGH_MAR - Me.MEGHKG, 5) > 0 Then
           DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "مقدار مرجوعي از مقدار فروش بيشتر است."
           CANCEL = True
      End If
      If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGH_MAR.TAG - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
           DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
           CANCEL = True
     End If
   End If
   rst.Close

End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
