Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Dim chek As Boolean
Private Sub Form_AfterUpdate()
                         On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, Temp As Long
   Dim ST As String
    rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH_BACK]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields(0) > 0 Then
        Forms![HEAD_LST_FROOSH_BACK]![Command100].Enabled = True
    Else
        Forms![HEAD_LST_FROOSH_BACK]![Command100].Enabled = True
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
   Forms![HEAD_LST_FROOSH_BACK]![LETSANAD] = True
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "select *  from STUF_def where code = '" & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     If rst.Fields("radah") = 2 Or rst.Fields("radah") = 3 Then
       Temp = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![head_lst_froosh_back2]![DATE_N])
       If Me.AVRAGE2 <> Temp Then
           Me.AVRAGE2 = Temp
       End If
     Else
      If Me.MEGH > 0 And Me.mabl > 0 Then
         DoCmd.OpenForm "PARA", , , , , acHidden
         Forms![Para]![ID] = Me.ID
         Forms![Para]![COD] = Me.CODE
         Forms![Para]![ANBAR] = Me.ANBAR
         Forms![Para]![dt] = Forms![HEAD_LST_FROOSH_BACK]![DATE_N]
         DoCmd.OpenForm "AVRAGE", , , , , acHidden
         rst.Close
         Set rst = New ADODB.Recordset
         Set rst = Forms![AVRAGE].Form.RecordsetClone
         If rst.RecordCount > 0 Then
          If (rst.Fields(1) + Me.MEGHk) <> 0 Then
          '   Temp = Round(((rst.Fields(2) + Me.MABL_K) / (rst.Fields(1) + Me.MEGHk)) * 100)
             Me.AVRAGE2 = Temp / 100
          Else
             Me.AVRAGE2 = 0
          End If

        End If
        rst.Close
        DoCmd.Close acForm, "AVRAGE"
        DoCmd.Close acForm, "Para"
     End If
    End If
  End If
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
            Forms![HEAD_LST_FROOSH_BACK]![MOGU] = Null
        Else
            Forms![HEAD_LST_FROOSH_BACK]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
   Call SETSECURITYSUB("HEAD_LST_FROOSH_BACK", Me.NAME, "FACTBFR", 3)
End Sub


Private Sub MEGH_MAR_AfterUpdate()
                      On Error Resume Next
 Dim rst As New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH_BACK]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH_BACK]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Forms![HEAD_LST_FROOSH_BACK]![MOGU].Refresh
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
