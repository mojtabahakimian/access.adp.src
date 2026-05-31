Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Dim chek As Boolean
Private Sub Form_AfterUpdate()
                         On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
   Dim ST As String

    rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![head_lst_froosh_back2]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields(0) > 0 Then
        Forms![head_lst_froosh_back2]![Command100].Enabled = True
    Else
        Forms![head_lst_froosh_back2]![Command100].Enabled = True
    End If
    rst.Close
    Set rst = New ADODB.Recordset
   If Me.ANBAR > 0 Then
        If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
              MAND = rst.Fields("MAND")
              If Round(rst.Fields("MAND") + (Me.MEGH_MAR.TAG - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                 DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                 Me.MEGH_MAR = Me.MEGH_MAR.TAG
                 chek = True
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                rst.Fields("MOGODI") = MAND
                rst.Fields("MOGODI_A") = 0
                rst.update
              End If
           End If
        Else
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
        End If
    End If
   rst.Close
   Forms![head_lst_froosh_back2]![LETSANAD] = True
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
            Forms![head_lst_froosh_back2]![MOGU] = Null
        Else
            Forms![head_lst_froosh_back2]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
   Call SETSECURITYSUB("HEAD_LST_FROOSH_BACK2", Me.NAME, "FACTBFR", 3)
End Sub


Private Sub Form_Open(CANCEL As Integer)
 If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
    Me.JAY.ColumnHidden = False
    Me.JAYO.ColumnHidden = False
 Else
    Me.JAY.ColumnHidden = True
    Me.JAYO.ColumnHidden = True
 End If

End Sub

Private Sub MEGH_MAR_AfterUpdate()
                      On Error Resume Next
 Dim rst As New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_froosh_back2]![MOGU] = Null
    Else
        Forms![head_lst_froosh_back2]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Forms![head_lst_froosh_back2]![MOGU].Refresh
    rst.Close
End Sub

Private Sub MEGH_MAR_BeforeUpdate(CANCEL As Integer)
                        On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double
   If Me.ANBAR > 0 Then
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
  End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
