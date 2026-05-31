Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_DblClick(CANCEL As Integer)
             On Error Resume Next
    Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = aDialogColor(Me.hWnd)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Call SETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
               On Error Resume Next
    If KeyAscii = 27 Then
       DoCmd.Close
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
               On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
           On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, sudm As Double, khareed As Double
    khareed = 0
    If IsLoaded("HEAD_LST_FROOSH22") Or IsLoaded("HEAD_LST_FROOSH1") Or IsLoaded("HEAD_LST_FROOSH2") Then
        If Forms![BASEKNOW]![TKHF] = 3 And IsLoaded("HEAD_LST_FROOSH22") Then
            rst.Open "select code ,anbar ,MABL, mabl_k, MEGHk  from invo_lst where tag = 2 and number = " & Forms!HEAD_LST_FROOSH22!NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst.EOF
               RST2.Open "SELECT INVO_LST.CODE, INVO_LST.MABL FROM HEAD_LST INNER JOIN INVO_LST ON (HEAD_LST.TAG = INVO_LST.TAG) AND (HEAD_LST.NUMBER = INVO_LST.NUMBER) WHERE (((INVO_LST.CODE) = '" & rst.Fields("code") & "') And ((HEAD_LST.TAG) = 1) And ((HEAD_LST.DATE_N) <= " & [Forms]![HEAD_LST_FROOSH22]![DATE_N] & ")) ORDER BY HEAD_LST.DATE_N DESC,HEAD_LST.NUMBER DESC,INVO_LST.RADIF DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If RST2.RecordCount > 0 Then
                  sudm = sudm + rst.Fields("mabl_k") - (RST2.Fields("mabl") * rst.Fields("MEGHk"))
                  khareed = RST2.Fields("mabl")
               Else
                   khareed = GetDayPrice(rst.Fields("code"), Forms!HEAD_LST_FROOSH22!DATE_N)
                   If khareed <> 0 Then
                       sudm = sudm + rst.Fields("mabl_k") - (khareed * rst.Fields("MEGHk"))
                   Else
                        Set RST2 = New ADODB.Recordset
                        RST2.Open "SELECT STUF_FSK.FI_A AS MABL FROM STUF_FSK WHERE (((STUF_FSK.CODE)='" & rst.Fields("code") & "') AND ((STUF_FSK.ANBAR)=" & rst.Fields("anbar") & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                            sudm = sudm + rst.Fields("mabl_k") - (RST2.Fields("MABL") * rst.Fields("MEGHk"))
                            khareed = RST2.Fields("mabl")
                        End If
                   End If
               End If
               Me.List7.AddItem Chr(34) & (rst.Fields("CODE") & Chr(34) & ";" & Chr(34) & GETKALANAME(rst.Fields("CODE")) & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL"), "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL_k"), "#,###") & Chr(34) & ";" & Chr(34) & Format(khareed, "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MEGHK") * khareed, "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL_k") - (khareed * rst.Fields("MEGHk")), "#,###") & Chr(34) & ";")
               rst.MoveNext
               Set RST2 = New ADODB.Recordset
            Wend
            Text0 = sudm - Forms!HEAD_LST_FROOSH22!NTKHFIF + Forms!HEAD_LST_FROOSH22!HKH
            DAR = Forms!HEAD_LST_FROOSH22!HKH
            KASR = Forms!HEAD_LST_FROOSH22!NTKHFIF
            SUD = sudm
        Else
            rst.Open "select code ,anbar ,MABL, mabl_k, MEGHk  from invo_lst where tag = 2 and number = " & Forms!HEAD_LST_FROOSH2!NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst.EOF
               RST2.Open "SELECT INVO_LST.CODE, INVO_LST.MABL FROM HEAD_LST INNER JOIN INVO_LST ON (HEAD_LST.TAG = INVO_LST.TAG) AND (HEAD_LST.NUMBER = INVO_LST.NUMBER) WHERE (((INVO_LST.CODE) = '" & rst.Fields("code") & "') And ((HEAD_LST.TAG) = 1) And ((HEAD_LST.DATE_N) <= " & [Forms]![HEAD_LST_FROOSH2]![DATE_N] & ")) ORDER BY HEAD_LST.DATE_N DESC,HEAD_LST.NUMBER DESC,INVO_LST.RADIF DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If RST2.RecordCount > 0 Then
                  sudm = sudm + rst.Fields("mabl_k") - (RST2.Fields("mabl") * rst.Fields("MEGHk"))
                  khareed = RST2.Fields("mabl")
               Else
                  ' khareed = GetDayPrice(RST.Fields("code"), Forms!HEAD_LST_FROOSH2!DATE_N)
                   khareed = LASTAVRAGE(rst.Fields("code"), rst.Fields("anbar"), Forms!HEAD_LST_FROOSH2!DATE_N)
                   If khareed <> 0 Then
                       sudm = sudm + rst.Fields("mabl_k") - (khareed * rst.Fields("MEGHk"))
                   Else
                        Set RST2 = New ADODB.Recordset
                        RST2.Open "SELECT STUF_FSK.FI_A AS MABL FROM STUF_FSK WHERE (((STUF_FSK.CODE)='" & rst.Fields("code") & "') AND ((STUF_FSK.ANBAR)=" & rst.Fields("anbar") & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                            sudm = sudm + rst.Fields("mabl_k") - (RST2.Fields("MABL") * rst.Fields("MEGHk"))
                            khareed = RST2.Fields("mabl")
                        End If
                   End If
               End If
             ' Debug.Print (RST.Fields("CODE") & ";" & GETKALANAME(RST.Fields("CODE")) & ";" & Format(RST.Fields("MABL"), "#,###") & ";" & Format(RST.Fields("MABL_k"), "#,###") & ";" & Format(khareed, "#,###") & ";" & Format(RST.Fields("MEGHK") * khareed, "#,###") & ";" & Format(RST.Fields("MABL_k") - (khareed * RST.Fields("MEGHk")), "#,###") & ";")
               Me.List7.AddItem Chr(34) & (rst.Fields("CODE") & Chr(34) & ";" & Chr(34) & GETKALANAME(rst.Fields("CODE")) & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL"), "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL_k"), "#,###") & Chr(34) & ";" & Chr(34) & Format(khareed, "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MEGHK") * khareed, "#,###") & Chr(34) & ";" & Chr(34) & Format(rst.Fields("MABL_k") - (khareed * rst.Fields("MEGHk")), "#,###") & Chr(34) & ";")
               rst.MoveNext
               Set RST2 = New ADODB.Recordset
            Wend
            Text0 = sudm - Forms!HEAD_LST_FROOSH2!NTKHFIF + Forms!HEAD_LST_FROOSH2!HKH
            DAR = Forms!HEAD_LST_FROOSH2!HKH
            KASR = Forms!HEAD_LST_FROOSH2!NTKHFIF
            SUD = sudm
        End If
   End If
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = Me.Detail.BackColor
     BKC = GETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Call UISetRoundRect(Me, 20, False)
End Sub
