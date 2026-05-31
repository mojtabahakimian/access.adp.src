Attribute VB_Name = "SIGNMUD"
Option Compare Database



Public Function LETSIGNER(ID As Integer, LEV As Integer, USERCO As Integer) As Boolean
    Dim rst As New ADODB.Recordset
    Select Case ID
        Case 20:
            rst.Open "SELECT   USERCO, FFRP_FROOSH FROM dbo.SIGN WHERE    FFRP_FROOSH = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                LETSIGNER = True
            Else
                LETSIGNER = False
            End If
        Case 1:
            Select Case LEV
                Case 1:
                    rst.Open "SELECT   USERCO, ANB_RASID_ANB FROM dbo.SIGN WHERE    KFR_BAZAR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 2:
                    rst.Open "SELECT   USERCO, ANB_RASID_QC   FROM dbo.SIGN WHERE    KFR_HESAB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 3:
                    rst.Open "SELECT   USERCO, ANB_RASID_QC   FROM dbo.SIGN WHERE    KFR_MODIR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
            End Select
        Case 6:
            Select Case LEV
                Case 1:
                    rst.Open "SELECT   USERCO, ENTGH_AZANB FROM dbo.SIGN WHERE    ENTGH_AZANB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 2:
                    rst.Open "SELECT   USERCO, ENTGH_BEANB   FROM dbo.SIGN WHERE    ENTGH_BEANB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 3:
                    rst.Open "SELECT   USERCO, ENTGH_MODIR   FROM dbo.SIGN WHERE    ENTGH_MODIR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
            End Select
        Case 2:
            rst.Open "SELECT   USERCO, ANB_HAVL_FROSH FROM dbo.SIGN WHERE    ANB_HAVL_FROSH = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                LETSIGNER = True
            Else
                LETSIGNER = False
            End If
        Case 100:
            Select Case LEV
                Case 1:
                    rst.Open "SELECT   USERCO, PAY_MVAHED FROM dbo.SIGN WHERE    PAY_MVAHED = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 2:
                    rst.Open "SELECT   USERCO, PAY_HESAB   FROM dbo.SIGN WHERE    PAY_HESAB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 3:
                    rst.Open "SELECT   USERCO, PAY_MAMEL   FROM dbo.SIGN WHERE    PAY_MAMEL = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
            End Select
        Case 13:
            Select Case LEV
                Case 1:
                    rst.Open "SELECT   USERCO, FFR_FROOSH FROM dbo.SIGN WHERE    FFR_FROOSH = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 2:
                    rst.Open "SELECT   USERCO, FFR_HESAB   FROM dbo.SIGN WHERE    FFR_HESAB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 3:
                    rst.Open "SELECT   USERCO, FFR_MODIR   FROM dbo.SIGN WHERE    FFR_MODIR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
            End Select
        Case 12:
            Select Case LEV
                Case 1:
                    rst.Open "SELECT   USERCO, KFR_BAZAR FROM dbo.SIGN WHERE    KFR_BAZAR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 2:
                    rst.Open "SELECT   USERCO, KFR_HESAB   FROM dbo.SIGN WHERE    KFR_HESAB = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
                Case 3:
                    rst.Open "SELECT   USERCO, KFR_MODIR   FROM dbo.SIGN WHERE    KFR_MODIR = 1  AND USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        LETSIGNER = True
                    Else
                        LETSIGNER = False
                    End If
            End Select
    End Select
End Function


Public Function GETSIGNER(ID As Integer, USERCO As Integer) As Integer
    Dim rst As New ADODB.Recordset
    Select Case ID
        Case 20:
                rst.Open "SELECT   USERCO, FFRP_FROOSH, FFRP_ANB, FFRP_HESAB,FFRP_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("FFRP_MODIR") = True Then
                        GETSIGNER = 4
                    Else
                        If rst.Fields("FFRP_HESAB") = True Then
                             GETSIGNER = 3
                        Else
                            If rst.Fields("FFRP_ANB") = True Then
                                 GETSIGNER = 2
                            Else
                                If rst.Fields("FFRP_FROOSH") = True Then
                                     GETSIGNER = 1
                                Else
                                    GETSIGNER = 0
                                End If
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
        Case 1:
                rst.Open "SELECT   USERCO, ANB_RASID_ANB, ANB_RASID_QC FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("ANB_RASID_QC") = True Then
                        GETSIGNER = 2
                    Else
                        If rst.Fields("ANB_RASID_ANB") = True Then
                             GETSIGNER = 1
                        Else
                            If rst.Fields("ANB_RASID_ANB") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
        Case 6:
                rst.Open "SELECT   USERCO, ENTGH_AZANB, ENTGH_BEANB, ENTGH_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("ENTGH_MODIR") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("ENTGH_BEANB") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("ENTGH_AZANB") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
        Case 13:
                rst.Open "SELECT   USERCO, FFR_FROOSH, FFR_HESAB, FFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("FFR_MODIR") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("FFR_HESAB") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("FFR_FROOSH") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
        Case 12:
                rst.Open "SELECT   USERCO, KFR_BAZAR, KFR_HESAB, KFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("KFR_MODIR") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("KFR_HESAB") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("KFR_BAZAR") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
        Case 2:
                rst.Open "SELECT   USERCO, ANB_HAVL_FROSH, ANB_HAVL_ANB, ANB_HAVL_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("ANB_HAVL_MODIR") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("ANB_HAVL_ANB") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("ANB_HAVL_FROSH") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
         Case 100:
                rst.Open "SELECT   USERCO, PAY_MVAHED, PAY_HESAB, PAY_MAMEL FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("PAY_MAMEL") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("PAY_HESAB") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("PAY_MVAHED") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
         Case 90:
                rst.Open "SELECT   USERCO, PAZ_CEO, PAZ_HES, PAZ_PAZ FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields("PAZ_CEO") = True Then
                        GETSIGNER = 3
                    Else
                        If rst.Fields("PAZ_HES") = True Then
                             GETSIGNER = 2
                        Else
                            If rst.Fields("PAZ_PAZ") = True Then
                                 GETSIGNER = 1
                            Else
                                 GETSIGNER = 0
                            End If
                         End If
                    End If
                Else
                    GETSIGNER = 0
                End If
    End Select
End Function

Public Sub LetSigneTick(frm As Form, ID As Integer, USERCO As Integer)
    Dim rst As New ADODB.Recordset
    Select Case ID
        Case 20:
                rst.Open "SELECT   USERCO, FFRP_FROOSH, FFRP_ANB, FFRP_HESAB,FFRP_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 1:
                rst.Open "SELECT   USERCO, ANB_RASID_ANB, ANB_RASID_QC,0 AS F3 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 6:
                rst.Open "SELECT   USERCO, ENTGH_AZANB, ENTGH_BEANB, ENTGH_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 4:
                rst.Open "SELECT   USERCO, FFRB_FROOSH, FFRB_ANB, FFRB_HESAB, FFRB_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               ' If rst.RecordCount > 0 Then
                '    If rst.Fields(4) = True Then
                 '       frm![SGN4].Enabled = True
                  '      frm![SGN4].Locked = False
                   ' Else
                    '    frm![SGN4].Enabled = False
                     '   frm![SGN4].Locked = True
               '     End If
               ' End If

        Case 3:
                rst.Open "SELECT   USERCO, KFRB_BAZAR, KFRB_ANB, KFRB_HESAB, KFRB_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields(4) = True Then
                        frm![SGN4].Enabled = True
                        frm![SGN4].Locked = False
                    Else
                        frm![SGN4].Enabled = False
                        frm![SGN4].Locked = True
                    End If
                End If

        Case 27:
                rst.Open "SELECT   USERCO, KFRB_BAZAR, KFRB_ANB, KFRB_HESAB, KFRB_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If rst.Fields(4) = True Then
                        frm![SGN4].Enabled = True
                        frm![SGN4].Locked = False
                    Else
                        frm![SGN4].Enabled = False
                        frm![SGN4].Locked = True
                    End If
                End If

        Case 13:
                rst.Open "SELECT   USERCO, FFR_FROOSH, FFR_HESAB,FFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 33: ' صادراتي '
                rst.Open "SELECT   USERCO, SGN0133 , SGN0233,SGN0333 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 32: ' صادراتي '
                rst.Open "SELECT   USERCO, SGN0132 , SGN0232,SGN0332 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 34: ' خزانه '
                rst.Open "SELECT   USERCO, SGN0134 , SGN0234,SGN0334 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 35: ' سفارشات '
                rst.Open "SELECT   USERCO, SGN0135 , SGN0235,SGN0335 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 36: ' درخواست خريد '
                rst.Open "SELECT   USERCO, SGN0136 , SGN0236,SGN0336 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 37: ' درخواست مرخصي '
                rst.Open "SELECT   USERCO, SGN0137 , SGN0237,SGN0337 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 38: ' خروج ساير '
                rst.Open "SELECT   USERCO, ANB_KHOROGS_MODIRT , ANB_KHOROGS_ANB, ANB_KHOROGS_ANB AS HES FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 39: ' حواله خروج  '
                rst.Open "SELECT   USERCO, ANB_KHOROG_MODIRT , ANB_KHOROG_ANB, ANB_KHOROG_ANB AS HES FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 40: 'ASSETS  '
                rst.Open "SELECT   USERCO, SGN0140 , SGN0240,SGN0340 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 41: 'درخواست کالا از انبار  '
                rst.Open "SELECT   USERCO, SGN0141 , SGN0241,SGN0341 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 42: 'اعلام خرابي  '
                rst.Open "SELECT   USERCO, SGN0142 , SGN0242,SGN0342 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 12:
                rst.Open "SELECT   USERCO, KFR_BAZAR, KFR_HESAB, KFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Case 2:
                rst.Open "SELECT   USERCO, ANB_HAVL_FROSH, ANB_HAVL_ANB, ANB_HAVL_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Case 100:
                rst.Open "SELECT   USERCO, PAY_MVAHED, PAY_HESAB, PAY_MAMEL FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Case 90:
                rst.Open "SELECT   USERCO, PAZ_PAZ , PAZ_HES,PAZ_CEO FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Case 24:
                rst.Open "SELECT   USERCO, SGN0124 , SGN0224,SGN0324 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Case 26:
                rst.Open "SELECT   USERCO, SGN0126 , SGN0226,SGN0326 FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         Case 0:
                rst.Open "SELECT   USERCO, SND_TAHI,SND_MALI,SND_MODIR FROM dbo.SIGN WHERE     USERCO = " & USERCO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End Select
    If rst.RecordCount > 0 Then
        If rst.Fields(3) = True Then
            frm![SGN3].Enabled = True
            frm![SGN3].Locked = False
        Else
            frm![SGN3].Enabled = False
            frm![SGN3].Locked = True
        End If
        If rst.Fields(2) = True Then
            frm![SGN2].Enabled = True
            frm![SGN2].Locked = False
        Else
            frm![SGN2].Enabled = False
            frm![SGN2].Locked = True
        End If
        If rst.Fields(1) = True Then
            frm![SGN1].Enabled = True
            frm![SGN1].Locked = False
        Else
            frm![SGN1].Enabled = False
            frm![SGN1].Locked = True
        End If
    Else
        frm![SGN3].Enabled = False
        frm![SGN3].Locked = True
        frm![SGN2].Enabled = False
        frm![SGN2].Locked = True
        frm![SGN1].Enabled = False
        frm![SGN1].Locked = True
    End If

End Sub
