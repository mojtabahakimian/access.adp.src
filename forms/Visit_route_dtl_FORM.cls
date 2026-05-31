Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COUST_NO_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, TBH As String
    TBH = ""
    TBH = GetHesLevel(Me.COUST_NO)
     Select Case TBH
        Case "TDETA_HES":
                    rst.Open "select   RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES2":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES3":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES4":
                    rst.Open "select   RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR))  = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    End Select
    If rst.RecordCount > 0 Then
        'اگرمشتري به هيچ مسيري مرتبط نيست
        If IsNull(rst.Fields("ROUTE_NAME")) Then
            rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
            rst.update
       Else
        'اگرمشتري به  مسيري مرتبط است
            If rst.Fields("ROUTE_NAME") <> Me.ROUTE_NAME Then
                RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.COUST_NO & "' and RACTIVE = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & rst.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
                    If Forms![BASEKNOW]![Text44] Then
                        RST2.Fields("ractive") = False
                        RST2.update
                        rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                        rst.update
                    Else
                      CANCEL = True
                      Me.Undo
                    End If
                Else
                    rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                    rst.update
                End If
            End If
        End If
    End If

End Sub

Private Sub COUST_NO_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If COUST_NO.Text = "-" Or COUST_NO.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "43"
  Response = acDataErrContinue
End If

End Sub

Private Sub COUST_NO2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, TBH As String
    TBH = ""
    TBH = GetHesLevel(Me.COUST_NO)
     Select Case TBH
        Case "TDETA_HES":
                    rst.Open "select   RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES2":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES3":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES4":
                    rst.Open "select   RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR))  = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    End Select
    If rst.RecordCount > 0 Then
        'اگرمشتري به هيچ مسيري مرتبط نيست
        If IsNull(rst.Fields("ROUTE_NAME")) Then
            rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
            rst.update
       Else
        'اگرمشتري به  مسيري مرتبط است
            If rst.Fields("ROUTE_NAME") <> Me.ROUTE_NAME Then
                RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.COUST_NO & "' and RACTIVE = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & rst.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
                    If Forms![BASEKNOW]![Text44] Then
                        RST2.Fields("ractive") = False
                        RST2.update
                        rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                        rst.update
                    Else
                      CANCEL = True
                      Me.Undo
                    End If
                Else
                    rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                    rst.update
                End If
            End If
        End If
    End If

End Sub

Private Sub COUST_NO2_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If COUST_NO.Text = "-" Or COUST_NO.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "34"
  Response = acDataErrContinue

End If

End Sub



Private Sub Form_AfterUpdate()
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, TBH As String
    TBH = ""
    TBH = GetHesLevel(Me.COUST_NO)
     Select Case TBH
        Case "TDETA_HES":
                    rst.Open "select   RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES2":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES3":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES4":
                    rst.Open "select   RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR))  = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    End Select
    If rst.RecordCount > 0 Then
        'اگرمشتري به هيچ مسيري مرتبط نيست
        If IsNull(rst.Fields("ROUTE_NAME")) Then
            rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
            rst.update
       Else
        'اگرمشتري به  مسيري مرتبط است
            If rst.Fields("ROUTE_NAME") <> Me.ROUTE_NAME Then
                RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.COUST_NO & "' and RACTIVE = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & rst.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
                    If Forms![BASEKNOW]![Text44] Then
                        RST2.Fields("ractive") = False
                        RST2.update
                        rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                        rst.update
                    Else
                      CANCEL = True
                      Me.Undo
                    End If
                Else
                    rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                    rst.update
                End If
            End If
        End If
    End If


End Sub

Private Sub RACTIVE_DblClick(CANCEL As Integer)
                    On Error Resume Next
   RACTIVE = Not RACTIVE
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, TBH As String
    TBH = ""
    TBH = GetHesLevel(Me.COUST_NO)
     Select Case TBH
        Case "TDETA_HES":
                    rst.Open "select   RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES2":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES3":
                    rst.Open "select    RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR))  AS Expr1, ROUTE_NAME  from " & TBH & "  where  RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))    + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        Case "TDETA_HES4":
                    rst.Open "select   RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS Expr1, ROUTE_NAME  from " & TBH & "  where RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))   + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR))  = '" & Me.COUST_NO & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    End Select
    If rst.RecordCount > 0 Then
        'اگرمشتري به هيچ مسيري مرتبط نيست
        If IsNull(rst.Fields("ROUTE_NAME")) Then
            rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
            rst.update
       Else
        'اگرمشتري به  مسيري مرتبط است
            If rst.Fields("ROUTE_NAME") <> Me.ROUTE_NAME Then
                RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.COUST_NO & "' and RACTIVE = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & rst.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
                    If Forms![BASEKNOW]![Text44] Then
                        RST2.Fields("ractive") = False
                        RST2.update
                        rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                        rst.update
                    Else
                      CANCEL = True
                      Me.Undo
                    End If
                Else
                    rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                    rst.update
                End If
            End If
        End If
    End If

End Sub
