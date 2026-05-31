Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
    Dim MAHLIST As String, XXD, XXM, XXS, mah, SAL As Long, SH As String, XXI As Long
    Dim rst As New ADODB.Recordset, i As Integer, TEDAD As Integer, rptname As String, RST2 As New ADODB.Recordset, NAGHDINEGI As Double
    Dim DT1 As Long, MH As Long, DT2 As Long, dt11 As Long, dt22 As Long, JAMBG As Double, NAGHD As Double
    BGIDD = IIf(IsNull(Me.OpenArgs), 1, Me.OpenArgs)
    DoCmd.OpenForm "BUN"
    Me.RecordSource = "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD
            Forms![BUN].Form.Refresh

    rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    XXD = UDay(rst.Fields("BGCDATE"))
    XXM = UMonth(rst.Fields("BGCDATE"))
    XXS = UYear(rst.Fields("BGCDATE"))
    mah = rst.Fields("BGMAH")
    SAL = rst.Fields("BGSAL")
    XXI = (SAL - XXS) * 12 + mah - XXM
    MH = XXM
    JAMBG = 0
        Forms![BUN].Form.Refresh
    For i = 0 To XXI
        Forms![BUN].Form.Refresh
        Me("M" & i + 1).CAPTION = MAHNAME(MH)
        Me("M" & i + 1).Visible = True
        If MH = 12 Then
            MH = 0
            SH = " " & SAL
        End If
        DT1 = XXS * 10000 + MH * 100 + XXD
        DT2 = XXS * 10000 + MH * 100 + 31
        dt11 = SAL * 10000 + mah * 100
        dt22 = SAL * 10000 + mah * 100 + 31
        MH = MH + 1
        GFRO = GETFROOSH(DT1, DT2)
        Me("BGFOCN" & i & "MONR").ControlSource = "=" & GFRO
        Me("BGFOCN" & i & "MONR").Visible = True
        If i = 0 Then
            Me("BGFOCN" & i & "MOND").ControlSource = "=" & Round(GFRO / IIf(rst.Fields("BGFOCXDAY") = 0, 1, rst.Fields("BGFOCXDAY")) * 100, 1)
        Else
            Me("BGFOCN" & i & "MOND").ControlSource = "=" & Round(GFRO / rst.Fields("BGFOCN" & i & "MON") * 100, 1)
        End If
        Me("BGFOCN" & i & "MOND").Visible = True
        Me("BGPCH" & i).ControlSource = "BGPCH" & XXI - i
        Me("BGPCH" & i).Visible = True
        Forms![BUN].Form.Refresh

        JAMBG = JAMBG + GETCHEKin(dt11, dt22, DT1, DT2)
        Me("BGMBCHEKR" & i).ControlSource = "=" & GETCHEKin(dt11, dt22, DT1, DT2)
        Me("BGMBCHEKR" & i).Visible = True
        Forms![BUN].Form.Refresh

        Me("BGMBCHEKJMR" & i).ControlSource = "=" & GETCHEKin(1, 99999999, DT1, DT2)
        Me("BGMBCHEKJMR" & i).Visible = True

        Forms![BUN].Form.Refresh
        Me("BGPCHD" & i).ControlSource = "=" & Round(GETCHEKin(dt11, dt22, DT1, DT2) / IIf(GFRO = 0, 1, GFRO) * 100, 1)
        Me("BGPCHD" & i).Visible = True
        If i = 0 Then
            Me("BGPCHM" & i).ControlSource = "=" & Round(rst.Fields("BGMBCHEKJM" & XXI - i) / IIf(rst.Fields("BGFOCXDAY") = 0, 1, rst.Fields("BGFOCXDAY")) * 100, 1)
        Else
            Me("BGPCHM" & i).ControlSource = "=" & Round(rst.Fields("BGMBCHEKJM" & XXI - i) / IIf(rst.Fields("BGFOCN" & i & "MON") = 0, 1, rst.Fields("BGFOCN" & i & "MON")) * 100, 1)
        End If
        Forms![BUN].Form.Refresh
        Me("BGPCHM" & i).Visible = True

        Forms![BUN].Form.Refresh
        Me("BGPCHR" & i).ControlSource = "=" & Round(GETCHEKin(1, 99999999, DT1, DT2) / IIf(GFRO = 0, 1, GFRO) * 100, 1)
        Me("BGPCHR" & i).Visible = True

        Forms![BUN].Form.Refresh
        XXD = 0
        Me("BGMBCHEK" & i).Visible = True
        Me("BGMBCHEKJM" & i).Visible = True
        Me("BGFOCN" & i + 1 & "MON").Visible = True
        Me("BGMBCHEKJM" & i).ControlSource = "BGMBCHEKJM" & XXI - i
        Me("BGMBCHEK" & i).ControlSource = "BGMBCHEK" & XXI - i
    Next i
    NAGHD = rst.Fields("BGBUGETMON")
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(DEED_DTL.BED) AS MAB FROM DEED_DTL INNER JOIN dbo.DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE     (DEED_HED.DATE_S BETWEEN " & dt11 & " AND " & dt22 & ")  AND (DEED_DTL.HES = N'110-9-9')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Forms![BUN].Form.Refresh
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MAB")) Then
        BGMBCASHD.ControlSource = "=" & rst.Fields("MAB")
        JAMBG = JAMBG + rst.Fields("MAB")
        BGPCASH0.ControlSource = "=" & Round(rst.Fields("MAB") / IIf(GETFROOSH(dt11, dt22) = 0, 1, GETFROOSH(dt11, dt22)) * 100, 1)
        Forms![BUN].Form.Refresh
    End If
     Me.DDD.ControlSource = "=" & JAMBG / IIf(NAGHD = 0, 1, NAGHD) * 100
   Me.BGBUGETMONR.ControlSource = "=" & JAMBG
        Forms![BUN].Form.Refresh
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.Close acForm, "bun"
End Sub
