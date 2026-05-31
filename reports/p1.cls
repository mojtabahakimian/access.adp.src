Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 8
Me.FontName = Me.CUST_NO2.FontName
Me.FontBold = Me.CUST_NO2.FontBold
    If Me.TextWidth(CUST_NO2) / 567 > 9.323 Then
        Me.FontSize = 7
        If Me.TextWidth(CUST_NO2) / 567 > 9.323 Then
            Me.FontSize = 6
            If Me.TextWidth(CUST_NO2) / 567 > 9.323 Then
               Me.FontSize = 5
               If Me.TextWidth(CUST_NO2) / 567 > 9.323 Then
                    Me.CUST_NO2.FontSize = 4
               Else
                     Me.CUST_NO2.FontSize = 5
               End If
            Else
                 Me.CUST_NO2.FontSize = 6
            End If
        Else
            Me.CUST_NO2.FontSize = 7
        End If
    Else
        Me.CUST_NO2.FontSize = 8
    End If
End Sub

Private Sub Report_Close()
    DoCmd.GoToRecord acDataForm, "head_lst_pishfroosh2", acNewRec
    Forms![HEAD_LST_PISHFROOSH2]![CUST_NO].SetFocus
End Sub
Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     HES = '" & Forms![HEAD_LST_PISHFROOSH2]![CUST_NO] & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          Me.MANDAH.CAPTION = "  كل مانده حساب : 0  "
        Else
          Me.MANDAH.CAPTION = "  كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
        End If
    End If


End Sub
