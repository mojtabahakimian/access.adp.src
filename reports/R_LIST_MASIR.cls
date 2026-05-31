Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
         On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.ADDRESS.FontName
Me.FontBold = Me.ADDRESS.FontBold
    If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
        Me.FontSize = 9
        If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
            Me.FontSize = 8
            If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
               Me.FontSize = 7
               If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
                    Me.ADDRESS.FontSize = 6
               Else
                     Me.ADDRESS.FontSize = 7
               End If
            Else
                 Me.ADDRESS.FontSize = 8
            End If
        Else
            Me.ADDRESS.FontSize = 9
        End If
    Else
        Me.ADDRESS.FontSize = 10
    End If
Me.FontSize = 10
Me.FontName = Me.NAME1.FontName
Me.FontBold = Me.NAME1.FontBold
    If Me.TextWidth(NAME1) / 567 > 6.042 Then
        Me.FontSize = 9
        If Me.TextWidth(NAME1) / 567 > 6.042 Then
            Me.FontSize = 8
            If Me.TextWidth(NAME1) / 567 > 6.042 Then
               Me.FontSize = 7
               If Me.TextWidth(NAME1) / 567 > 6.042 Then
                    Me.NAME1.FontSize = 6
               Else
                     Me.NAME1.FontSize = 7
               End If
            Else
                 Me.NAME1.FontSize = 8
            End If
        Else
            Me.NAME1.FontSize = 9
        End If
    Else
        Me.NAME1.FontSize = 10
    End If

End Sub

Private Sub Report_Open(CANCEL As Integer)
         On Error Resume Next
    Me.RecordSource = "SELECT     Visit_route.ROUTE_NAME, Visit_route.HES, Visit_route.IYALAT, Visit_route.CITY, Visit_route.District, Visit_route.CDATE, Visit_route.USERNAME, Visit_route.RACTIVE AS Visit_route_RACTIVE, Visit_route.OKF, Visit_route_dtl.COUST_NO, Visit_route_dtl.RACTIVE AS Visit_route_dtl_RACTIVE, Visit_route_dtl.IDR, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, CUST_HESAB.TEL, CUST_HESAB.MOBILE, CUST_HESAB_1.NAME AS VISITOR, Visit_route_dtl.CLASS FROM         Visit_route INNER JOIN  Visit_route_dtl ON Visit_route.ROUTE_NAME = Visit_route_dtl.ROUTE_NAME INNER JOIN  CUST_HESAB ON Visit_route_dtl.COUST_NO = CUST_HESAB.hes INNER JOIN    CUST_HESAB CUST_HESAB_1 ON Visit_route.HES = CUST_HESAB_1.hes WHERE     (Visit_route_dtl.RACTIVE = 1) AND (Visit_route.ROUTE_NAME = N'" & Forms![Visit_route_FORM]![ROUTE_NAME] & "'  or Visit_route.ROUTE_NAME = N'" & Fixp(Forms![Visit_route_FORM]![ROUTE_NAME]) & "' ) AND (Visit_route.HES = N'" & Forms![Visit_route_FORM]![HES] & "')"
End Sub
