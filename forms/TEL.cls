Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Text2_AfterUpdate()
    If IsNull(Me.Text2) Then
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE name Like N'%" & Text2 & "%' "
    End If
End Sub
Private Sub Text8_AfterUpdate()
    If IsNull(Me.Text8) Then
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE TEL Like N'%" & Text8 & "%' "
    End If
End Sub
