Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command16_Click()
    Text3 = ""
End Sub

Private Sub Frame10_AfterUpdate()
                On Error Resume Next
 If Me.Frame10 = 2 Then
   If IsNull(Me.Text3) Then
        Me.TELSUB.Form.RecordSource = "SELECT        NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM            dbo.CUST_HESAB"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE name Like N'%" & Text3 & "%' "
    End If
   Me.TELSUB.Form.AllowEdits = False
   Me.TELSUB.Form.AllowAdditions = False
  Else
    Me.TELSUB.Form.AllowEdits = True
    Me.TELSUB.Form.AllowAdditions = True
    If IsNull(Me.Text3) Then
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef WHERE name Like N'%" & Text3 & "%' "
    End If
End If
End Sub

Private Sub Text3_AfterUpdate()
                On Error Resume Next
 If Me.Frame10 = 2 Then
   If IsNull(Me.Text3) Then
        Me.TELSUB.Form.RecordSource = "SELECT        NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM            dbo.CUST_HESAB"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE name Like N'%" & Text3 & "%' "
    End If
   Me.TELSUB.Form.AllowEdits = False
   Me.TELSUB.Form.AllowAdditions = False
  Else
    Me.TELSUB.Form.AllowEdits = True
    Me.TELSUB.Form.AllowAdditions = True
    If IsNull(Me.Text3) Then
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef WHERE name Like N'%" & Text3 & "%' "
    End If
End If
End Sub
Private Sub Command15_Click()
    Text1 = ""
End Sub

Private Sub Text1_AfterUpdate()
                On Error Resume Next
 If Me.Frame10 = 2 Then
   If IsNull(Me.Text1) Then
        Me.TELSUB.Form.RecordSource = "SELECT        NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM            dbo.CUST_HESAB"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE name Like N'%" & Text1 & "%' "
    End If
   Me.TELSUB.Form.AllowEdits = False
   Me.TELSUB.Form.AllowAdditions = False
  Else
    Me.TELSUB.Form.AllowEdits = True
    Me.TELSUB.Form.AllowAdditions = True
    If IsNull(Me.Text3) Then
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef"
    Else
        Me.TELSUB.Form.RecordSource = "SELECT NAME, TEL FROM telef WHERE name Like N'%" & Text3 & "%' "
    End If
End If
End Sub
