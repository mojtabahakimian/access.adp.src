Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
Dim rst As New ADODB.Recordset
                 On Error Resume Next
If Not IsNull(Me.MOLAH) Then
 If RIGHT(Me.MOLAH, 1) = "." Then
     rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.ORDR_LST.MEGHk) AS Weight FROM         dbo.ORDR_LST INNER JOIN   dbo.STUF_DEF ON dbo.ORDR_LST.CODE = dbo.STUF_DEF.CODE WHERE      (dbo.ORDR_LST.ID = " & Forms![ORDR_HED]![ID] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) And rst.Fields("Weight") > 0 Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Format(Round(rst.Fields("Weight")), "#,##")
        End If
    End If
 End If
 End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
