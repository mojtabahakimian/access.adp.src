Attribute VB_Name = "modColorPicker"


Option Compare Database
Option Explicit
'***********  Code Start  ***********
Private Type COLORSTRUC
  lStructSize As Long
  hWnd As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Const CC_SOLIDCOLOR = &H80

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
  (pChoosecolor As COLORSTRUC) As Long

Public Function aDialogColor(ByVal hWnd As Long, Optional ByVal dtcol As Long) As Long
  Dim x As Long, CS As COLORSTRUC, CustColor(16) As Long

  CS.lStructSize = Len(CS)
  If hWnd <> 0 Then
    CS.hWnd = hWnd
  Else
    CS.hWnd = Application.hWndAccessApp
  End If
  CS.Flags = CC_SOLIDCOLOR
  CS.lpCustColors = String$(16 * 4, 0)
  If Not IsNull(dtcol) Then
    CS.rgbResult = dtcol
  End If
  x = ChooseColor(CS)
  If x = 0 Then
    ' ERROR - use Default White
    'prop = RGB(255, 255, 255) ' White
    aDialogColor = -1 'False
    Exit Function
  Else
    ' Normal processing
     aDialogColor = CS.rgbResult
  End If

End Function
'***********  Code End   ***********

Public Function GETFLDvlu(USERID As Integer, frm As String, fld As String, VLU As Variant) As Variant
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT * FROM USEROPTION WHERE USERID = " & USERID & " AND FLD ='" & fld & "' AND FRM = '" & frm & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        CurrentProject.Connection.Execute ("INSERT  INTO  USEROPTION (USERID,FLD,FRM,VLU) VALUES(" & USERID & ",'" & fld & "','" & frm & "'," & VLU & ")")
        GETFLDvlu = VLU
    Else
        GETFLDvlu = rst.Fields("VLU")
    End If
End Function

Public Sub SETFLDvlu(USERID As Integer, frm As String, fld As String, VLU As Variant)
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT * FROM USEROPTION WHERE USERID = " & USERID & " AND FLD ='" & fld & "' AND FRM = '" & frm & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        CurrentProject.Connection.Execute ("INSERT  INTO  USEROPTION (USERID,FLD,FRM,VLU) VALUES(" & USERID & ",'" & fld & "','" & frm & "'," & VLU & ")")
    Else
        CurrentProject.Connection.Execute ("UPDATE  USEROPTION SET VLU = " & VLU & " WHERE (USERID = " & USERID & " AND FLD = '" & fld & "' AND FRM = '" & frm & "')")
    End If
End Sub
