Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_DataSetChange()
DoCmd.Beep
End Sub

Private Sub t1_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
    Set rst = Me.RecordsetClone
    Dim fld As Field
    Dim i As Integer
    For Each fld In rst.Fields
       i = 1 + i
       Me("T" & i).ControlSource = fld.NAME
    Next
    DoCmd.RunCommand acCmdProperties
End Sub
