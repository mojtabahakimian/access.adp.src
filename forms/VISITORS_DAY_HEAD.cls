Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command100_Click()
    DoCmd.OpenReport "R_LIST_VISIT_DAY", acPreview, , "HES ='" & Me.HES & "' AND  VDATE = " & Me.VDATE
End Sub

Private Sub Command23_Click()

                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "visitselect", acFormDS, , "hes ='" & Me.HES & "'"
End Sub

Private Sub Command23_LostFocus()
    Me.VISITORS_DAY_DTL_SUB.Requery
End Sub

Private Sub Form_AfterUpdate()
    Me.VISITORS_DAY_DTL_SUB.Form.COUST_NO.RowSource = "SELECT     TOP 100 PERCENT dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.hes AS Expr1 FROM  dbo.Visit_route INNER JOIN                    dbo.Visit_route_dtl ON dbo.Visit_route.ROUTE_NAME = dbo.Visit_route_dtl.ROUTE_NAME INNER JOIN  dbo.CUST_HESAB ON dbo.Visit_route_dtl.COUST_NO = dbo.CUST_HESAB.hes WHERE     (dbo.Visit_route.HES = N'" & Me.HES2 & "') ORDER BY dbo.CUST_HESAB.NAME"
    Me.VISITORS_DAY_DTL_SUB.Form.COUST_NO2.RowSource = "SELECT     TOP 100 PERCENT dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.hes AS Expr1 FROM  dbo.Visit_route INNER JOIN                    dbo.Visit_route_dtl ON dbo.Visit_route.ROUTE_NAME = dbo.Visit_route_dtl.ROUTE_NAME INNER JOIN  dbo.CUST_HESAB ON dbo.Visit_route_dtl.COUST_NO = dbo.CUST_HESAB.hes WHERE     (dbo.Visit_route.HES = N'" & Me.HES2 & "') ORDER BY dbo.CUST_HESAB.NAME"
End Sub

Private Sub Form_Current()
                 On Error Resume Next
If Me.NewRecord Then
     Me.VISITORS_DAY_DTL_SUB.Locked = True
  Else
     Me.VISITORS_DAY_DTL_SUB.Locked = False
 End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
