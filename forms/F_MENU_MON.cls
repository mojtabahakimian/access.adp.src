Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  Select Case Me.OpenArgs
    Case 1:
       CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'CT1')   DROP VIEW  CT1")
       CurrentProject.Connection.Execute ("CREATE VIEW  CT1 as  SELECT      dbo.STUF_DEF.NAME,SUM(dbo.INVO_LST.MEGHk) AS MEGHT FROM   dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN   dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN    dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE WHERE (dbo.HEAD_LST.TAG = 9) And (dbo.Umonth(dbo.HEAD_LST.DATE_N) = " & Me.MON & " )  And (dbo.TCOD_STUFGROUP.CODE = " & Me.GRP & " ) AND (dbo.TCOD_VAHEDS.CODE = " & Me.VAH & ") GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_STUFGROUP.NAMES, dbo.TCOD_VAHEDS.NAMES")
       DoCmd.OpenReport "CHART_TOLID_MON", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
   Case 2:
        DoCmd.OpenForm "AMAR_KOL_TOLID_KALA", acFormPivotChart, , "GRP = " & Me.GRP & " AND VH = " & Me.VAH
        DoCmd.OpenForm Me.NAME, , , , , acHidden
   Case 3:
        DoCmd.OpenForm "AMAR_KOL_FROOSH", acFormPivotChart, , "GRP = " & Me.GRP & " AND VH = " & Me.VAH
        DoCmd.OpenForm Me.NAME, , , , , acHidden
 End Select
End Sub

Private Sub Command6_Click()
                    On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                 On Error Resume Next
 Select Case Me.OpenArgs
    Case 2:
           Me.MON.Visible = False
           Me.GRP.Left = 2850
           Me.Label33.Left = 4900
    Case 3:
           Me.MON.Visible = False
           Me.GRP.Left = 2850
           Me.Label33.Left = 4900
    Case "pchs": Me.HelpContextId = 800601

    Case "dchs": Me.HelpContextId = 800501

    Case "pchss": Me.HelpContextId = 800602

    Case "dchss": Me.HelpContextId = 800502

    Case "chkm": Me.HelpContextId = 8009

    Case "chkb": Me.HelpContextId = 8008

    Case "chkv": Me.HelpContextId = 800503

    Case "chkp": Me.HelpContextId = 8016
  End Select
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
