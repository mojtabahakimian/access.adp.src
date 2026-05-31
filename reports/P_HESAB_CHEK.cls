Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
   rst.Open "FORMAT_HESAB", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = ("BANK = " & Forms![CHK_V_PRINT]![CHK_V_PRINTSUB].Form![BKK] & " And KIND = " & Forms![CHK_V_PRINT]![CHK])
    Set rst = rst
   If rst.RecordCount = 0 Then
   Else
       Me.DATE.Width = rst.Fields("DATE_WITH") * 567
       Me.DATE.Top = rst.Fields("DATE_TOP") * 567
       Me.DATE.Left = rst.Fields("DATE_LEFT") * 567
       Me.DATE.FontName = rst.Fields("DATE_FONT")
       Me.DATE.FontSize = rst.Fields("DATE_SIZE")
       Me.DATE.Visible = rst.Fields("DATE_p")

       Me.SHO.Width = rst.Fields("SHO_WITH") * 567
       Me.SHO.Top = rst.Fields("SHO_TOP") * 567
       Me.SHO.Left = rst.Fields("SHO_LEFT") * 567
       Me.SHO.FontName = rst.Fields("SHO_FONT")
       Me.SHO.FontSize = rst.Fields("SHO_SIZE")
       Me.SHO.Visible = rst.Fields("SHO_p")

       Me.SHES.Width = rst.Fields("SHES_WITH") * 567
       Me.SHES.Top = rst.Fields("SHES_TOP") * 567
       Me.SHES.Left = rst.Fields("SHES_LEFT") * 567
       Me.SHES.FontName = rst.Fields("SHES_FONT")
       Me.SHES.FontSize = rst.Fields("SHES_SIZE")
       Me.SHES.Visible = rst.Fields("SHES_p")

       Me.NAMEE.Width = rst.Fields("NAME_WITH") * 567
       Me.NAMEE.Top = rst.Fields("NAME_TOP") * 567
       Me.NAMEE.Left = rst.Fields("NAME_LEFT") * 567
       Me.NAMEE.FontName = rst.Fields("NAME_FONT")
       Me.NAMEE.FontSize = rst.Fields("NAME_SIZE")
       Me.NAMEE.Visible = rst.Fields("NAME_p")

       Me.CDATE.Width = rst.Fields("CDATE_WITH") * 567
       Me.CDATE.Top = rst.Fields("CDATE_TOP") * 567
       Me.CDATE.Left = rst.Fields("CDATE_LEFT") * 567
       Me.CDATE.FontName = rst.Fields("CDATE_FONT")
       Me.CDATE.FontSize = rst.Fields("CDATE_SIZE")
       Me.CDATE.Visible = rst.Fields("CDATE_p")

       Me.SER.Width = rst.Fields("SER_WITH") * 567
       Me.SER.Top = rst.Fields("SER_TOP") * 567
       Me.SER.Left = rst.Fields("SER_LEFT") * 567
       Me.SER.FontName = rst.Fields("SER_FONT")
       Me.SER.FontSize = rst.Fields("SER_SIZE")
       Me.SER.Visible = rst.Fields("SER_p")

       Me.CHES.Width = rst.Fields("CHES_WITH") * 567
       Me.CHES.Top = rst.Fields("CHES_TOP") * 567
       Me.CHES.Left = rst.Fields("CHES_LEFT") * 567
       Me.CHES.FontName = rst.Fields("CHES_FONT")
       Me.CHES.FontSize = rst.Fields("CHES_SIZE")
       Me.CHES.Visible = rst.Fields("CHES_p")

       Me.BSHO.Width = rst.Fields("BSHO_WITH") * 567
       Me.BSHO.Top = rst.Fields("BSHO_TOP") * 567
       Me.BSHO.Left = rst.Fields("BSHO_LEFT") * 567
       Me.BSHO.FontName = rst.Fields("BSHO_FONT")
       Me.BSHO.FontSize = rst.Fields("BSHO_SIZE")
       Me.BSHO.Visible = rst.Fields("BSHO_p")

       Me.mabl.Width = rst.Fields("MAB_WITH") * 567
       Me.mabl.Top = rst.Fields("MAB_TOP") * 567
       Me.mabl.Left = rst.Fields("MAB_LEFT") * 567
       Me.mabl.FontName = rst.Fields("MAB_FONT")
       Me.mabl.FontSize = rst.Fields("MAB_SIZE")
       Me.mabl.Visible = rst.Fields("MAB_p")

       Me.TEXTMABL.Width = rst.Fields("AMAB_WITH") * 567
       Me.TEXTMABL.Top = rst.Fields("AMAB_TOP") * 567
       Me.TEXTMABL.Left = rst.Fields("AMAB_LEFT") * 567
       Me.TEXTMABL.FontName = rst.Fields("AMAB_FONT")
       Me.TEXTMABL.FontSize = rst.Fields("AMAB_SIZE")
       Me.TEXTMABL.Visible = rst.Fields("AMAB_p")

       Me.CITY.Width = rst.Fields("CITY_WITH") * 567
       Me.CITY.Top = rst.Fields("CITY_TOP") * 567
       Me.CITY.Left = rst.Fields("CITY_LEFT") * 567
       Me.CITY.FontName = rst.Fields("CITY_FONT")
       Me.CITY.FontSize = rst.Fields("CITY_SIZE")
       Me.CITY.Visible = rst.Fields("CITY_p")
   End If
   rst.Close
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
