Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ANBAR_AfterUpdate()
                      On Error Resume Next
   CANBAR = ANBAR
   Call set_AfterUpdate
End Sub

Private Sub CANBAR_AfterUpdate()
                      On Error Resume Next
   ANBAR = CANBAR
   Call set_AfterUpdate
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
Dim rst As New ADODB.Recordset
Me.mogudi_search.Visible = False
  Me.ANBAR.RowSource = "SELECT     CODE, NAMES FROM         TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE IN (SELECT  ANBCO  FROM dbo.OPANBACCESS  WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & ")))"
  Me.CANBAR.RowSource = "SELECT CODE FROM TCOD_ANBAR WHERE     (CODE <> 0) AND (CODE IN (SELECT  ANBCO  FROM dbo.OPANBACCESS  WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & ")))"
    If MID(Forms![BASEKNOW]![OPTIONSS], 9, 1) = "5" Then
       rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
          Me.ANBAR.DefaultValue = rst.Fields(0)
          Me.CANBAR.DefaultValue = rst.Fields(0)
       End If
    End If
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
          Me.Width = Me.WindowWidth - 100
          Me.mogudi_search.Left = 50
          Me.mogudi_search.Width = Me.Width - 200
          Me.Detail.Height = Me.WindowHeight - 200
          Me.mogudi_search.Height = Me.Detail.Height - 2000
       '   Me.ANBAR.Left = 0.5
       '   Me.CANBAR.Left = Me.ANBAR.Left + Me.ANBAR.Width + 0.3
        '  Me.Label4.Left = Me.CANBAR.Left + Me.CANBAR.Width + 0.2
        End If
        Me.Repaint
    End If
    Me.mogudi_search.Form.RecordSource = " SELECT  DISTINCT    CODE, MAND, ANBAR, ANBARN,  NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES FROM         dbo.AKMOGUDI_KOL_ANBAR(99999999," & Me.ANBAR & ") AKMOGUDI_KOL_ANBAR "
Me.mogudi_search.Visible = True
End Sub

Private Sub set_AfterUpdate()
                      On Error Resume Next
    Dim shar As String
    If IsNull(Me.Set) Then
        shar = ""
    Else
        shar = SHARTCRATOR(Me.Set)
    End If
    Me.mogudi_search.Form.RecordSource = " SELECT    DISTINCT  CODE, MAND, ANBAR, ANBARN,  NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES FROM         dbo.AKMOGUDI_KOL_ANBAR(99999999," & Me.ANBAR & ") AKMOGUDI_KOL_ANBAR where " & IIf(shar = "", " NAME LIKE '%' ", shar)
   ' Debug.Print Me.mogudi_search.Form.RecordSource
End Sub

Private Sub set_Change()
                      On Error Resume Next
    If Me.ONS Then
        Dim shar As String
        If IsNull(Me.Set.Text) Then
            shar = ""
        Else
            shar = SHARTCRATOR(Me.Set.Text)
        End If
         Me.mogudi_search.Form.RecordSource = " SELECT    DISTINCT  CODE, MAND, ANBAR, ANBARN,  NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES FROM         dbo.AKMOGUDI_KOL_ANBAR(99999999," & Me.ANBAR & ") AKMOGUDI_KOL_ANBAR where " & IIf(shar = "", " NAME LIKE '%' ", shar)
    End If
End Sub


Public Function SHARTCRATOR(TXT As String) As String
                      On Error Resume Next
    Dim i, K As Integer, TXTO As String, SHART, SHAT2 As String
   i = 1
   K = 1
   TXT = Trim(TXT)
   While i < Len(TXT)
        If MID(TXT, i, 1) = " " Then
          If Len(SHART) > 0 Then
            SHART = SHART & " and NAME LIKE replace(replace('%" & MID(TXT, K + 1, i - K - 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = shart2 & " and NAME LIKE replace(replace('%" & MID(TXT, K + 1, i - K - 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
            K = i
          Else
            SHART = "NAME LIKE replace(replace('%" & MID(TXT, K, i - K) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = "NAME LIKE replace(replace('%" & MID(TXT, K, i - K) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
            K = i
          End If
        End If
        i = i + 1
   Wend
        If i <> K Then
               If Len(SHART) > 0 Then
                 SHART = SHART & " and NAME LIKE replace(replace('%" & MID(TXT, K + 1, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
                 shart2 = shart2 & " and NAME LIKE replace(replace('%" & MID(TXT, K + 1, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
                 K = i
               Else
                 SHART = "NAME LIKE  replace(replace('%" & MID(TXT, K, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
                 shart2 = "NAME LIKE replace(replace('%" & MID(TXT, K, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
                 K = i
               End If
        End If
    If Len(TXT) = 1 Then
            SHART = "NAME LIKE replace(replace('%" & TXT & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = "NAME LIKE replace(replace('%" & TXT & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
    End If
  If Len(SHART) > 1 Then
    SHART = "(( " & SHART & ") or (" & shart2 & "))"
  End If
  SHARTCRATOR = SHART
End Function
