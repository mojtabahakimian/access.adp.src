Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Dim nam As String

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim i As Integer, SHART, shart2 As String, K As Integer, anbars As String
                     On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
   If KeyAscii = 13 Then
   i = 1
   K = 1
   Text0 = Trim(Text0)
   While i < Len(Text0)
        If MID(Text0, i, 1) = " " Then
          If Len(SHART) > 0 Then
            SHART = SHART & " and NAME LIKE replace(replace('%" & MID(Text0, K + 1, i - K - 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = shart2 & " and NAME LIKE replace(replace('%" & MID(Text0, K + 1, i - K - 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
            K = i
          Else
            SHART = "NAME LIKE replace(replace('%" & MID(Text0, K, i - K) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = "NAME LIKE replace(replace('%" & MID(Text0, K, i - K) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
            K = i
          End If
        End If
        i = i + 1
   Wend
        If i <> K Then
               If Len(SHART) > 0 Then
                 SHART = SHART & " and NAME LIKE replace(replace('%" & MID(Text0, K + 1, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
                 shart2 = shart2 & " and NAME LIKE replace(replace('%" & MID(Text0, K + 1, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
                 K = i
               Else
                 SHART = "NAME LIKE  replace(replace('%" & MID(Text0, K, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
                 shart2 = "NAME LIKE replace(replace('%" & MID(Text0, K, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
                 K = i
               End If
        End If
    If Len(Text0) = 1 Then
            SHART = "NAME LIKE replace(replace('%" & Text0 & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
            shart2 = "NAME LIKE replace(replace('%" & Text0 & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
    End If
   Select Case nam
      Case "F_MENU_ANBAR_FRKH": anbars = Forms![F_MENU_ANBAR_FRKH].Form![ANBAR]
      Case "F_MENU_KART": anbars = Forms![F_MENU_KART].Form![ANBAR]
      Case "F_MENU_ANBAR": anbars = Forms![F_MENU_ANBAR].Form![ANBAR]
      Case "HEAD_LST_MOSTAGHIM": anbars = Forms![HEAD_LST_MOSTAGHIM]![INVO_LST_MO_SUB].Form![ANBAR]
      Case "HEAD_LST_RASID": anbars = Forms![head_lst_rasid]![INVO_LST_RASID_SUB].Form![ANBAR]
      Case "HEAD_LST_HAVL": anbars = Forms![head_lst_havl]![INVO_LST_HAVL_SUB].Form![ANBAR]
      Case "OHEAD_LST_FROOSH": anbars = Forms![OHEAD_LST_FROOSH]![OINVOLST_SUB].Form![ANBAR]
      Case "HEAD_LST_PISHFROOSH": anbars = Forms![head_lst_pishfroosh]![INVO_LST_PISHFROOSH_SUB].Form![ANBAR]
      Case "HEAD_LST_PISHFROOSH2": anbars = Forms![HEAD_LST_PISHFROOSH2]![INVO_LST_PISHFROOSH_SUB].Form![ANBAR]
      Case "HEAD_LST_FROOSH": anbars = Forms![HEAD_LST_FROOSH]![INVO_LST_sub].Form![ANBAR]
      Case "HEAD_LST_KHADAMAT": anbars = Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![ANBAR]
      Case "HEAD_LST_KHAREED": anbars = Forms![head_lst_khareed]![INVO_LST_KH_SUB].Form![ANBAR]
      Case "HAVALAH_EXIT": anbars = Forms![HAVALAH_EXIT]![HAVALA_EXIT_SUB].Form![ANBAR]
      Case "HEAD_LST_ENTEGHAL": anbars = Forms![HEAD_LST_ENTEGHAL]![INVO_LST_ENTEGHAL_SUB].Form![ANBAR]
      Case "HAVALAH_ENTER": anbars = Forms![havalah_enter]![HAVALAH_ENTER_SUB].Form![ANBAR]
      Case "HEAD_MANF": anbars = Forms![HEAD_MANF]![DTL_MANF_SUB1].Form![ANBAR]
      Case "HAVALAH_EXIT_SAYER": anbars = Forms![havalah_exit_sayer]![HAVALAH_EXIT_SAYER_SUB].Form![ANBAR]
      Case "HEAD_LST_REQUEST": anbars = Forms![HEAD_LST_REQUEST]![INVO_LST_REQUEST_SUB].Form![ANBAR]
      Case "HEAD_LST_RASID_OTHER": anbars = Forms![HEAD_LST_RASID_OTHER]![INVO_LST_RASID_SUB_OTHER].Form![ANBAR]
      Case "HEAD_LST_RASID_OTHER": anbars = Forms![HEAD_LST_HAV_other]![INVO_LST_HAV_SUB_OTHER].Form![ANBAR]
  End Select
  If Len(SHART) > 1 Then
    SHART = "(( " & SHART & ") or (" & shart2 & "))"
  End If
  If Forms![BASEKNOW]![KALA] = 2 And nam <> "ORDR_HED" And nam <> "VISITORS_PORSANT_HEAD" Then
        If IsNull(Text0) Then
            SHART = SHART & "MOG > 0"
            If IsNull(anbars) Or anbars = "" Then
              DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART, , , nam
            Else
              DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART & "and anbar = " & anbars, , , nam
            End If
        Else
            SHART = SHART & "and MOG > 0"
            If IsNull(anbars) Or anbars = "" Then
              DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART, , , nam
            Else
              DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART & " and anbar = " & anbars, , , nam
            End If
        End If
  Else
      If IsNull(Text0) Then
          If IsNull(anbars) Or anbars = "" Then
            DoCmd.OpenForm "SERCHKAL", acFormDS, , , , , nam
          Else
            DoCmd.OpenForm "SERCHKAL", acFormDS, , "anbar = " & anbars, , , nam
          End If

      Else
          If IsNull(anbars) Or anbars = "" Then
            DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART, , , nam
          Else
            DoCmd.OpenForm "SERCHKAL", acFormDS, , SHART & " and anbar = " & anbars, , , nam
          End If
      End If
    End If
    DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub Form_Load()
                      On Error Resume Next
   nam = Screen.Application.CurrentObjectName
   DoCmd.Beep
End Sub
