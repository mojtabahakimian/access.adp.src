Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
  Dim incmd$, i, dt As String, fs, a, TEXTST As String
  dt = "Version :7.1.1: date : " & Format(FARSIDATEl(DATE), "####/##/##") & "  " & Format(Now(), "hh:mm")
  DoCmd.OpenForm "About", acDesign

  Forms![About].version.CAPTION = dt
  DoCmd.Save
  DoCmd.Close acForm, "About", acSaveYes

  DoCmd.OpenForm "YE", acDesign
  Forms![YE].version.CAPTION = dt
  DoCmd.Save
  DoCmd.Close acForm, "YE", acSaveYes

  DoCmd.OpenForm "USER_LOGIN", acDesign

  Forms![USER_LOGIN].version.CAPTION = dt
  DoCmd.Save
  DoCmd.Close acForm, "USER_LOGIN", acSaveYes
  DoCmd.OpenForm "startup", acDesign
  Forms![Startup].version.CAPTION = dt
  DoCmd.Save
  DoCmd.Close acForm, "startup", acSaveYes
  DoCmd.OpenForm "KALAS_MAIN", acDesign
  Forms![KALAS_MAIN].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "KALAS_MAIN", acSaveYes

  DoCmd.OpenForm "TDETA_HES_SHEET", acDesign
  Forms![TDETA_HES_SHEET].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TDETA_HES_SHEET", acSaveYes


  DoCmd.OpenForm "SERCHKAL", acDesign
  Forms![SERCHKAL].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "SERCHKAL", acSaveYes


  DoCmd.OpenForm "TOTA_HES_SHEET", acDesign
  Forms![TOTA_HES_SHEET].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TOTA_HES_SHEET", acSaveYes


  DoCmd.OpenForm "TARAZ_4_MOIN", acDesign
  Forms![TARAZ_4_MOIN].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TARAZ_4_MOIN", acSaveYes

  DoCmd.OpenForm "TDETA_HES_SHEET2", acDesign
  Forms![TDETA_HES_SHEET2].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TDETA_HES_SHEET2", acSaveYes

  DoCmd.OpenForm "TDETA_HES_SHEET3", acDesign
  Forms![TDETA_HES_SHEET3].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TDETA_HES_SHEET3", acSaveYes

  DoCmd.OpenForm "TDETA_HES_SHEET4", acDesign
  Forms![TDETA_HES_SHEET4].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "TDETA_HES_SHEET4", acSaveYes

  DoCmd.OpenForm "DETA_HES_SHEET", acDesign
  Forms![DETA_HES_SHEET].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "DETA_HES_SHEET", acSaveYes

  DoCmd.OpenForm "HEAD_LST_FROOSH22", acDesign
  Forms![HEAD_LST_FROOSH22].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "HEAD_LST_FROOSH22", acSaveYes


  DoCmd.OpenForm "HEAD_LST_FROOSH2", acDesign
  Forms![HEAD_LST_FROOSH2].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "HEAD_LST_FROOSH2", acSaveYes

  DoCmd.OpenForm "HEAD_LST_RASID", acDesign
  Forms![head_lst_rasid].AllowEdits = True
  Forms![head_lst_rasid].AllowDeletions = True
  Forms![head_lst_rasid].AllowAdditions = True
  DoCmd.Save
  DoCmd.Close acForm, "HEAD_LST_RASID", acSaveYes

  DoCmd.OpenForm "INVO_LST_RASID_SUB", acDesign
  Forms![INVO_LST_RASID_SUB].AllowEdits = True
  Forms![INVO_LST_RASID_SUB].AllowDeletions = True
  Forms![INVO_LST_RASID_SUB].AllowAdditions = True
  DoCmd.Save
  DoCmd.Close acForm, "INVO_LST_RASID_SUB", acSaveYes

  DoCmd.OpenForm "DEED_HEAD", acDesign
  Forms![DEED_HEAD]![Child14].Form.AllowEdits = True
  Forms![DEED_HEAD]![Child14].Form.AllowDeletions = True
  Forms![DEED_HEAD]![Child14].Form.AllowAdditions = True
  Forms![DEED_HEAD].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "DEED_HEAD", acSaveYes

  DoCmd.OpenForm "HEAD_LST_PISHFROOSH2", acDesign
  Forms![HEAD_LST_PISHFROOSH2].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "HEAD_LST_PISHFROOSH2", acSaveYes

  DoCmd.OpenForm "HEAD_LST_HAVL", acDesign
  Forms![head_lst_havl].ServerFilter = ""
  DoCmd.Save
  DoCmd.Close acForm, "HEAD_LST_HAVL", acSaveYes

  Err.Clear
    i = 0
 '  Set fs = CreateObject("Scripting.FileSystemObject")
 '  Set B = fs.CREATETextFile(CurrentProject.PATH & "\" & "negin1.scp", True)
'   B.writeline dt'
  ' Set a = fs.OPENTextFile(CurrentProject.PATH & "\" & "negin.scp", 1)
  ' TEXTST = a.readline
  ' While Not a.AtEndOfStream
  '     TEXTST = a.readline
  '     B.writeline TEXTST
  '  Wend
  ' a.Close
  ' B.Close
  ' Kill CurrentProject.PATH & "\" & "negin.scp"
  ' Name CurrentProject.PATH & "\" & "negin1.scp" As CurrentProject.PATH & "\" & "negin.scp"
  ' FileCopy CurrentProject.PATH & "\" & "negin.scp", "f:\software\ade\negin.scp"
  ' FileCopy CurrentProject.PATH & "\" & "negin.scp", "C:\negin.scp"
 DoCmd.Close acForm, Me.NAME
End Sub
