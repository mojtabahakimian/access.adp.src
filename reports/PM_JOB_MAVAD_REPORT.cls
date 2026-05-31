Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_kart"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                   On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    Me.JO_PERSON.RowSource = GetUserList
    Me.RecordSource = "SELECT PM_MVAD.JOB_IDD, PM_MVAD.MV_ANBAR, PM_MVAD.MV_CODE, PM_MVAD.MV_MEGH, PM_MVAD.MV_MEGHk, PM_MVAD.MV_VAHED_K, PM_MVAD.MV_DESCRIPTION, PM_JOB.AS_IDD, PM_JOB.JO_STDATE, PM_JOB.JO_JOB, PM_JOB.JO_PERSON FROM PM_JOB INNER JOIN PM_MVAD ON PM_JOB.IDD = PM_MVAD.JOB_IDD WHERE (PM_MVAD.JOB_IDD IN (select idd from pm_job where " & Me.OpenArgs & "))"
End Sub
