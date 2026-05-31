Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Dim tg As Integer, num As Integer, Meidnum As Double
Private Sub Command120_Click()
                 On Error Resume Next

    Select Case tg
        Case 20:
            If IsLoaded("HEAD_LST_PISHFROOSH2") Then
                Forms![HEAD_LST_PISHFROOSH2]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
        Case 100:
            If IsLoaded("PAYMENTFORMORDER") Then
                Forms![PAYMENTFORMORDER]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
                Forms![PAYMENTFORMORDER]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid") = Forms![BASEKNOW]![USERCOD]
            End If
            CurrentProject.Connection.Execute ("UPDATE PAYORDER SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 1 WHERE  IDD = " & num)
            CurrentProject.Connection.Execute ("UPDATE PAYORDER SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid" & " = " & Forms![BASEKNOW]![USERCOD] & " WHERE  IDD = " & num)
        Case 90:
            If IsLoaded("PTAMIR") Then
                Forms![PTAMIR]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
                Forms![PTAMIR]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid") = Forms![BASEKNOW]![USERCOD]
            End If
            CurrentProject.Connection.Execute ("UPDATE PTAMIRAT SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 1 WHERE  PIDT = " & num)
            CurrentProject.Connection.Execute ("UPDATE PTAMIRAT SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid" & " = " & Forms![BASEKNOW]![USERCOD] & " WHERE  PIDT = " & num)
        Case 2:
            If IsLoaded("HEAD_LST_HAVL") Then
                Forms![head_lst_havl]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
        Case 6:
            If GETUSERTASK(Meidnum) = Forms![BASEKNOW]![USERCOD] Then
                If IsLoaded("HEAD_LST_ENTEGHAL") Then
                    Forms![HEAD_LST_ENTEGHAL]("SGN1") = True
                End If
                CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN1 = 1 WHERE TAG =5 AND NUMBER = " & num)
            Else
                If IsLoaded("HEAD_LST_ENTEGHAL") Then
                    Forms![HEAD_LST_ENTEGHAL]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
                End If
                CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 1 WHERE TAG =5 AND NUMBER = " & num)
            End If
        Case 1:
            If IsLoaded("HEAD_LST_RASID") Then
                Forms![head_lst_rasid]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
        Case 13:
            If IsLoaded("HEAD_LST_FROOSH2") Then
                Forms![HEAD_LST_FROOSH2]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
            If IsLoaded("HEAD_LST_FROOSH22") Then
                Forms![HEAD_LST_FROOSH22]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
        Case 12:
            If IsLoaded("HEAD_LST_KHAREED1") Then
                Forms![head_lst_khareed1]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
            End If
        Case 32:
            If IsLoaded("HEAD_LST_KHAREED1_SADER") Then
                Forms![HEAD_LST_KHAREED1_SADER]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = True
                CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(12, Forms![BASEKNOW]![USERCOD]) & " = 1 WHERE TAG = 12 AND NUMBER = " & num)
            End If
    End Select
    CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & Meidnum & ",'" & UCurrentUser() & "','" & "مورد تاييد است " & IIf(IsNull(Me.TXT), " ", Me.TXT) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & num & "," & tg & " )")
    If tg <> 6 Then
        CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 1 WHERE TAG = " & tg & " AND NUMBER = " & num)
        CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid" & " = 1 WHERE TAG = " & tg & " AND NUMBER = " & num)
    End If
    If tg <> 100 Then
          CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(Meidnum) & " WHERE IDNUM = " & Meidnum)
          DoCmd.OpenForm "mesageform", , , , , acDialog, "  تاييد شد و به کارتابل فرستنده ارسال شد"
    Else
          DoCmd.OpenForm "mesageform", , , , , acDialog, "  تاييد شد و به کارتابل شما ارسال شد . در کارتابل به شخص مربوطه ارجاع دهيد."
    End If
    Forms![main]![TASKS].Form.Refresh
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command121_Click()
                 On Error Resume Next
   CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & Meidnum & ",'" & UCurrentUser() & "','" & "اصلاحيه لازم است:" & IIf(IsNull(Me.TXT), " ", Me.TXT) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & num & "," & tg & " )")
   CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(Meidnum) & " WHERE IDNUM = " & Meidnum)
   DoCmd.OpenForm "mesageform", , , , , acDialog, "  اصلاحيه لازم است  به کارتابل فرستنده ارسال شد"
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command122_Click()
                On Error Resume Next
    Select Case tg
        Case 20:
            If IsLoaded("HEAD_LST_PISHFROOSH2") Then
                Forms![HEAD_LST_PISHFROOSH2]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 2:
            If IsLoaded("HEAD_LST_HAVL") Then
                Forms![head_lst_havl]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 1:
            If IsLoaded("HEAD_LST_RASID") Then
                Forms![head_lst_rasid]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 6:
            If IsLoaded("HEAD_LST_ENTEGHAL") Then
                Forms![HEAD_LST_ENTEGHAL]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 100:
            If IsLoaded("PAYMENTFORMORDER") Then
                Forms![PAYMENTFORMORDER]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
            CurrentProject.Connection.Execute ("UPDATE PAYORDER SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 0 WHERE IDD = " & num)
        Case 90:
            If IsLoaded("PTAMIR") Then
                Forms![PTAMIR]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
            CurrentProject.Connection.Execute ("UPDATE PTAMIRAT SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 0 WHERE  PIDT = " & num)
        Case 13:
            If IsLoaded("HEAD_LST_FROOSH2") Then
                Forms![HEAD_LST_FROOSH2]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
            If IsLoaded("HEAD_LST_FROOSH22") Then
                Forms![HEAD_LST_FROOSH22]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 12:
            If IsLoaded("HEAD_LST_KHAREED1") Then
                Forms![head_lst_khareed1]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
            End If
        Case 32:
            If IsLoaded("HEAD_LST_KHAREED1") Then
                Forms![HEAD_LST_KHAREED1_SADER]("SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD])) = False
                CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(12, Forms![BASEKNOW]![USERCOD]) & " = 0 WHERE TAG = 12 AND NUMBER = " & num)
            End If
    End Select
    CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & Meidnum & ",'" & UCurrentUser() & "','" & "مورد تاييد نميباشد " & IIf(IsNull(Me.TXT), " ", Me.TXT) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & num & "," & tg & " )")
    CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & " = 0 WHERE TAG = " & tg & " AND NUMBER = " & num)
    CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET SGN" & GETSIGNER(tg, Forms![BASEKNOW]![USERCOD]) & "usid" & " = 1 WHERE TAG = " & tg & " AND NUMBER = " & num)
    CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(Meidnum) & " WHERE IDNUM = " & Meidnum)
    DoCmd.OpenForm "mesageform", , , , , acDialog, "  تاييد نشد و به کارتابل فرستنده ارسال شد"
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Command4_Click()
                On Error Resume Next
    Forms![BASEKNOW]![MATN] = Me.TXT
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command8_Click()
    DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                On Error Resume Next
    tg = MID(Me.OpenArgs, 1, InStr(1, Me.OpenArgs, "-") - 1)
    num = MID(Me.OpenArgs, InStr(1, Me.OpenArgs, "-") + 1, InStr(3, Me.OpenArgs, "+") - 4)
   Meidnum = MID(Me.OpenArgs, InStr(3, Me.OpenArgs, "+") + 1, Len(Me.OpenArgs) - InStr(3, Me.OpenArgs, "+"))

End Sub
