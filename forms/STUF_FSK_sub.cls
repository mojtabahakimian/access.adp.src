Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NEWR As Boolean
Private Sub ANBAR_AfterUpdate()
    Dim rst As New ADODB.Recordset
                      On Error Resume Next
    If Not IsNull(Me.ANBAR) Then
    If ANBAR.TAG <> ANBAR And ANBAR.TAG <> 0 Then
        rst.Open "SELECT * FROM stuf_stk CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR.TAG, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                Me.ANBAR = Me.ANBAR.TAG
        Else
            If rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - Me.MOGODI_A < Forms![STUF_DEF]![MIN_M] Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تغيير كد انبار مقدار موجودي را به كمتر از حد مجاز مي رساند!"
                Me.ANBAR = Me.ANBAR.TAG
            Else
            End If
        End If
      ' rst.Close
      Set rst = New ADODB.Recordset
    Else
    End If
    If Forms![STUF_DEF]![RADAH] = 1 Then
       rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND  NUMBER = " & Me.ANBAR & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 0 Then
           rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![MOGODIA]
        rst.Fields("NUMBER") = Me.ANBAR
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("BED_BES") = -1
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.update
           'فروش
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = 1 AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        Else
           ' rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM DETA_HES N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Me.CODE
            rst.Fields("NAME") = "فروش " & Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE] & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        End If
        ' rst.Close
        Set rst = New ADODB.Recordset
       'پاياپاي هزينه توليد-پاياپاي مواد مصرفي
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![PHAZ_TOL] & " AND  NUMBER = 1  AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![PHAZ_TOL]
        rst.Fields("NUMBER") = 1
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
 Else
        If Forms![STUF_DEF]![RADAH] = 3 Then
             'پاياپاي هزينه توليد-پاياپاي مواد مصرفي
            rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![PHAZ_TOL] & " AND  NUMBER = 1  AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![PHAZ_TOL]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
        End If
       'فروش
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = 1 AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        Else
           ' rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM DETA_HES N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Me.CODE
            rst.Fields("NAME") = "فروش " & Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE] & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        End If
        ' rst.Close
        Set rst = New ADODB.Recordset
        'انبارمحصول يا نيمه ساخته
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND  NUMBER = " & Me.ANBAR & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![MOGODIA]
        rst.Fields("NUMBER") = Me.ANBAR
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        'كنترل مواد
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![HAZ_TOL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = "مواد مصرفي " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![GHEYMAT]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = " قيمت تمام شده  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![GHEYMAT]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = Me.CODE
        rst.Fields("NAME") = " قيمت تمام شده  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        'كنترل كالاي در جريان ساخت
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' دستمزد
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND TNUMBER = 99999996", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999996
        rst.Fields("NAME") = "دستمزد"
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = 99999997", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999997
        rst.Fields("NAME") = "سربار"
        rst.Fields("BED_BES") = -1
        rst.update

        'جذب دستمزد
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND TNUMBER = 99999999", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999999
        rst.Fields("NAME") = "جذب دستمزد"
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = 99999998", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999998
        rst.Fields("NAME") = "جذب سربار"
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![AMALKARD] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = " عملكرد  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
    End If
    End If
End Sub

Private Sub ANBAR_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
Dim rst As New ADODB.Recordset
If Forms![STUF_DEF]![RADAH] = 1 Then
      rst.Open "SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE (((TCOD_ANBAR.CODE)=" & Me.ANBAR & ") AND ((TCOD_ANBAR.KIND)<>0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
         DoCmd.OpenForm "mesageform", , , , , acDialog, "كالاي فوق متعلق به انبار مواد اوليه مي باشد و نمي تواند در اين انبار قرار گيرد بازدن دوبار كليد ESC از اين وضعيت خارج شويد."
        ' Me.ANBAR.Undo
         CANCEL = True
         Exit Sub
      End If

    End If
End Sub

Private Sub ANBAR_DblClick(CANCEL As Integer)
     Dim rst As New ADODB.Recordset
                     On Error Resume Next
    If Forms![STUF_DEF]![RADAH] = 1 Then
       rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND  NUMBER = " & Me.ANBAR & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 0 Then
           rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![MOGODIA]
        rst.Fields("NUMBER") = Me.ANBAR
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("BED_BES") = -1
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.update
           'فروش
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = 1 AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        Else
           ' rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM DETA_HES N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Me.CODE
            rst.Fields("NAME") = "فروش " & Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE] & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        End If
        ' rst.Close
        Set rst = New ADODB.Recordset
       'پاياپاي هزينه توليد-پاياپاي مواد مصرفي
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![PHAZ_TOL] & " AND  NUMBER = 1  AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![PHAZ_TOL]
        rst.Fields("NUMBER") = 1
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
 Else
        If Forms![STUF_DEF]![RADAH] = 3 Then
             'پاياپاي هزينه توليد-پاياپاي مواد مصرفي
            rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![PHAZ_TOL] & " AND  NUMBER = 1  AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![PHAZ_TOL]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
        End If
       'فروش
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
         '   ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = 1 AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = 1
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        Else
'           ' rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Me.CODE
            rst.Fields("NAME") = "فروش " & Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
            ' rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT * FROM  TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Forms![STUF_DEF]![CODE] & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                rst.AddNew
            End If
            rst.Fields("N_KOL") = Forms![BASEKNOW]![FROSH]
            rst.Fields("NUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.Fields("BED_BES") = -1
            rst.update
        End If
        ' rst.Close
        Set rst = New ADODB.Recordset
        'انبارمحصول يا نيمه ساخته
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND  NUMBER = " & Me.ANBAR & " AND TNUMBER = " & Forms![STUF_DEF]![CODE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![MOGODIA]
        rst.Fields("NUMBER") = Me.ANBAR
        rst.Fields("TNUMBER") = Forms![STUF_DEF]![CODE]
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        'كنترل مواد
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![HAZ_TOL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = "مواد مصرفي " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![GHEYMAT]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = " قيمت تمام شده  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![GHEYMAT]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = Me.CODE
        rst.Fields("NAME") = " قيمت تمام شده  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        'كنترل كالاي در جريان ساخت
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        ' دستمزد
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND TNUMBER = 99999996", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999996
        rst.Fields("NAME") = "دستمزد"
        rst.Fields("BED_BES") = -1
        rst.update
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = 99999997", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999997
        rst.Fields("NAME") = "سربار"
        rst.Fields("BED_BES") = -1
        rst.update

        'جذب دستمزد
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND TNUMBER = 99999999", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999999
        rst.Fields("NAME") = "جذب دستمزد"
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = 99999998", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("TNUMBER") = 99999998
        rst.Fields("NAME") = "جذب سربار"
        rst.Fields("BED_BES") = -1
        rst.update

        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![AMALKARD] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
        rst.Fields("NUMBER") = Me.CODE
        rst.Fields("NAME") = " عملكرد  " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
    End If


End Sub

Private Sub FI_A_AfterUpdate()
                      On Error Resume Next
   If Me.FI_A = 0 Then
     Me.MABL_A.TabStop = True
   Else
     Me.MABL_A = Me.FI_A * Me.MOGODI_A
   End If
End Sub

Private Sub FI_A_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM stuf_stk WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR.TAG, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If ANBAR.TAG <> ANBAR And ANBAR.TAG <> 0 Then
        If rst.RecordCount = 0 Then
             '   DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
               ' Me.ANBAR = Me.ANBAR.TAG
        Else

            rst.Fields("MOGODI_A") = rst.Fields("MOGODI_A") - Me.MOGODI_A.TAG
            rst.update
        End If
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM stuf_stk WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.Fields("CODE") = CODE
            rst.Fields("ANBAR") = ANBAR
            If IsNull(MOGODI_A) Then
               MOGODI_A = 0
            End If
            rst.Fields("MOGODI_A") = MOGODI_A
            rst.update
        Else
            rst.Fields("MOGODI_A") = rst.Fields("MOGODI_A") + Me.MOGODI_A
            rst.update
        End If
        ' rst.Close
    Else
        ' rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM stuf_stk WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("CODE") = CODE
          rst.Fields("ANBAR") = ANBAR
            If IsNull(MOGODI_A) Then
               MOGODI_A = 0
            End If
          rst.Fields("MOGODI_A") = MOGODI_A
          rst.update
          ' rst.Close
        Else
            rst.Fields("MOGODI_A") = rst.Fields("MOGODI_A") - (MOGODI_A.TAG - MOGODI_A)
            rst.update
        End If
    End If

End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
   Response = acDataErrContinue

End Sub

Private Sub Form_Current()
                       On Error Resume Next
  If Me.NewRecord Then
   Me.CODE.TAG = ""
   Me.ANBAR.TAG = 0
   Me.MOGODI_A.TAG = 0
   Me.FI_A.TAG = 0
   Me.MABL_A.TAG = 0
  Else
   Me.CODE.TAG = Me.CODE
   Me.ANBAR.TAG = Me.ANBAR
   Me.MOGODI_A.TAG = Me.MOGODI_A
   Me.FI_A.TAG = Me.FI_A
   Me.MABL_A.TAG = Me.MABL_A
  End If
  Call SETSECURITYSUB("STUF_DEF", Me.NAME, "KALA", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
    rst.Open "SELECT * FROM stuf_stk WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
        CANCEL = True
    Else
       If rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - Me.MOGODI_A < Forms![STUF_DEF]![MIN_M] Then
          DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "حذف كالا مقدار موجودي را به كمتر از حد مجاز مي رساند!"
          CANCEL = True
       Else
          rst.Fields("MOGODI_A") = rst.Fields("MOGODI_A") - Me.MOGODI_A
          rst.update
       End If
    End If
    ' rst.Close
   Else
     CANCEL = True
   End If
End Sub

Private Sub Form_Load()
                    On Error Resume Next
  Call SETSECURITYSUB("STUF_DEF", Me.NAME, "KALA", 3)
End Sub

Private Sub MABL_A_AfterUpdate()
  If Me.MOGODI_A <> 0 Then
     Me.FI_A = Me.MABL_A / Me.MOGODI_A
  End If
End Sub

Private Sub MABL_A_BeforeUpdate(CANCEL As Integer)
  If Me.MOGODI_A = 0 Then
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "موجودي صفر است بنابر اين مبلغ نمي تواند داشته باشد"
    CANCEL = True
  End If

End Sub

Private Sub MABL_A_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub MOGODI_A_AfterUpdate()
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
    If MOGODI_A < MOGODI_A.TAG And ANBAR.TAG = ANBAR And Not Me.NewRecord Then
        rst.Open "SELECT * FROM stuf_stk WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                MOGODI_A = MOGODI_A.TAG
        Else
            If rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - (MOGODI_A.TAG - MOGODI_A) < Forms![STUF_DEF]![MIN_M] Then
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تغيير موجودي مقدار كالا را به كمتر از حد مجاز مي رساند!"
                MOGODI_A = MOGODI_A.TAG
            Else
            End If
        End If
        ' rst.Close
    Else
    End If
   If Me.FI_A = 0 Then
     Me.MABL_A.TabStop = True
   Else
     Me.MABL_A = Me.FI_A * Me.MOGODI_A
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub MOGODI_A_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub
