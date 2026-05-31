Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Private Sub Combo34_AfterUpdate()
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
  If Not IsNull(Me.Combo34) Then
    Call GETTAF3(Me.Combo34, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    Me.HKOL = KOL
    Me.HMOIN = MOIN
    Me.HTAF = taf
    Me.HTAF2 = TAF2
    Me.HTAF3 = TAF3
    Me.HTAF4 = TAF4
  End If
  Me.Combo36 = Me.Combo34
  Me.Combo36.Requery
End Sub

Private Sub Combo34_Exit(CANCEL As Integer)
    Call Combo34_AfterUpdate
    If Not IsNull(Me.Combo34) Then
        If ISTAF(Me.Combo34) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub Combo34_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.Combo34.Text = "-" Or Me.Combo34.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "05"
  End If
  Response = acDataErrContinue
End Sub

Private Sub Combo36_AfterUpdate()
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   Me.Combo34 = Me.Combo36
  If Not IsNull(Me.Combo34) Then
    Call GETTAF3(Me.Combo34, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    Me.HKOL = KOL
    Me.HMOIN = MOIN
    Me.HTAF = taf
    Me.HTAF2 = TAF2
    Me.HTAF3 = TAF3
    Me.HTAF4 = TAF4
  End If
   Me.Combo34.Requery
End Sub

Private Sub Combo36_Exit(CANCEL As Integer)
    If Not IsNull(Me.Combo34) Then
        If ISTAF(Me.Combo34) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub Command5_Click()
            On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, SORTT As String, i As Integer, MAN As Double
  If IsNull(Me.HTAF2) Then
    Me.HTTAF = Me.HKOL & "-" & Me.HMOIN & "-" & Me.HTAF
  Else
    Me.HTTAF = Me.HKOL & "-" & Me.HMOIN & "-" & Me.HTAF & "-" & Me.HTAF2
  End If
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.HKOL) Or IsNull(Me.HMOIN) Or IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Select Case Frame34
    Case 1: SORTT = "N_S, BED DESC"
    Case 2: SORTT = "base, BED DESC"
    Case 3: SORTT = "DATE_S, BED DESC"
  End Select
  Me.txtSort = SORTT
If BLOCKEDMK(Me.Combo34) Then
   DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
   Exit Sub
End If
If Not Frisok(Me.Combo34) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "بعضي از فاکتور هاي اين مشتري داراي مغايرت باسند حسابداري است فاکتورهاي داراي اشکال در فايل errorfr  در درايو c ميتوانيد ببينيد"
End If
If Not Khisok(Me.Combo34) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "بعضي از فاکتور هاي خريد اين فروشنده داراي مغايرت باسند حسابداري است فاکتورهاي داراي اشکال در فايل errorKH  در درايو c ميتوانيد ببينيد"
End If
    CurrentProject.Connection.Execute ("INSERT INTO AMALIAT (USERID,USERNAME,ADATE,AMALID) VALUES (" & UCurrentUserIDD & ",'" & Me.Combo34 & "',GETDATE(),'" & Me.OpenArgs & "')")
If Me.OpenArgs = "TAF" Then
    If Me.Frame34 = 1 Then
       DoCmd.OpenReport "R_DAFTAR_TAFZILY__2", acViewPreview
    Else
       If Me.Frame34 = 2 Then
           DoCmd.OpenReport "R_DAFTAR_TAFZILY2__2", acViewPreview
       Else
           DoCmd.OpenReport "R_DAFTAR_TAFZILY_2_2", acViewPreview
       End If
    End If
    DoCmd.OpenForm Me.NAME, , , , , acHidden
    DoCmd.OpenForm Me.NAME, , , , , acHidden
End If
If Me.OpenArgs = "VAZ" Then
    DoCmd.OpenReport "R_GARDESH_KHFR_DAFTAR", acViewPreview
    DoCmd.OpenForm Me.NAME, , , , , acHidden
End If
If Me.OpenArgs = "FRKMA4" Then
    DoCmd.OpenReport "R_GARDESH_KHFR_DAFTAR_A4", acViewPreview
    DoCmd.OpenForm Me.NAME, , , , , acHidden
End If
 If Me.OpenArgs = "RMOIN" Then
    '    CurrentProject.Connection.Execute ("IF EXISTS (SELECT 1 FROM sys.objects WHERE type = 'P' AND name = 'usp_TafzilLedger') DROP PROCEDURE usp_TafzilLedger")
'        CurrentProject.Connection.Execute ("CREATE  PROC [dbo].[usp_TafzilLedger] @FromDate int,@ToDate int,@TafzilCode nvarchar(50),@SortExpr nvarchar(400)=N'DATE_S, BED DESC' AS BEGIN SET ARITHABORT ON;SET NOCOUNT ON;" & _
                                            " DECLARE @WhiteList TABLE(Col sysname PRIMARY KEY);INSERT @WhiteList(Col) VALUES('N_S'),('DATE_S'),('base'),('HES_K'),('HES_M'),('HES_T'),('HES_T2'),('TAFZILN'),('SHARH'),('BED'),('BES'),('MAND'),('NO_S'),('N_SERI'),('BANK'),('NUMBER'),('TAG'),('ARZD'),('id'),('HES');" & _
                                            " DECLARE @ExprUpper nvarchar(4000);DECLARE @ExprNoDir nvarchar(4000);DECLARE @Xml xml;SET @ExprUpper=UPPER(@SortExpr);SET @ExprNoDir=REPLACE(REPLACE(@ExprUpper,' DESC',''),' ASC','');SET @Xml=CAST('<r>'+REPLACE(@ExprNoDir,',','</r><r>')+'</r>' AS xml);" & _
                                            " DECLARE @Tokens TABLE(token sysname);INSERT @Tokens(token) SELECT LTRIM(RTRIM(N.value('.','nvarchar(100)'))) FROM @Xml.nodes('/r') AS T(N);IF EXISTS(SELECT 1 FROM @Tokens t WHERE NOT EXISTS(SELECT 1 FROM @WhiteList w WHERE w.Col=t.token)) BEGIN RAISERROR(N'يک يا چند نام ستون غيرمجاز در SortExpr وجود دارد.',16,1);RETURN;END;" & _
                                            " DECLARE @sql nvarchar(max)=N';WITH CTE AS(SELECT N_S,DATE_S,HES_K,HES_M,HES_T,HES_T2,TAFZILN,SHARH,BED,BES,NO_S,N_SERI,BANK,[NUMBER],TAG,ARZD,base,id,DiffAmt=BED-BES FROM dbo.QDAFTARTAFZIL2_H(@FromDate,@ToDate,@TafzilCode)),Ordered AS(SELECT *,rn=ROW_NUMBER() OVER(ORDER BY '+@SortExpr+N', id) FROM CTE) SELECT o.N_S,o.DATE_S,o.HES_K,o.HES_M,o.HES_T,o.HES_T2,o.TAFZILN,o.SHARH,o.BED,o.BES,ABS(c.SumAmt) AS MAND,CASE WHEN c.SumAmt>0 THEN N''بد'' WHEN c.SumAmt=0 THEN N''--'' ELSE N''بس'' END AS TASH,@TafzilCode AS hes,o.NO_S,o.N_SERI,o.BANK,o.[NUMBER],o.TAG,o.ARZD,o.base FROM Ordered o CROSS APPLY (SELECT SumAmt=SUM(i.DiffAmt) FROM Ordered i WHERE i.rn<=o.rn) c ORDER BY '+@SortExpr+N', id;';" & _
                                            " EXEC sp_executesql @sql,N'@FromDate int,@ToDate int,@TafzilCode nvarchar(50)',@FromDate,@ToDate,@TafzilCode;END")
      If Not IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
            With Forms!R_DAFTAR_MOIN_LIST
                .RecordSource = "EXEC usp_TafzilLedger  " & Me.DT1 & " , " & Me.DT2 & " , '" & Me.Combo34 & "', '" & SORTT & "'"
                'Requery                  ' بارگذاري مجدد داده
                .Visible = True           ' نمايش فرم
            End With
          DoCmd.OpenForm Me.NAME, , , , , acHidden
      Else
      If Not IsLoaded("R_DAFTAR_MOIN_LIST2") Then
            DoCmd.OpenForm "R_DAFTAR_MOIN_LIST2", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
            With Forms!R_DAFTAR_MOIN_LIST2
                .RecordSource = "EXEC usp_TafzilLedger  " & Me.DT1 & " , " & Me.DT2 & " , '" & Me.Combo34 & "', '" & SORTT & "'"
                'Requery                  ' بارگذاري مجدد داده
                .Visible = True           ' نمايش فرم
            End With
            DoCmd.OpenForm Me.NAME, , , , , acHidden
      Else
            DoCmd.OpenForm "R_DAFTAR_MOIN_LIST3", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
            With Forms!R_DAFTAR_MOIN_LIST3
                .RecordSource = "EXEC usp_TafzilLedger  " & Me.DT1 & " , " & Me.DT2 & " , '" & Me.Combo34 & "', '" & SORTT & "'"
                'Requery                  ' بارگذاري مجدد داده
                .Visible = True           ' نمايش فرم
            End With
            DoCmd.OpenForm Me.NAME, , , , , acHidden
          End If
     End If
     DoCmd.OpenForm Me.NAME, , , , , acHidden
 End If
 If Me.OpenArgs = "NABZMOSH" Then
     DoCmd.OpenForm "NABZ_MOSHTARI", , , , , , Me.Combo34
 End If
 If Me.OpenArgs = "NABZKAR" Then
     DoCmd.OpenForm "NABZ_KARSHENAS", , , , , , Me.Combo34
 End If
End Sub

Private Sub Command6_Click()
                     On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub


Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  DoCmd.restore
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub HTAF_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.HTAF.Text = "-" Or Me.HTAF.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "05"
  End If
Response = acDataErrContinue
End Sub
