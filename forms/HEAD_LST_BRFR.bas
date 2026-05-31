Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =15090
    DatasheetFontHeight =10
    ItemSuffix =118
    Left =3135
    Top =630
    Right =18225
    Bottom =10290
    HelpContextId =6002
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x6817d2c48e16e640
    End
    RecordSource ="SELECT NUMBER, TAG AS htag, 24 AS Dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N"
        "_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, M"
        "OIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, TICMBAA, MBAA, HMBAA, DEP"
        "ATMAN, SHIFT, CUST_KIND, OKF, SADER, SGN1, SGN2, SGN3, SGN4, sgn1usid, sgn2usid,"
        " sgn3usid, UID FROM HEAD_LST WHERE (TAG = 25) ORDER BY NUMBER"
    Caption ="فاكتور برگشت فروش رسيد شده"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9675
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Style =1
                    Width =15090
                    Height =9675
                    FontWeight =700
                    Name ="TabCtl56"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =105
                            Top =330
                            Width =14850
                            Height =9210
                            Name ="Page57"
                            Caption ="&روي فاكتور"
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11990
                                    Top =589
                                    Width =951
                                    Height =315
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="NUMBER1"
                                    ControlSource ="NUMBER1"
                                    DefaultValue ="0"
                                    Tag ="1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13085
                                            Top =594
                                            Width =1110
                                            Height =300
                                            Name ="Label1"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =9020
                                    Top =590
                                    Width =1761
                                    Height =315
                                    TabIndex =1
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="DATE_N"
                                    ControlSource ="DATE_N"
                                    BeforeUpdate ="[Event Procedure]"
                                    DefaultValue ="=FARSIDATE(Date())"
                                    InputMask ="####/##/##"
                                    Tag ="890222"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =10880
                                            Top =590
                                            Width =1065
                                            Height =300
                                            Name ="Label9"
                                            Caption ="تاريخ فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6770
                                    Top =575
                                    Width =966
                                    Height =315
                                    TabIndex =2
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="FNUMCO"
                                    ControlSource ="FNUMCO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            ReadingOrder =2
                                            NumeralShapes =2
                                            Left =7757
                                            Top =575
                                            Width =1080
                                            Height =300
                                            Name ="Label103"
                                            Caption ="شماره داخلي:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =3547
                                    Top =556
                                    Width =1671
                                    Height =315
                                    TabIndex =3
                                    BackColor =13434828
                                    ForeColor =8388608
                                    Name ="NUMBER"
                                    ControlSource ="NUMBER"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT NUMBER, TAG FROM HEAD_LST WHERE (TAG = 24) AND (NOT (NUMBER IN (SELECT HE"
                                        "AD_LST.NUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 25))))) ORDER BY NUMBER"
                                    AfterUpdate ="[Event Procedure]"
                                    Tag ="0"
                                    OnNotInList ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =5362
                                            Top =556
                                            Width =1335
                                            Height =300
                                            Name ="Label110"
                                            Caption ="شماره رسيد انبار:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =8670
                                    Top =945
                                    Width =4266
                                    Height =315
                                    TabIndex =4
                                    ForeColor =8388608
                                    Name ="DEPATMAN"
                                    ControlSource ="DEPATMAN"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                                    ColumnWidths ="0"
                                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                                    OnNotInList ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13086
                                            Top =945
                                            Width =750
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label107"
                                            Caption ="واحد :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =4410
                                    Top =1695
                                    Width =1191
                                    Height =315
                                    TabIndex =5
                                    ForeColor =8388608
                                    Name ="SHIFT"
                                    ControlSource ="SHIFT"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                                    ColumnWidths ="0"
                                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =5640
                                            Top =1695
                                            Width =645
                                            Height =240
                                            ForeColor =8388608
                                            Name ="Label109"
                                            Caption ="شيفت :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5640
                                    Top =945
                                    Width =1896
                                    Height =315
                                    TabIndex =6
                                    ForeColor =8388608
                                    Name ="CUST_KIND"
                                    ControlSource ="CUST_KIND"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                                        "TKNAME;"
                                    ColumnWidths ="0"
                                    ValidationRule ="Not Is Null"
                                    OnDblClick ="[Event Procedure]"
                                    DefaultValue ="[Forms]![BASEKNOW]![DEFTKH]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7575
                                            Top =945
                                            Width =1020
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label118"
                                            Caption ="نوع مشتري :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =7620
                                    Top =1305
                                    Width =5316
                                    Height =315
                                    TabIndex =7
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="CUST_NO"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                                    ColumnWidths ="0;2835;567"
                                    OnDblClick ="[Event Procedure]"
                                    Tag ="0"
                                    OnNotInList ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13080
                                            Top =1305
                                            Width =780
                                            Height =300
                                            Name ="Label19"
                                            Caption =" مشتري:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5640
                                    Top =1305
                                    Width =1926
                                    Height =315
                                    TabIndex =8
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="CUST_NO2"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes FROM CUST_HESAB"
                                    OnNotInList ="[Event Procedure]"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7635
                                    Top =1680
                                    Width =5295
                                    Height =315
                                    TabIndex =9
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="MOLAH"
                                    ControlSource ="MOLAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =13080
                                            Top =1680
                                            Width =795
                                            Height =300
                                            Name ="Label21"
                                            Caption ="ملاحظات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6390
                                    Top =1695
                                    Width =666
                                    Height =315
                                    TabIndex =10
                                    BackColor =16777164
                                    ForeColor =8388608
                                    Name ="MAS"
                                    ControlSource ="MAS"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =7110
                                            Top =1695
                                            Width =450
                                            Height =300
                                            Name ="Label106"
                                            Caption ="مدت:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =105
                                    Top =2093
                                    Width =14280
                                    Height =6225
                                    TabIndex =11
                                    Name ="INVO_LST_BRFR_SUB1"
                                    SourceObject ="Form.INVO_LST_BRFR_SUB1"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;DTAG"
                                    OnExit ="[Event Procedure]"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4857
                                    Top =386
                                    Width =291
                                    Height =315
                                    TabIndex =12
                                    Name ="hTAG"
                                    ControlSource ="htag"
                                    DefaultValue ="25"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1000
                                    Top =8447
                                    Width =2106
                                    Height =315
                                    TabIndex =13
                                    BackColor =13434828
                                    Name ="Text59"
                                    ControlSource ="=[INVO_LST_BRFR_SUB1].[Form]![SMABLK]"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =3205
                                            Top =8447
                                            Width =1260
                                            Height =300
                                            Name ="Label61"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4521
                                    Top =8451
                                    Width =1881
                                    Height =315
                                    TabIndex =14
                                    BackColor =13434828
                                    Name ="MOGU"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =6501
                                            Top =8451
                                            Width =750
                                            Height =300
                                            Name ="Label63"
                                            Caption ="موجودي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1515
                                    Top =1732
                                    Width =1341
                                    Height =300
                                    TabIndex =15
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="N_S"
                                    ControlSource ="N_S"
                                    OnDblClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2865
                                            Top =1732
                                            Width =1485
                                            Height =300
                                            BackColor =8421376
                                            ForeColor =65535
                                            Name ="lsanad"
                                            Caption ="ثبت در سند شماره:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =135
                                    Top =1417
                                    Width =2721
                                    Height =300
                                    TabIndex =16
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="MANDAH"
                                    OnDblClick ="[Event Procedure]"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2865
                                            Top =1417
                                            Width =1485
                                            Height =300
                                            BackColor =8421376
                                            ForeColor =65535
                                            Name ="Label69"
                                            Caption ="مانده حساب:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =141
                                    TextFontCharSet =0
                                    Left =150
                                    Top =907
                                    Width =1065
                                    Height =465
                                    FontWeight =700
                                    TabIndex =17
                                    Name ="Command100"
                                    Caption ="&چاپ فاكتور"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =1215
                                    Top =907
                                    Width =1290
                                    Height =465
                                    FontWeight =700
                                    TabIndex =18
                                    Name ="Command106"
                                    Caption ="فاکتور 1"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =209
                                    TextFontCharSet =0
                                    Left =2505
                                    Top =907
                                    Width =1110
                                    Height =465
                                    FontWeight =700
                                    TabIndex =19
                                    Name ="Command108"
                                    Caption =" &رسيد انبار2"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5359
                                    Top =330
                                    Width =411
                                    Height =345
                                    TabIndex =20
                                    Name ="LETSANAD"
                                    DefaultValue ="False"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =6054
                                    Top =443
                                    Width =396
                                    TabIndex =21
                                    Name ="Dtag"
                                    ControlSource ="Dtag"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1335
                                    Top =556
                                    Width =1596
                                    Height =317
                                    TabIndex =22
                                    ForeColor =8388608
                                    Name ="USER_NAME"
                                    ControlSource ="USER_NAME"
                                    DefaultValue ="UCurrentUser()"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =2988
                                            Top =556
                                            Width =525
                                            Height =300
                                            ForeColor =8388608
                                            Name ="Label111"
                                            Caption ="كاربر:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =135
                                    Top =1732
                                    Width =1371
                                    Height =300
                                    TabIndex =23
                                    BackColor =8421376
                                    ForeColor =65535
                                    Name ="MABNA"
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =3644
                                    Top =939
                                    Width =340
                                    Height =218
                                    TabIndex =24
                                    Name ="TICMBAA"
                                    ControlSource ="TICMBAA"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =4037
                                            Top =885
                                            Width =1515
                                            Height =240
                                            ForeColor =255
                                            Name ="Label142"
                                            Caption ="محاسبه.ماليات.ب.ا.ا"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =3644
                                    Top =1179
                                    Width =340
                                    Height =218
                                    TabIndex =25
                                    Name ="OKF"
                                    ControlSource ="OKF"
                                    DefaultValue ="False"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =4034
                                            Top =1125
                                            Width =975
                                            Height =240
                                            ForeColor =255
                                            Name ="Label146"
                                            Caption ="تائيد فاكتور"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =150
                                    Top =375
                                    Width =1080
                                    Height =465
                                    FontWeight =700
                                    TabIndex =26
                                    Name ="ESLAH"
                                    Caption ="اصلاح "
                                    OnClick ="[Event Procedure]"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =6616
                                    Top =9026
                                    Width =3660
                                    Height =330
                                    TabIndex =27
                                    ConditionalFormat = Begin
                                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                                        0x310000000000320000000000330000000000
                                    End
                                    Name ="PERSONEL"
                                    RowSourceType ="Value List"
                                    RowSource ="1;\"حكيميان\";2;\"آقاي فلاح\";4;\"چم حيدر\";12;\"ملكي\";125;\"محمودي\";127;\"حكي"
                                        "ميان -2\";131;\"انكوباتور\";137;\"تكنيك ها و مثالها\";140;\"كارآموز\";141;\"آبيا"
                                        "ر\";142;\"جعفري\""
                                    ColumnWidths ="0"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =10336
                                            Top =9026
                                            Width =720
                                            Height =240
                                            Name ="Label5"
                                            Caption ="ارجاع به :"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =13440
                                    Top =8442
                                    TabIndex =28
                                    Name ="SGN1"
                                    ControlSource ="SGN1"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =13881
                                            Top =8385
                                            Width =495
                                            Height =240
                                            Name ="Label149"
                                            Caption ="فروش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    ReadingOrder =1
                                    Left =13455
                                    Top =9041
                                    TabIndex =29
                                    Name ="SGN3"
                                    ControlSource ="SGN3"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =13889
                                            Top =8985
                                            Width =795
                                            Height =240
                                            Name ="Label150"
                                            Caption ="حسابداري"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =13455
                                    Top =9300
                                    TabIndex =30
                                    Name ="SGN4"
                                    ControlSource ="SGN4"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =215
                                            Left =13889
                                            Top =9300
                                            Width =780
                                            Height =240
                                            Name ="Label65"
                                            Caption ="مدير عامل"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    ReadingOrder =1
                                    Left =13455
                                    Top =8742
                                    TabIndex =31
                                    Name ="SGN2"
                                    ControlSource ="SGN2"
                                    OnClick ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =13896
                                            Top =8685
                                            Width =495
                                            Height =240
                                            Name ="Label77"
                                            Caption ="انبار"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =11540
                                    Top =8477
                                    Width =1875
                                    TabIndex =32
                                    ConditionalFormat = Begin
                                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                                        0x310000000000320000000000330000000000
                                    End
                                    Name ="SGN1usid"
                                    ControlSource ="SGN1usid"
                                    RowSourceType ="Value List"
                                    RowSource ="154;\"S.Dehestani-Tehran\";179;\"S.M.hosseinipoor\";182;\"Saeed.Zolfaghari\";81;"
                                        "\"بنفشه صادقي\";5;\"خانم برگريزان\";46;\"خانم بيات\";26;\"خانم جبينيان\";64;\"خا"
                                        "نم خوشنود\";59;\"خانم سجاد\";100;\"خانم سهند علي نژاد\";19;\"خانم کرمي\";45;\"خا"
                                        "نم کارگر\";67;\"خانم دشتي\";139;\"خانم تارا رادنيا\";74;\"خانم رفيعي\";115;\"خان"
                                        "م زکيه عرب طاهري زاده\";132;\"خانم حساني\";84;\"خانم احمدي\";69;\"خانم اميري\";4"
                                        "0;\"خانم عبدالوند\";9;\"خانم صادقي\";76;\"خانم صادقي نژاد\";87;\"خانم عليميرزايي"
                                        "\";78;\"خانم فتوحي\";110;\"خانم فاطمه جعفري\";116;\"خانم فاطمه نادريان\";3;\"خان"
                                        "م قويدل\";70;\"خانم مجاهد\";112;\"خانم مريم صمدي\";75;\"خانم محسني\";89;\"خانم م"
                                        "طوري\";135;\"خانم منصوره غريب\";125;\"خانم نرگس معين زاده\";23;\"خانم نيره قاسمي"
                                        "\";109;\"خانم نينا اميراني\";129;\"خانم هانيه ميرعلي\";108;\"خانم شريف\";126;\"خ"
                                        "انم شادي قامتي مقدم\";170;\"سميرا بهجت\";88;\"کنترل\";137;\"کنترل 2\";185;\"P.To"
                                        "rabpoor\";177;\"R.asadi\";97;\"دفترتهران\";36;\"دشتبان\";117;\"آقاي سپهر سلمانيا"
                                        "ن\";90;\"آقاي دکتر حبيب زارع\";83;\"آقاي حسيني نسب\";102;\"آقاي حميد فلاح\";120;"
                                        "\"آقاي احسان دهقان طالبي\";92;\"آقاي احمدي\";167;\"آقاي علي اکبر غفوري\";130;\"آ"
                                        "قاي عليرضا محسني\";118;\"آقاي محمد مهدي درويشي\";51;\"آقاي مطلوبي\";93;\"آقاي شر"
                                        "افت\";101;\"تهران انبار\";18;\"DrDamaki\";175;\"D.yazdani\";169;\"رسول آيتي\";10"
                                        "3;\"رضا فلاح علي آباد\";111;\"ريحانه رفيعي\";171;\"زهرا ابراهيمي\";162;\"F.Morad"
                                        "khani\";138;\"Farahmand\";134;\"اقاي سادات اخوي\";113;\"اقاي احسان احمدي\";131;\""
                                        "اقاي علي دايي\";128;\"انبار اشکذر\";187;\"H.Abdolee\";176;\"A.H.lari\";153;\"K.S"
                                        "almanian-Tehran\";164;\"Mrs.Barzegar\";160;\"Mrs.H.Hoseininasab\";159;\"Mrs.Youn"
                                        "esabadi\";183;\"Mrs.Zare\";184;\"M.Pasandeh\";161;\"M.Motiee\";178;\"M.fallah\";"
                                        "155;\"A.Abdollahpour\";157;\"jazira-S.H.mirhashemi\";172;\"jazire-mrs-sojodi\";1"
                                        "73;\"jazire-mrs-Gholampor\";158;\"jazire-mrs-gholami\";163;\"jazire-mr-ghazian\""
                                        ";186;\"m.rughangar\";165;\"mr.h.khojasteh\";166;\"ms.s.omidvar\";156;\"Anbar.sal"
                                        "hei\";181;\"علي دهقان\";30;\"علي نژاد\";180;\"عهديه قويدل\";71;\"فروشگاه يزد\";2"
                                        ";\"مدير سيستم\";136;\"مديرفروش ميبد-آقاي دهقاني\";91;\"مرخصي\";140;\"محمد حسين ش"
                                        "ايق\";141;\"مهسا تقي پور\";168;\"مهدي موسوي\";142;\"مهندس جعفري\";1;\"¬Administe"
                                        "r\""
                                    ColumnWidths ="0"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =11533
                                    Top =8755
                                    Width =1875
                                    TabIndex =33
                                    ConditionalFormat = Begin
                                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                                        0x310000000000320000000000330000000000
                                    End
                                    Name ="sgn2usid"
                                    ControlSource ="sgn2usid"
                                    RowSourceType ="Value List"
                                    RowSource ="154;\"S.Dehestani-Tehran\";179;\"S.M.hosseinipoor\";182;\"Saeed.Zolfaghari\";81;"
                                        "\"بنفشه صادقي\";77;\"خانم برگريزان\";46;\"خانم بيات\";26;\"خانم جبينيان\";64;\"خ"
                                        "انم خوشنود\";59;\"خانم سجاد\";100;\"خانم سهند علي نژاد\";19;\"خانم کرمي\";45;\"خ"
                                        "انم کارگر\";67;\"خانم دشتي\";139;\"خانم تارا رادنيا\";74;\"خانم رفيعي\";115;\"خا"
                                        "نم زکيه عرب طاهري زاده\";132;\"خانم حساني\";84;\"خانم احمدي\";69;\"خانم اميري\";"
                                        "40;\"خانم عبدالوند\";9;\"خانم صادقي\";76;\"خانم صادقي نژاد\";87;\"خانم عليميرزاي"
                                        "ي\";78;\"خانم فتوحي\";110;\"خانم فاطمه جعفري\";116;\"خانم فاطمه نادريان\";3;\"خا"
                                        "نم قويدل\";70;\"خانم مجاهد\";112;\"خانم مريم صمدي\";75;\"خانم محسني\";89;\"خانم "
                                        "مطوري\";135;\"خانم منصوره غريب\";125;\"خانم نرگس معين زاده\";23;\"خانم نيره قاسم"
                                        "ي\";109;\"خانم نينا اميراني\";129;\"خانم هانيه ميرعلي\";108;\"خانم شريف\";126;\""
                                        "خانم شادي قامتي مقدم\";170;\"سميرا بهجت\";88;\"کنترل\";137;\"کنترل 2\";185;\"P.T"
                                        "orabpoor\";177;\"R.asadi\";97;\"دفترتهران\";36;\"دشتبان\";117;\"آقاي سپهر سلماني"
                                        "ان\";90;\"آقاي دکتر حبيب زارع\";83;\"آقاي حسيني نسب\";102;\"آقاي حميد فلاح\";120"
                                        ";\"آقاي احسان دهقان طالبي\";92;\"آقاي احمدي\";167;\"آقاي علي اکبر غفوري\";130;\""
                                        "آقاي عليرضا محسني\";118;\"آقاي محمد مهدي درويشي\";51;\"آقاي مطلوبي\";93;\"آقاي ش"
                                        "رافت\";101;\"تهران انبار\";18;\"DrDamaki\";175;\"D.yazdani\";169;\"رسول آيتي\";1"
                                        "03;\"رضا فلاح علي آباد\";111;\"ريحانه رفيعي\";171;\"زهرا ابراهيمي\";162;\"F.Mora"
                                        "dkhani\";138;\"Farahmand\";134;\"اقاي سادات اخوي\";113;\"اقاي احسان احمدي\";131;"
                                        "\"اقاي علي دايي\";128;\"انبار اشکذر\";187;\"H.Abdolee\";176;\"A.H.lari\";153;\"K"
                                        ".Salmanian-Tehran\";164;\"Mrs.Barzegar\";160;\"Mrs.H.Hoseininasab\";159;\"Mrs.Yo"
                                        "unesabadi\";183;\"Mrs.Zare\";184;\"M.Pasandeh\";161;\"M.Motiee\";178;\"M.fallah\""
                                        ";155;\"A.Abdollahpour\";157;\"jazira-S.H.mirhashemi\";172;\"jazire-mrs-sojodi\";"
                                        "173;\"jazire-mrs-Gholampor\";158;\"jazire-mrs-gholami\";163;\"jazire-mr-ghazian\""
                                        ";186;\"m.rughangar\";165;\"mr.h.khojasteh\";166;\"ms.s.omidvar\";156;\"Anbar.sal"
                                        "hei\";181;\"علي دهقان\";30;\"علي نژاد\";180;\"عهديه قويدل\";71;\"فروشگاه يزد\";2"
                                        ";\"مدير سيستم\";136;\"مديرفروش ميبد-آقاي دهقاني\";91;\"مرخصي\";140;\"محمد حسين ش"
                                        "ايق\";141;\"مهسا تقي پور\";168;\"مهدي موسوي\";142;\"مهندس جعفري\";1;\"¬Administe"
                                        "r\""
                                    ColumnWidths ="0"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Visible = NotDefault
                                    Locked = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =11533
                                    Top =9040
                                    Width =1875
                                    TabIndex =34
                                    ConditionalFormat = Begin
                                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                                        0x310000000000320000000000330000000000
                                    End
                                    Name ="sgn3usid"
                                    ControlSource ="sgn3usid"
                                    RowSourceType ="Value List"
                                    RowSource ="154;\"S.Dehestani-Tehran\";179;\"S.M.hosseinipoor\";182;\"Saeed.Zolfaghari\";81;"
                                        "\"بنفشه صادقي\";77;\"خانم برگريزان\";46;\"خانم بيات\";26;\"خانم جبينيان\";64;\"خ"
                                        "انم خوشنود\";59;\"خانم سجاد\";100;\"خانم سهند علي نژاد\";19;\"خانم کرمي\";45;\"خ"
                                        "انم کارگر\";67;\"خانم دشتي\";139;\"خانم تارا رادنيا\";74;\"خانم رفيعي\";115;\"خا"
                                        "نم زکيه عرب طاهري زاده\";132;\"خانم حساني\";84;\"خانم احمدي\";69;\"خانم اميري\";"
                                        "40;\"خانم عبدالوند\";9;\"خانم صادقي\";76;\"خانم صادقي نژاد\";87;\"خانم عليميرزاي"
                                        "ي\";78;\"خانم فتوحي\";110;\"خانم فاطمه جعفري\";116;\"خانم فاطمه نادريان\";3;\"خا"
                                        "نم قويدل\";70;\"خانم مجاهد\";112;\"خانم مريم صمدي\";75;\"خانم محسني\";89;\"خانم "
                                        "مطوري\";135;\"خانم منصوره غريب\";125;\"خانم نرگس معين زاده\";23;\"خانم نيره قاسم"
                                        "ي\";109;\"خانم نينا اميراني\";129;\"خانم هانيه ميرعلي\";108;\"خانم شريف\";126;\""
                                        "خانم شادي قامتي مقدم\";170;\"سميرا بهجت\";88;\"کنترل\";137;\"کنترل 2\";185;\"P.T"
                                        "orabpoor\";177;\"R.asadi\";97;\"دفترتهران\";36;\"دشتبان\";117;\"آقاي سپهر سلماني"
                                        "ان\";90;\"آقاي دکتر حبيب زارع\";83;\"آقاي حسيني نسب\";102;\"آقاي حميد فلاح\";120"
                                        ";\"آقاي احسان دهقان طالبي\";92;\"آقاي احمدي\";167;\"آقاي علي اکبر غفوري\";130;\""
                                        "آقاي عليرضا محسني\";118;\"آقاي محمد مهدي درويشي\";51;\"آقاي مطلوبي\";93;\"آقاي ش"
                                        "رافت\";101;\"تهران انبار\";18;\"DrDamaki\";175;\"D.yazdani\";169;\"رسول آيتي\";1"
                                        "03;\"رضا فلاح علي آباد\";111;\"ريحانه رفيعي\";171;\"زهرا ابراهيمي\";162;\"F.Mora"
                                        "dkhani\";138;\"Farahmand\";134;\"اقاي سادات اخوي\";113;\"اقاي احسان احمدي\";131;"
                                        "\"اقاي علي دايي\";128;\"انبار اشکذر\";187;\"H.Abdolee\";176;\"A.H.lari\";153;\"K"
                                        ".Salmanian-Tehran\";164;\"Mrs.Barzegar\";160;\"Mrs.H.Hoseininasab\";159;\"Mrs.Yo"
                                        "unesabadi\";183;\"Mrs.Zare\";184;\"M.Pasandeh\";161;\"M.Motiee\";178;\"M.fallah\""
                                        ";155;\"A.Abdollahpour\";157;\"jazira-S.H.mirhashemi\";172;\"jazire-mrs-sojodi\";"
                                        "173;\"jazire-mrs-Gholampor\";158;\"jazire-mrs-gholami\";163;\"jazire-mr-ghazian\""
                                        ";186;\"m.rughangar\";165;\"mr.h.khojasteh\";166;\"ms.s.omidvar\";156;\"Anbar.sal"
                                        "hei\";181;\"علي دهقان\";30;\"علي نژاد\";180;\"عهديه قويدل\";71;\"فروشگاه يزد\";2"
                                        ";\"مدير سيستم\";136;\"مديرفروش ميبد-آقاي دهقاني\";91;\"مرخصي\";140;\"محمد حسين ش"
                                        "ايق\";141;\"مهسا تقي پور\";168;\"مهدي موسوي\";142;\"مهندس جعفري\";1;\"¬Administe"
                                        "r\""
                                    ColumnWidths ="0"
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    IMESentenceMode =3
                                    Left =1936
                                    Top =9000
                                    TabIndex =35
                                    Name ="UID"
                                    ControlSource ="UID"
                                    DefaultValue ="=UCurrentUserIDD()"
                                End
                            End
                        End
                        Begin Page
                            Enabled = NotDefault
                            OverlapFlags =247
                            AccessKey =129
                            Left =75
                            Top =450
                            Width =14880
                            Height =9090
                            Name ="Page58"
                            Caption ="&پشت فاكتور"
                            Begin
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11392
                                    Top =1125
                                    Height =315
                                    Name ="M_NAGHD"
                                    ControlSource ="M_NAGHD"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13236
                                            Top =1125
                                            Width =750
                                            Height =300
                                            Name ="Label23"
                                            Caption ="مبلغ نقد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11390
                                    Top =1485
                                    Height =315
                                    TabIndex =1
                                    Name ="TAKHFIF"
                                    ControlSource ="TAKHFIF"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13228
                                            Top =1500
                                            Width =990
                                            Height =300
                                            Name ="Label37"
                                            Caption ="مبلغ تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =9150
                                    Top =1470
                                    Width =1101
                                    Height =315
                                    TabIndex =2
                                    Name ="Text101"
                                    Format ="Fixed"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10296
                                            Top =1485
                                            Width =1020
                                            Height =300
                                            Name ="Label102"
                                            Caption ="درصد تخفيف:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11390
                                    Top =1890
                                    Height =315
                                    TabIndex =3
                                    Name ="MABL_HAZ"
                                    ControlSource ="MABL_HAZ"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13228
                                            Top =1890
                                            Width =555
                                            Height =300
                                            Name ="Label33"
                                            Caption ="هزينه:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =3
                                    Left =5355
                                    Top =1890
                                    Width =4896
                                    Height =315
                                    TabIndex =4
                                    Name ="MOIN_HAZ"
                                    ControlSource ="MOIN_HAZ"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)"
                                        ") + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS n"
                                        "varchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS "
                                        "nvarchar)) AS Expr2 FROM TDETA_HES"
                                    ColumnWidths ="0;1990;851"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10296
                                            Top =1890
                                            Width =990
                                            Height =300
                                            Name ="Label35"
                                            Caption ="معين هزينه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =9585
                                    Top =501
                                    Width =1221
                                    Height =300
                                    TabIndex =5
                                    Name ="Text73"
                                    ControlSource ="DATE_N"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =10815
                                            Top =495
                                            Width =1065
                                            Height =300
                                            Name ="Label74"
                                            Caption ="تاريخ فاكتور:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =1575
                                    Top =501
                                    Width =7236
                                    Height =300
                                    TabIndex =6
                                    Name ="Combo75"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                                    ColumnWidths ="0;2835;567"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8820
                                            Top =495
                                            Width =780
                                            Height =300
                                            Name ="Label76"
                                            Caption ="مشتري:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =90
                                    Top =450
                                    Width =14130
                                    Height =371
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =5085
                                    Top =960
                                    Width =9135
                                    Height =1931
                                    Name ="Box81"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =11909
                                    Top =501
                                    Width =1116
                                    Height =300
                                    TabIndex =7
                                    Name ="Text78"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13034
                                            Top =499
                                            Width =1110
                                            Height =300
                                            Name ="Label79"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =120
                                    Top =3210
                                    Width =14145
                                    Height =5295
                                    TabIndex =8
                                    Name ="PAY_GETP_SUB"
                                    SourceObject ="Form.PAY_GETP_SUB_BRFR"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;DTAG"
                                    OnExit ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =13005
                                            Top =2940
                                            Width =1215
                                            Height =300
                                            Name ="PAY_GETD_SUB Label"
                                            Caption ="چكهاي پرداختي"
                                            EventProcPrefix ="PAY_GETD_SUB_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =225
                                    Top =2370
                                    Width =2661
                                    Height =300
                                    TabIndex =9
                                    Name ="NPAR"
                                    ControlSource ="=[M_NAGHD]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2986
                                            Top =2377
                                            Width =1545
                                            Height =300
                                            Name ="Label86"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =226
                                    Top =1050
                                    Width =2661
                                    Height =300
                                    TabIndex =10
                                    Name ="JF"
                                    ControlSource ="=[INVO_LST_BRFR_SUB1].[Form]![SMABLK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2991
                                            Top =1050
                                            Width =1260
                                            Height =300
                                            Name ="Label88"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =226
                                    Top =1380
                                    Width =2661
                                    Height =300
                                    TabIndex =11
                                    Name ="HKH"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2985
                                            Top =1380
                                            Width =1080
                                            Height =300
                                            Name ="Label92"
                                            Caption ="هزينه خدمات:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =75
                                    Top =960
                                    Width =4737
                                    Height =2154
                                    Name ="Box95"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =226
                                    Top =1714
                                    Width =2661
                                    Height =300
                                    TabIndex =12
                                    Name ="NTKHFIF"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2993
                                            Top =1714
                                            Width =735
                                            Height =300
                                            Name ="Label94"
                                            Caption ="تخفيفات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =231
                                    Top =2044
                                    Width =2661
                                    Height =300
                                    TabIndex =13
                                    BackColor =16777164
                                    Name ="GHABEL"
                                    ControlSource ="=[JF]+[HKH]+[MBAA]-[NTKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2991
                                            Top =2051
                                            Width =1440
                                            Height =300
                                            Name ="Label97"
                                            Caption ="مبلغ قابل پرداخت:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =225
                                    Top =2700
                                    Width =2661
                                    Height =300
                                    TabIndex =14
                                    BackColor =16777164
                                    Name ="MAN"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2986
                                            Top =2707
                                            Width =525
                                            Height =300
                                            Name ="Label99"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1623
                                    Top =8617
                                    Width =2106
                                    Height =315
                                    TabIndex =15
                                    Name ="NCHK"
                                    ControlSource ="=IIf(IsNull([PAY_GETP_SUB].[Form]![SCMAB]),0,[PAY_GETP_SUB].[Form]![SCMAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3822
                                            Top =8617
                                            Width =1260
                                            Height =300
                                            Name ="Label90"
                                            Caption ="جمع مبلغ چكها:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =11389
                                    Top =2327
                                    Height =315
                                    TabIndex =16
                                    Name ="MBAA"
                                    ControlSource ="MBAA"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    AfterUpdate ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13230
                                            Top =2325
                                            Width =945
                                            Height =240
                                            Name ="Label115"
                                            Caption ="ماليات ب.ا.ا:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =5370
                                    Top =2325
                                    Width =4896
                                    Height =315
                                    TabIndex =17
                                    Name ="HMBAA"
                                    ControlSource ="HMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarc"
                                        "har)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvar"
                                        "char)) AS Expr2 FROM TDETA_HES"
                                    ColumnWidths ="851"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =10311
                                            Top =2325
                                            Width =990
                                            Height =300
                                            Name ="Label117"
                                            Caption ="معين ماليات:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =105
                            Top =330
                            Width =14850
                            Height =9210
                            Name ="Page155"
                            Caption ="ساير"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =270
                                    Top =1135
                                    Width =14295
                                    Height =2700
                                    Name ="VISITOR_DTL_SUB"
                                    SourceObject ="Form.VISITOR_DTL_SUBBACKFR"
                                    LinkChildFields ="NUMBER;TAG"
                                    LinkMasterFields ="NUMBER;dTAG"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =105
                                    Top =330
                                    Width =14580
                                    Height =371
                                    Name ="Box170"
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =12239
                                    Top =381
                                    Width =1116
                                    Height =300
                                    TabIndex =1
                                    Name ="Text171"
                                    ControlSource ="NUMBER"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =13364
                                            Top =379
                                            Width =1110
                                            Height =300
                                            Name ="Label172"
                                            Caption ="شماره فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    Left =10005
                                    Top =381
                                    Width =1131
                                    Height =300
                                    TabIndex =2
                                    Name ="Text173"
                                    ControlSource ="DATE_N"
                                    InputMask ="####/##/##"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =11145
                                            Top =375
                                            Width =1065
                                            Height =300
                                            Name ="Label174"
                                            Caption ="تاريخ فاكتور:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    BackStyle =0
                                    NumeralShapes =2
                                    ColumnCount =2
                                    Left =270
                                    Top =381
                                    Width =8871
                                    Height =300
                                    TabIndex =3
                                    Name ="Combo175"
                                    ControlSource ="CUST_NO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                                    ColumnWidths ="0;2835;567"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9150
                                            Top =375
                                            Width =795
                                            Height =300
                                            Name ="Label176"
                                            Caption =" مشتري:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =105
                                    Top =903
                                    Width =14568
                                    Height =3510
                                    Name ="Box182"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =12510
                                    Top =795
                                    Width =1245
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="VISITOR_DTL_SUB Label"
                                    Caption ="پورسانت"
                                    EventProcPrefix ="VISITOR_DTL_SUB_Label"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =364
                                    Top =5210
                                    Width =4047
                                    Height =2154
                                    Name ="Box177"
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =471
                                    Top =6613
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =4
                                    Name ="Text178"
                                    ControlSource ="=[M_NAGHD]+[MABL_VAR]+[MABL_HAV]+[NCHK]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2677
                                            Top =6620
                                            Width =1545
                                            Height =300
                                            Name ="Label179"
                                            Caption ="جمع مبالغ پرداختي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =472
                                    Top =5293
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =5
                                    Name ="Text180"
                                    ControlSource ="=IIf(IsError([INVO_LST_BRF_SUB].[Form]![SMABLK]) Or IsNull([INVO_LST_BRF_SUB].[F"
                                        "orm]![SMABLK]),0,[INVO_LST_BRF_SUB].[Form]![SMABLK])"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2677
                                            Top =5293
                                            Width =1260
                                            Height =300
                                            Name ="Label181"
                                            Caption ="جمع كل فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =471
                                    Top =5623
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =6
                                    Name ="Text182"
                                    ControlSource ="=[MABL_HAZ]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2676
                                            Top =5623
                                            Width =1080
                                            Height =300
                                            Name ="Label183"
                                            Caption ="هزينه خدمات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =471
                                    Top =5953
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =7
                                    Name ="Text184"
                                    ControlSource ="=[TAKHFIF]"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2676
                                            Top =5953
                                            Width =735
                                            Height =300
                                            Name ="Label185"
                                            Caption ="تخفيفات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =471
                                    Top =6283
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =8
                                    BackColor =16777164
                                    Name ="Text186"
                                    ControlSource ="=IIf([VAS]=1 Or IsNull([VAS]),[JF]+[HKH]+[MBAA]-[NTKHFIF],[JF]-[HKH]+[MBAA]-[NTK"
                                        "HFIF])"
                                    Format ="#,##0.00;-#,##0.00"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2676
                                            Top =6290
                                            Width =1440
                                            Height =300
                                            Name ="Label187"
                                            Caption ="مبلغ قابل پرداخت:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    TextAlign =3
                                    TextFontFamily =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =471
                                    Top =6943
                                    Width =2106
                                    Height =300
                                    FontWeight =700
                                    TabIndex =9
                                    BackColor =16777164
                                    Name ="Text188"
                                    ControlSource ="=[GHABEL]-[NPAR]"
                                    Format ="#,##0.00;#,##0.00-"
                                    FontName ="Titr"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2677
                                            Top =6950
                                            Width =525
                                            Height =300
                                            Name ="Label189"
                                            Caption ="مانده:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =499
                                    Top =3965
                                    Width =2106
                                    Height =315
                                    TabIndex =10
                                    BackColor =13408767
                                    Name ="Text190"
                                    ControlSource ="=IIf(IsNull([VISITOR_DTL_SUB].[Form]![SCMAB]) Or IsError([VISITOR_DTL_SUB].[Form"
                                        "]![SCMAB]),0,[VISITOR_DTL_SUB].[Form]![SCMAB])"
                                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2705
                                            Top =3965
                                            Width =1260
                                            Height =300
                                            Name ="Label191"
                                            Caption ="جمع پورسانت:"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_LST_BRFR.cls"
