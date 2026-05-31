Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14910
    DatasheetFontHeight =10
    ItemSuffix =133
    Left =6930
    Top =1335
    Right =18720
    Bottom =8025
    HelpContextId =6003
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Filter ="((pish_view.CUST_NO=\"128-3-1\"))"
    RecSrcDt = Begin
        0x5e0172178e16e640
    End
    RecordSource ="pish_view"
    Caption ="پيش فاكتور فروش"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnResize ="[Event Procedure]"
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
            Height =9850
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11460
                    Top =90
                    Width =1056
                    Height =315
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =12660
                            Top =90
                            Width =585
                            Height =240
                            Name ="Label9"
                            Caption ="*تاريخ :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =7965
                    Top =484
                    Width =5886
                    Height =315
                    TabIndex =1
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;3969;567"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13995
                            Top =484
                            Width =870
                            Height =300
                            Name ="Label19"
                            Caption =" مشتري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6015
                    Top =495
                    Width =1866
                    Height =315
                    TabIndex =2
                    Name ="CUST_NO2"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7215
                    Top =840
                    Width =6630
                    Height =315
                    TabIndex =3
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13995
                            Top =840
                            Width =795
                            Height =300
                            Name ="Label21"
                            Caption ="ملاحظات:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =56
                    Top =1814
                    Width =14730
                    Height =5280
                    TabIndex =4
                    Name ="INVO_LST_PISHFROOSH_SUB"
                    SourceObject ="Form.INVO_LST_PISHFROOSH_SUB2"
                    LinkChildFields ="NUMBER;TAG"
                    LinkMasterFields ="NUMBER;HTAG"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13215
                    Top =120
                    Width =636
                    Height =315
                    TabIndex =6
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="0"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13995
                            Top =121
                            Width =630
                            Height =240
                            Name ="Label1"
                            Caption ="شماره :"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4918
                    Top =233
                    Width =291
                    Height =255
                    TabIndex =7
                    Name ="hTAG"
                    ControlSource ="TAG"
                    DefaultValue ="20"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =7143
                    Width =2106
                    Height =315
                    TabIndex =8
                    Name ="Text59"
                    ControlSource ="=[INVO_LST_PISHFROOSH_sub].[Form]![SMABLK]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2375
                            Top =7143
                            Width =1110
                            Height =240
                            Name ="Label61"
                            Caption ="جمع كل فاكتور:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =7155
                    Width =1131
                    Height =315
                    TabIndex =9
                    Name ="MOGU"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4740
                            Top =7155
                            Width =705
                            Height =240
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
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1360
                    Width =3381
                    Height =300
                    TabIndex =10
                    BackColor =32768
                    ForeColor =65535
                    Name ="MANDAH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =3503
                            Top =1360
                            Width =975
                            Height =300
                            BackColor =32768
                            ForeColor =65535
                            Name ="lsanad"
                            Caption ="مانده حساب:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =1635
                    Top =90
                    Width =825
                    Height =405
                    FontWeight =700
                    TabIndex =11
                    Name ="Command100"
                    Caption ="&چاپ"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7710
                    Top =7155
                    Width =1461
                    Height =315
                    TabIndex =5
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9209
                            Top =7170
                            Width =990
                            Height =300
                            Name ="Label37"
                            Caption ="مبلغ تخفيف:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6015
                    Top =93
                    Width =2151
                    Height =315
                    TabIndex =12
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8223
                            Top =90
                            Width =525
                            Height =300
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="*كاربر:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =10290
                    Top =1245
                    Width =3576
                    Height =330
                    TabIndex =13
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14010
                            Top =1245
                            Width =495
                            Height =240
                            ForeColor =8388608
                            Name ="Label107"
                            Caption ="واحد :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =12756
                    Width =1191
                    Height =315
                    TabIndex =14
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
                            Left =13986
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label109"
                            Caption ="شيفت :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8790
                    Top =105
                    Width =1566
                    Height =315
                    TabIndex =15
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                        "TKNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10395
                            Top =105
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label113"
                            Caption ="نوع مشتري :"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7710
                    Top =7550
                    Width =1461
                    Height =315
                    TabIndex =16
                    Name ="MABL_HAZ"
                    ControlSource ="MABL_HAZ"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9220
                            Top =7567
                            Width =600
                            Height =240
                            Name ="Label105"
                            Caption ="خدمات:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    EnterKeyBehavior = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =7995
                    Width =9066
                    Height =990
                    TabIndex =17
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    OnDblClick ="[Event Procedure]"
                    Tag ="با دابل کليک در متن شرح فاکتور قبلي انتقال داده ميشود"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9224
                            Top =8010
                            Width =810
                            Height =240
                            Name ="Label108"
                            Caption ="متن پائين :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5505
                    Top =7155
                    Width =1191
                    Height =315
                    TabIndex =18
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6747
                            Top =7175
                            Width =900
                            Height =240
                            Name ="Label110"
                            Caption ="ماليات.ب.ا.ا"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =4500
                    Top =114
                    Width =340
                    Height =218
                    TabIndex =19
                    Name ="TICMBAA"
                    ControlSource ="TICMBAA"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4893
                            Top =60
                            Width =1515
                            Height =240
                            ForeColor =255
                            Name ="Label142"
                            Caption ="ارزش افزوده"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2880
                    Top =90
                    Width =420
                    Height =405
                    FontWeight =700
                    TabIndex =20
                    ForeColor =255
                    Name ="custprint"
                    Caption ="H"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =105
                    Top =540
                    Width =1380
                    Height =405
                    TabIndex =21
                    Name ="Command113"
                    Caption ="تبديل به فاكتور"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =4560
                    Top =714
                    Width =280
                    Height =218
                    TabIndex =22
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4890
                            Top =660
                            Width =810
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تبديل شده"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =2040
                    Top =1190
                    Width =441
                    TabIndex =23
                    Name ="CDDATE"
                    ControlSource ="CDDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1500
                    Top =505
                    Width =441
                    TabIndex =24
                    Name ="CDTIME"
                    ControlSource ="CDTIME"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1500
                    Top =845
                    Width =441
                    TabIndex =25
                    Name ="OKDATE"
                    ControlSource ="OKDATE"
                    DefaultValue ="Now()"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =1500
                    Top =1185
                    Width =441
                    TabIndex =26
                    Name ="OKTIME"
                    ControlSource ="OKTIME"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =0
                    Left =3767
                    Top =86
                    Width =825
                    Height =405
                    FontWeight =700
                    TabIndex =27
                    Name ="Command139"
                    Caption ="چاپ 1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2460
                    Top =90
                    Width =420
                    Height =405
                    FontWeight =700
                    TabIndex =28
                    ForeColor =255
                    Name ="Command114"
                    Caption ="O"
                    OnClick ="[Event Procedure]"
                    Tag ="چاپ سفارش"
                End
                Begin CheckBox
                    OverlapFlags =255
                    ReadingOrder =1
                    Left =4500
                    Top =414
                    Width =340
                    Height =218
                    TabIndex =29
                    Name ="JAY"
                    ControlSource ="JAY"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4893
                            Top =360
                            Width =1020
                            Height =240
                            ForeColor =255
                            Name ="JAYl"
                            Caption ="محاسبه جايزه"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3315
                    Top =90
                    Width =420
                    Height =405
                    FontWeight =700
                    TabIndex =30
                    ForeColor =255
                    Name ="cccc"
                    Caption ="C"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =106
                    Top =45
                    Width =1380
                    Height =465
                    TabIndex =31
                    Name ="Command116"
                    Caption ="تبديل به حواله"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2607
                    Top =602
                    Width =1986
                    Height =405
                    TabIndex =32
                    BackColor =65535
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    RowSourceType ="Value List"
                    RowSource ="0;\"عادي\";1;\"رزرو شده\";2;\"تبديل به حواله\";3;\"لغو شد\";4;\"رزو قطعي\""
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    DefaultValue ="0"
                    OnChange ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =105
                    Top =960
                    Width =1380
                    Height =405
                    TabIndex =33
                    Name ="Command118"
                    Caption ="تبديل به سفارش"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =12699
                    Top =8390
                    Width =1956
                    Height =591
                    TabIndex =34
                    Name ="Command119"
                    Caption ="ارجاع به مدير"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000082000000270000000100180000000000ba3b0000232e0000232e0000 ,
                        0x0000000000000000cbac12d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11cbac120000dda911e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810dda9110001d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa110505 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a80fe6a507e7a70ce7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a70de7a70defc257ea ,
                        0xb127e6a508e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a70de7a6 ,
                        0x0cf1ce76fffefbf9e7bce9ae22e6a508e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a80ee7a913f4d58cfffefefffffffffffff9e7bbe9ae22e6a508e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a70ce9ae1ffaeccafffffffffffffffffffffffffaebc6 ,
                        0xeab22be6a507e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a60beab22af9e7bdffffffff ,
                        0xfffffffffffffffffaeccaeab22ae6a507e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810efc881e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a5 ,
                        0x07e9ae20f9e7bcfffffffffffffffffffffffffaeccaeab22ae6a507e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810eac179eac179e9bd69e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a80fe7a810e7a810e7a810e7a810e7a810e7a810e7a810eb ,
                        0xd0a2e8b245e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a811e6a509e6a60ae7a811e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810ecbb5ef6dfbafdfaf4fffffffcf5e9 ,
                        0xf2d39fe9af31e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a811e7a80ee7a810e7a810e7a810e7a810e7a810eec171fffffffa ,
                        0xefdeecbb5ee7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e6a508e9ae21f9e7bdfffffffffffffffffffffffffaeb ,
                        0xc9eab129e6a406e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810eac687eee6dbefefefefefefefefefedddc5e9b858e7a810e7a810 ,
                        0xe7a810e9b858e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e9b858efefefeeebe5ebcb95e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a508ebb639eab3 ,
                        0x2fe6a609e7a810e7a810e7a810e7a810e7a810e7a810efc881ffffffffffffff ,
                        0xfffffcf5e9fffffffffffffaefdee9af31e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a70de7a811e7a810e7a810e7a810e7a810e7a8 ,
                        0x10eab54af9ead2fffffffdfaf4efc881e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e6a508e9ae21f9e7bdffffff ,
                        0xfffffffffffffffffffaecc9ebb533e6a506e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e9b858efefefeee6dbebcb95eac179ecd4aeefefefee ,
                        0xe6dbe7a810e7a810eac687efefefeac687e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e9bd69ede2d0efefefebcb95e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a60a ,
                        0xeab32ffefaf3fcf3e1e9ad1fe7a70ce7a810e7a810e7a810e7a810e9af31fdfa ,
                        0xf4fffffff2d39fe7a810e7a810eab54afaefdefffffff6dfbae7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a70ee7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810f2d39fffffffffffffecbb5ee7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6 ,
                        0xa508e9af22f8e3b2fffffffffffffffffffffffffbf1d8ebb739e6a506e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810ecd4aeeeebe5e8b245e7a810e7a8 ,
                        0x10e7a810ebcb95efefefebcb95e7a810e8ad2fecd9bae8ad2fe7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e8ad2fede2d0 ,
                        0xefefefe9b858e7a810e7a70ee7a80fe7a810e7a810e7a810e7a810e7a80fe7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a70ee6a509e8aa16f8e5baf6dea8e7a912e7a70ee7a810e7a810e7a810 ,
                        0xe7a80feec170fffffff2d39fe7a810e7a810e7a810e7a810e9af31fdfaf4fdfa ,
                        0xf4e7a810e7a810e7a810e7a810e7a810e7a810e7a80fe7a70de7a80fe7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810f1ce90ffffff ,
                        0xfcf5e9e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e6a609e7a914f6dfa6fffffffffffffffffffffffffc ,
                        0xf2ddebb739e6a406e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810ede2d0ede2d0 ,
                        0xe7a810e7a810e7a810e7a810e8ad2fefefefebd0a2e9bd69eac179e8b245e7a8 ,
                        0x10e7a810e7a810e7a810e7a810eac179eac179eac179e8ad2fe7a810e7a810e7 ,
                        0xa810e7a810e9b858efefefecd9bae7a810e7a912e7a80feac179eac179eac179 ,
                        0xe9bd69e7a80fe7a70ee7a811e7a810e7a810e7a810e7a810e7a810e7a810e7a7 ,
                        0x0ee6a60be7a60be6a406e7a913e8ad1fe8ac1be9ad20e7a60be6a403e7a811e7 ,
                        0xa810e7a810e7a810e7a70ef2d39fffffffeec170e7a810e7a810e7a810e7a810 ,
                        0xe7a810f9ead2ffffffe7a810e7a810e7a810e7a810eab54aecbb5eeab54ae7a9 ,
                        0x12e7a810eec170eec170e7a810e7a810eec170eec170e9af31e7a810e7a810e7 ,
                        0xa810e7a810f9ead2ffffffeec170e7a810eab54ae7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a60ae7aa16f6dfa5ffff ,
                        0xfffffffffffffffffffffcf1d9ebb635e6a404e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810ecd9baeee6dbe7a810e7a810e7a810e7a810e8ad2fefefefefefefede2d0 ,
                        0xefefefefefefecd9bae8ad2fe7a810e9b858eeebe5efefefefefefefefefefef ,
                        0xefecd9bae8ad2fe7a810e7a810e7a810ede2d0ede2d0e7a810e9ba5ceeebe5ef ,
                        0xefefefefefefefefefefefede2d0e9b859e6a70ce7a810e7a810e7a810e7a810 ,
                        0xe7a810e6a405e7a811e9af25e8ac1deec157fbefd5fefaf3fefbf5fdf7e9f8e4 ,
                        0xb8edbc49e6a507e7a810e7a810e7a810e7a70eeec170fffffff2d39fe7a810e7 ,
                        0xa810e7a810e7a810e7a810f9ead2ffffffe7a810e7a810e9af31f9ead2ffffff ,
                        0xf9ead2fffffff6dfbaf9ead2fffffffffffffaefdef6dfbaffffffffffffffff ,
                        0xfff1ce90e7a810e7a810e7a810f2d39ffffffff2d39fe7a810fffffffaefdee7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe6a60ae8ab17f6dfa7fffffffffffffffffffffffffbf0d4ecb739e6a406e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810ebcb95efefefeac179e7a810e7a810e7a810e8ad2fef ,
                        0xefefeeebe5ebcb95ebd0a2ecd9baefefefebcb95e7a810edddc5efefefecd9ba ,
                        0xebd0a2ebd0a2ecd9baefefefebcb95e7a810e7a810e8ad2fefefefedddc5e7a8 ,
                        0x10eddec5efefefecd9baebd0a2ebd0a2ecd9baefefefeddec5e7a60be7a810e7 ,
                        0xa810e7a810e7a810e7a509efc563fbefd6fffdfbfefbf5fffefcfffffffbf1d9 ,
                        0xfaeed2fcf2ddfffffffdf9eee9b025e7a60be7a810e7a810e7a810eab549ffff ,
                        0xfff9ead2e7a810e7a810e7a810e7a810e7a810f9ead2ffffffe7a810e7a810f2 ,
                        0xd39ffffffffcf5e9f4d9adfcf5e9fffffffffffffaefdef9ead2ffffffffffff ,
                        0xfcf5e9f4d9adfdfaf4fdfaf4e7a810e7a810e7a810f2d39ffffffff1ce90e7a8 ,
                        0x10ffffffffffffe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e6a60ae8aa16f7e1acffffffffffffffffffffffff ,
                        0xfcf3ddecba42e6a508e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e8ad2febcb95e8ad2fe7a810e7a8 ,
                        0x10e7a810e8ad2fefefefebd0a2e7a810e7a810e7a810eee6dbebd0a2e7a810ed ,
                        0xe2d0ede2d0e7a810e7a810e7a810e7a810eee6dbede2d0e7a810e7a810eac687 ,
                        0xefefefebcb95e7a810ede2d0ede2d0e7a810e7a810e7a810e7a810ecd4aeefef ,
                        0xefe8ad2ee7a70ee7a810e7a810e6a609eab433fffefdfdf8edf9e9c5fbf2dbff ,
                        0xfffff7e0ace6a508e7a70ee7a509f2d183fffffff0c86be6a303e7a810e7a810 ,
                        0xe7a810e7a80ff9ead2f6dfbae7a810e7a810e7a810e7a810e7a810f9ead2ffff ,
                        0xffe7a810e7a810f2d39fffffffeec170e7a810e9af31fcf5e9ffffffefc881e7 ,
                        0xa810fcf5e9ffffffefc881e7a810f2d39fffffffeec170e7a810e7a810faefde ,
                        0xffffffecbb5ee7a810ffffffffffffe7a80fe7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a60ae7a911f6dea3ff ,
                        0xfffffffffffffffffffffffdf7e7ebb637e6a60ae7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e8b245efefefebd0a2e7a810e7a810e9b858efef ,
                        0xefebd0a2e7a810ede2d0ede2d0e7a810e7a810e7a810e8b245efefefebd0a2e7 ,
                        0xa810e7a810eac687eee6dbe9bd69e7a810ede2d0ede2d0e7a810e7a810e7a810 ,
                        0xe7a810ecd9baeeebe5e7a70fe7a80fe7a810e7a810e6a507ecb941fffffff3d2 ,
                        0x86e39800f0c766fffffff1cb73e6a200e7a80fe6a200f0c96dffffffefc768e6 ,
                        0xa303e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810f9ead2ffffffe7a810e7a810f2d39fffffffeec170e7a810e7a810f1ce ,
                        0x90fffffff2d39fe7a810f2d39ffffffff2d39fe7a810f6dfbaffffffeab54ae7 ,
                        0xa810eab54afffffffcf5e9e7a810e7a810ffffffffffffe7a80fe7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a5 ,
                        0x07eab32df9e8c0fffffffffffffffffffffffffaeac4eab127e7a60ce7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810eac179efefefebd0a2e7a810 ,
                        0xe7a810ebcb95efefefe9bd69e7a810ede2d0ede2d0e7a810e7a810e7a810ebcb ,
                        0x95efefefeac687e7a810e7a810e7a810e7a810e7a810e7a810ede2d0ede2d0e7 ,
                        0xa810e7a810e7a810e9bd69efefefecd9bae7a60be7a810e7a810e7a810e7a70f ,
                        0xe7a80ff9e8c1fffffff3d183f5da9affffffeec053e6a405e7a80fe7a60bfae8 ,
                        0xc2ffffffebb534e6a609e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810f9ead2ffffffe7a810e7a810f2d39fffffffeec170 ,
                        0xe7a810e7a810e9af31fffffff7e4c6e7a810ecbb5efffffff4d9ade7a810fdfa ,
                        0xf4fdfaf4e7a810e7a810e7a810efc881f1ce90e7a810e7a810ffffffffffffe7 ,
                        0xa70fe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e6a506ebb532fbefd3fffffffffffffffffffffffff8e5b5e9af22e6a5 ,
                        0x09e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810eac179ef ,
                        0xefefebd0a2e7a810e7a810e8b245ebd0a2e7a810e7a810ede2d0ede2d0e7a810 ,
                        0xe7a810e7a810e8b245ecd4aee7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10ede2d0ede2d0e7a810e7a810e9b858eeebe5eeebe5e8ad2fe7a70ee7a810e7 ,
                        0xa810e7a810e7a810e7a60ae8ad1ef8e4b7fffffffffffffffefdebb83ce6a507 ,
                        0xe7a60beab028fefaf3fbefd3e7a70de7a80fe7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810faefdeffffffe7a810e7a810f2 ,
                        0xd39fffffffeec170e7a810e7a810e7a810e9af31e7a810e7a810e7a810eab54a ,
                        0xe7a810f1ce90fffffff6dfbae7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10ffffffffffffe7a70ee7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e6a506ebb431fbefd2fffffffffffffffffffffffff7e2ad ,
                        0xe8ab18e6a609e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810eac179efefefeac179e7a810e7a810e7a810e7a810e7a810e7a810ed ,
                        0xe2d0ede2d0e7a811e7a811e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810ede2d0ede2d0e7a810e9b858eeebe5eeebe5e9b858e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a811e6a60ae6a50aecbc4afdf9efff ,
                        0xffffe9b028e6a60ae7a810e7a70febb638ecb93fe6a60ae7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a80fe7a810e7a810e7a810e7a810e7a810ffffffffff ,
                        0xffe7a810e7a810f2d39fffffffeec170e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810f1ce90eab54ae7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810ffffffffffffe7a70ee7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e6a405ebb634fbf0d5ffffffffffffffffffff ,
                        0xfffff7e1abe8aa16e6a609e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810eac179efefefeac179e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810eee6dbede2d0e7a70ee9b657eee6dbe9bd69ecd4aeecd4aee7 ,
                        0xa810e7a811e7a810e7a810e7a810e7a810eee6dbede2d0e7a810eac179efefef ,
                        0xedddc5e9bd69e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a7 ,
                        0x0de6a405f4d998f9eac8e8ac1ce7a70de7a810e7a810e6a508e6a508e7a811e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a70ce7a810e7a810e7a810e7a810 ,
                        0xe7a810ffffffffffffe7a810e7a810f7e4c6ffffffeec170e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810ffffffffffffe7a80fe7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e6a506ebb532fbefd1ffffffffff ,
                        0xfffffffffffffff8e4b3e8ac1be6a609e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810eac179efefefeac179e7a810 ,
                        0xe7a810e7a810e7a810e7a810e8ad2fefefefede2d0e7a810e9b858eee6dbe9bd ,
                        0x69ecd4aeebd0a2e7a80fe7a913e7a914e7a810e7a810e8ad2fefefefede2d0e7 ,
                        0xa810e8ad2fecd9baefefefefefefedddc5ebcb95e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e6a508e7a812e7a811e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a70fe7a810e7 ,
                        0xa810e7a810e7a810e7a810fffffff9ead2e7a810e7a810f9ead2ffffffeec170 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810fffffffcf5e9e7 ,
                        0xa80fe7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a405ecba40fbf2da ,
                        0xfffffffffffffffffffffffff8e3b2e8ad1ce6a609e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e9bd69ed ,
                        0xe2d0e9bd69e7a810e7a810e7a810e7a810e7a810e7a810ecd4aeebcb95e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a70de6a60be7a810e7a810e7a8 ,
                        0x10ecd4aeebcb95e7a810e7a810e7a810e9b858ebd0a2eeebe5efefefeac179e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a70ee7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a80fe7a810e7a810e7a810e7a810fffffff9ead2e7a810e7a810f9 ,
                        0xead2ffffffeab54ae7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810eec1 ,
                        0x70fffffff9ead2e7a80fe7a810e7a810e7a810e7a810e7a810e7a810e6a405ec ,
                        0xba3ffcf4dffffffffffffffffffffffffff6dda1e8ab18e6a609e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e8ad ,
                        0x2febcb95e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810ecbb5efffffff9ea ,
                        0xd2e7a810e7a810faefdeffffffe7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810ecbb5efffffff9ead2e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e6a405ecbb44fcf4e0fffffffffffffffffffffffff5db9be7a912e7a60ae7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810eab54aecbb5ee7a810e7a810e7a810eec170e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e9af31ecbb5ee7a810e7a810e7a810 ,
                        0xe7a810e7a810e6a404ecb93ffcf4e0fffffffffffffffffffffffff6dda1e8a9 ,
                        0x11e7a60be7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a70deebf4ffcf4dfffffffffffffffffffffffff ,
                        0xf6dea2e8a913e7a60be7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a70ce9ae21fcf1d8ffffffffffffff ,
                        0xfffffffffff6db9ae8aa14e6a60ae7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a80fe7a70deec2 ,
                        0x55fcf6e6fffffffffffff4d488e7a60be6a60ce7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a80fe6a405edbe4bfcf4e1f5d892e7a70ee7a70ce7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a80fe6a60aeab42fe8aa16e7a70ce7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810d9aa112020d9aa11e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e6a60ae7 ,
                        0xa70ee7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810d9aa117674 ,
                        0xd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810d9aa11446fd9aa11e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810d9aa112020dda911e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7 ,
                        0xa810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810 ,
                        0xe7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a810e7a8 ,
                        0x10e7a810e7a810e7a810e7a810e7a810e7a810dda9113037cbac12d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa ,
                        0x11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9 ,
                        0xaa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11 ,
                        0xd9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11d9aa11cbac12672f ,
                        0x0000
                    End
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    Picture ="C:\\Users\\Administrator\\Desktop\\Buttons\\Forwardtoadmin.bmp"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =12705
                    Top =9000
                    Width =1956
                    Height =591
                    TabIndex =35
                    Name ="Command120"
                    Caption ="پاسخ مدير "
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000082000000270000000100180000000000ba3b0000232e0000232e0000 ,
                        0x000000000000000041bc0830be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0641bc0800002bbf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf062bbf06000130be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be060505 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0642c42c29bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf062bbf0729bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0688d67fc9ebc5eaf7 ,
                        0xe9fffffff5fbf4e0f3dea4de9d42c42c29bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf06a4de9da4de9d29bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf065cc32690c43c29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf062dbf0729bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06bde7b9 ,
                        0xfffffffffffffffffffffffffffffffffffffffffff5fbf4a4de9d29bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf062abf072abf0729bf0629bf06e0f3dee0f3de29bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0658c11cc1d77ebce29829bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0697da8efffffff5fbf4b1e3ab69cd5b42c42c42c42c79d16ebde7b9ffffff ,
                        0x97da8e29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf062cc00933c21246c52fd5efd288d67f88 ,
                        0xd67fd5efd242c42c29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0670ce4af6f5e3cfecba44be0e37be0929 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf06e0f3deffffffa4de9d29bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0656c84629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf062ebf09a0da ,
                        0x91ffffffbde7b9bde7b9ffffff97da8e29bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf067ecf5058c72c29bf0629bf064cc82a ,
                        0xacdc86adde89acde88acde88acde88acde88abde87addf8ed0eec1fffffff0f9 ,
                        0xe9b7dc89acd6752cbf0629bf062fbf0829bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0642c42cffffffffffff42c42c29bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0688d67f29bf0629bf0688d67f29bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0645bd0ad6e8b2b9d9824a ,
                        0xc01429bf0667d14cfdf8edfef9f0fdf9effdf9effdf9effdf9effdf9effdfaf1 ,
                        0xfefdfafffffffffffffffffffbf0d729bf0629bf0630bf0929bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06f5fbf4ffffff ,
                        0x56c84629bf0629bf0629bf0629bf0629bf0629bf0629bf0669cd5ba4de9d79d1 ,
                        0x6e29bf0629bf0679d16e88d67f42c42c29bf0679d16e88d67f69cd5b29bf0629 ,
                        0xbf0629bf0629bf0629bf0679d16ea4de9d89d67f7cd16f29bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0698c8 ,
                        0x49ecf5dfedf1d182cb4629bf0656cb36fffffffffefdfffefdfffefdfffefdff ,
                        0xfefdfffefefffefcfffffffffffffffffffffffffefbf529bf0629bf062abf06 ,
                        0x2abf0729bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf06c9ebc5ffffffbde7b929bf0629bf0629bf0629bf0629bf0697da8eeaf7e9 ,
                        0xffffffffffffffffffe0f3debde7b9fffffffffffff5fbf4b1e3abffffffffff ,
                        0xffffffffbde7b929bf0629bf0669cd5beaf7e9fffffffffffffffffffffffff5 ,
                        0xfbf4b2e3ac2abf0729bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x47c0145dc72edde7b7fffbf7ffffffb3e3985ec83183ce4efdf3dffbf1d9fbf1 ,
                        0xd9fbf1d9fbf1d9fbf1d9fbf1dafbf0d7fcf3e0fefbf4fffffffffefdfdf7e929 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0642c42cf5fbf4ffffffd5efd288d67f29bf0629bf0669 ,
                        0xcd5bffffffffffffe0f3deffffffe0f3defffffffffffff5fbf4c9ebc5ffffff ,
                        0xfffffff5fbf4bde7b9eaf7e9ffffff79d16e29bf06d5efd2ffffffe0f3deffff ,
                        0xffe0f3dee0f3def5fbf4ffffff97da8e29bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf06accd5de9e0a4fffbf4fffffffffffffdf7e9f6e6b9f2e4b3 ,
                        0xfaebc8faeac5faeac5faeac5faeac5faeac5faeac6f9e9c3f7e4b5f9eccaffff ,
                        0xfffffefdfdf7ea29bf0629bf0629bf062abf0729bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0669cd5be0f3deffffffffff ,
                        0xffeaf7e9bde7b9d5efd2ffffffbde7b942c42c29bf0629bf06c9ebc5ffffffbd ,
                        0xe7b929bf06c9ebc5ffffffbde7b929bf0688d67fffffffb1e3ab29bf06e0f3de ,
                        0xffffff42c42c29bf0629bf0629bf0679d16effffffbde7b929bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf06fffcf7ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfbf2d9f8e6bbfffffdfffefdfdf7ea29bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be060a2030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf06a4de9de0f3defffffffffffffffffffffffffffffff5fbf429bf0629bf ,
                        0x0669cd5bffffffe0f3de29bf0679d16effffffe0f3de29bf06a4de9dffffffa4 ,
                        0xde9d29bf06e0f3deffffff42c42c29bf0629bf0629bf0679d16effffffb1e3ab ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06fdf7eafffefcffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffcf2dbf9e8c2fffffffffffffefbf429bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0642c42c29bf0629bf0656c84688d67fc9ebc5f5fbf4ffffff ,
                        0xffffff79d16e29bf0629bf06e0f3def5fbf429bf0629bf06f5fbf4e0f3de29bf ,
                        0x06d5efd2ffffff79d16e29bf06e0f3deffffff42c42c29bf0629bf0629bf06b1 ,
                        0xe3abffffffa4de9d29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0xfdf7e9fffefcffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffaf2d9ffe5bbfff8f1fcf4e1faeac529 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0656c846e0f3dee0f3de69cd5b56c846a4de9deaf7e9ff ,
                        0xfffff5fbf4b1e3ab79d16e29bf0629bf0629bf0656c84629bf0629bf0629bf06 ,
                        0x69cd5b42c42c69cd5bffffffeaf7e929bf0629bf06e0f3deffffff42c42c29bf ,
                        0x0629bf0629bf06eaf7e9ffffff69cd5b29bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf06fdf7e9fffefcffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffefedc783c63a44c6 ,
                        0x2360cf465fc93529bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf06e0f3deffffffffffffffff ,
                        0xffffffffffffffbde7b942c42c29bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0688d67f88d67f29bf0629bf06e0f3de ,
                        0xffffff42c42c29bf0629bf0629bf0669cd5ba4de9d29bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf06fdf7e9fffefcffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xe9eabf60be1928bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf062abf0636be0929bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0642c42c ,
                        0xb1e3abe0f3dee0f3dec9ebc588d67f29bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf06f5fbf4ffffff42c42c29bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06fdf7e8fffefcffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffe9eac069c02328bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0656c84629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0642c42cffffffffffff42c42c29bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0xfdfaf0fffffeffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffefebc56fc12829bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf06bde7b9f5fbf469cd5b29bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0642c42cffffffeaf7e929bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf06fbf5e1fefcf6fefdfafefdf9fefdf9fefdf9fefdf9fefdf9 ,
                        0xfefdf9fefdf9fefdf9fefdf9ffffffffffffffffffffffffd6e5aa51c01729bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06bde7b9f5fb ,
                        0xf469cd5b29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0656c846ffffff ,
                        0xe0f3de29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf06d9dd97deedc2dfefc9dfeec7dfeec7df ,
                        0xeec7dfeec7dfeec7dfeec7dfeec7dfeec7dfefc9cae9b1c5e8acc3e8acd6efc9 ,
                        0x77ce4b29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0656c84629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0656c84679d16e29bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0628bf062abf0729bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be062020 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be06202030be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06202030be0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0630be06202030be0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0630be063330 ,
                        0x30be0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0630be064f7230be0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0630be06703a2bbf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629 ,
                        0xbf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf06 ,
                        0x29bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf0629bf ,
                        0x0629bf0629bf0629bf0629bf0629bf0629bf062bbf06343a41bc0830be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be ,
                        0x0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be0630 ,
                        0xbe0630be0630be0630be0630be0630be0630be0630be0630be0630be0630be06 ,
                        0x30be0630be0630be0630be0630be0630be0630be0630be0630be0641bc08312f ,
                        0x0000
                    End
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    Picture ="C:\\Users\\Administrator\\Desktop\\Buttons\\answertoadmin.bmp"
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13635
                    Top =7212
                    TabIndex =36
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13896
                            Top =7155
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =13605
                    Top =7466
                    Width =290
                    TabIndex =37
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13889
                            Top =7410
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13635
                    Top =7725
                    TabIndex =38
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13889
                            Top =7662
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4680
                    Top =1200
                    Width =1395
                    Height =405
                    FontWeight =700
                    TabIndex =39
                    Name ="okpish"
                    Caption ="تاييد پيش فاکتور"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =6088
                    Top =1202
                    Width =1395
                    Height =405
                    FontWeight =700
                    TabIndex =40
                    Name ="estelam"
                    Caption ="بررسي وضعيت"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =7530
                    Top =1260
                    Width =1716
                    Height =330
                    TabIndex =41
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="MODAT_PPID"
                    ControlSource ="MODAT_PPID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PPID, PPAME, MODAT FROM PRICE_PAYNO UNION SELECT 0, 'آزاد', 0"
                    ColumnWidths ="0;1077;285"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9300
                            Top =1260
                            Width =975
                            Height =285
                            ForeColor =8388608
                            Name ="Label106"
                            Caption ="نحوه پرداخت:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =6053
                    Top =9070
                    Width =5166
                    Height =330
                    TabIndex =42
                    ForeColor =8388608
                    Name ="PEPID"
                    ControlSource ="PEPID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PEPID, PEPNAME, PEPDATE, PEPDEPART FROM PRICE_ELAMIE ORDER BY PEPNAME DES"
                        "C"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11363
                            Top =9070
                            Width =1125
                            Height =330
                            ForeColor =8388608
                            Name ="PEPIDl"
                            Caption ="اعلاميه قيمت :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =165
                    Top =9090
                    Width =4596
                    Height =330
                    TabIndex =43
                    ForeColor =8388608
                    Name ="PEID"
                    ControlSource ="PEID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PEID, PENAME, PEDATE, PEPDEPART FROM PRICE_ELAMIETF"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4853
                            Top =9095
                            Width =1170
                            Height =285
                            ForeColor =8388608
                            Name ="PEIDl"
                            Caption ="اعلاميه تخفيف :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5895
                    Top =855
                    Width =771
                    Height =317
                    TabIndex =44
                    ForeColor =8388608
                    Name ="MAS"
                    ControlSource ="MAS"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6705
                            Top =855
                            Width =450
                            Height =300
                            ForeColor =8388608
                            Name ="Label129"
                            Caption ="مدت:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    Left =1635
                    Top =585
                    Width =735
                    Height =405
                    FontWeight =700
                    TabIndex =45
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3907
                    Top =7485
                    Width =420
                    Height =405
                    FontWeight =700
                    TabIndex =46
                    ForeColor =255
                    Name ="Command130"
                    Caption ="HA"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3458
                    Top =7483
                    Width =420
                    Height =405
                    FontWeight =700
                    TabIndex =47
                    ForeColor =255
                    Name ="Command131"
                    Caption ="FA"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10080
                    Top =8055
                    Width =3540
                    Height =330
                    TabIndex =48
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
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13680
                            Top =8055
                            Width =720
                            Height =240
                            Name ="Label5"
                            Caption ="ارجاع به :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11731
                    Top =7200
                    Width =1875
                    TabIndex =49
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
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11730
                    Top =7478
                    Width =1875
                    TabIndex =50
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11730
                    Top =7763
                    Width =1875
                    TabIndex =51
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
                    OverlapFlags =247
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1870
                    Top =1133
                    TabIndex =52
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=UCurrentUserIDD()"
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_LST_PISHFROOSH2.cls"
