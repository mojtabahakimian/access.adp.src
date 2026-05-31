Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15363
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =4140
    Right =19500
    Bottom =10005
    HelpContextId =5003
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="PGET_HED.N_S"
    RecSrcDt = Begin
        0x2413040f135ee640
    End
    RecordSource ="PGET_HED"
    Caption ="خزانه داري"
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
        Begin Line
            Width =1701
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
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Tab
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
            Height =11316
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13188
                    Top =86
                    Width =1101
                    Height =315
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="DATE"
                    ControlSource ="DATE"
                    ValidationRule ="Is Not Null And >10101"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14392
                            Top =86
                            Width =585
                            Height =300
                            Name ="Label1"
                            Caption ="تاريخ :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9234
                    Top =460
                    Width =5055
                    Height =315
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14392
                            Top =457
                            Width =840
                            Height =300
                            Name ="Label3"
                            Caption ="ملاحظات :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13
                    Top =469
                    Width =1281
                    Height =300
                    TabIndex =2
                    BackColor =8421376
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1307
                            Top =471
                            Width =975
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label5"
                            Caption ="شماره سند :"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Top =1140
                    Width =15343
                    Height =7921
                    TabIndex =3
                    Name ="PGET_LST_SUB"
                    SourceObject ="Form.PGET_LST_SUB"
                    LinkChildFields ="ID;DATE"
                    LinkMasterFields ="ID;DATE"
                    OnExit ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =141
                    TextFontCharSet =0
                    Left =13
                    Top =60
                    Width =2295
                    Height =390
                    FontWeight =700
                    TabIndex =4
                    ForeColor =16711680
                    Name ="Command12"
                    Caption ="&چاپ  عملكرد خزانه"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =11098
                    Top =825
                    Width =1525
                    Height =309
                    TabIndex =5
                    BackColor =65535
                    Name ="Frame15"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextFontCharSet =0
                            TextAlign =2
                            Left =11309
                            Top =833
                            Width =1035
                            Height =330
                            BackColor =65535
                            Name ="Label16"
                            Caption ="بدهكار "
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =9418
                    Top =830
                    Width =1645
                    Height =309
                    TabIndex =6
                    BackColor =10079487
                    Name ="Frame17"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =2
                            Left =9794
                            Top =841
                            Width =900
                            Height =270
                            BackColor =10079487
                            Name ="Label18"
                            Caption ="بستانكار "
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
                    Left =11343
                    Top =90
                    Width =741
                    Height =315
                    TabIndex =7
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="ID"
                    ControlSource ="ID"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12132
                            Top =90
                            Width =990
                            Height =240
                            Name ="Label20"
                            Caption ="شماره خزانه:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5022
                    Top =86
                    Width =1026
                    TabIndex =8
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =93
                    AccessKey =70
                    TextFontCharSet =0
                    Left =2323
                    Top =60
                    Width =1185
                    Height =391
                    FontWeight =700
                    TabIndex =9
                    ForeColor =16711680
                    Name ="Command23"
                    Caption =" س&ند دريافت"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1606
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =51
                    TextFontCharSet =0
                    Left =2323
                    Top =469
                    Width =1185
                    Height =391
                    FontWeight =700
                    TabIndex =10
                    ForeColor =16711680
                    Name ="Command24"
                    Caption ="&سند پرداخت"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1587
                End
                Begin ComboBox
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8808
                    Top =90
                    Width =1656
                    Height =330
                    TabIndex =11
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="0;\"عادي\";2;\"سند دريافت\";3;\"سند پرداخت\""
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10573
                            Top =90
                            Width =690
                            Height =315
                            Name ="Label26"
                            Caption ="نوع برگه:"
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
                    Left =7068
                    Top =90
                    Width =696
                    Height =315
                    TabIndex =12
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="IDK"
                    ControlSource ="IDK"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7812
                            Top =90
                            Width =915
                            Height =240
                            Name ="Label31"
                            Caption ="شماره برگه:"
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13
                    Top =754
                    Width =1281
                    Height =300
                    TabIndex =13
                    BackColor =8421376
                    ForeColor =65535
                    Name ="MABNA"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =1307
                            Top =756
                            Width =975
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label33"
                            Caption ="شماره مبنا :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4807
                    Top =86
                    Width =1686
                    Height =345
                    TabIndex =14
                    BackColor =13434828
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6517
                            Top =86
                            Width =465
                            Height =300
                            Name ="Label2"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    ReadingOrder =1
                    Left =3448
                    Top =129
                    Width =340
                    Height =218
                    TabIndex =15
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3838
                            Top =75
                            Width =765
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد خزانه"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3562
                    Top =358
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =16
                    Name ="ESLAH"
                    Caption ="اصلاح خزانه"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =6673
                    Top =483
                    Width =1716
                    Height =285
                    TabIndex =17
                    ForeColor =8388608
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8473
                            Top =483
                            Width =495
                            Height =270
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =4813
                    Top =483
                    Width =1191
                    Height =285
                    TabIndex =18
                    ForeColor =8388608
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =6043
                            Top =483
                            Width =645
                            Height =270
                            ForeColor =8388608
                            Name ="Label109"
                            Caption ="شيفت :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =190
                    Top =9111
                    Width =1761
                    Height =315
                    TabIndex =19
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MABL"
                    ControlSource ="=[PGET_LST_SUB].[Form]![Text26]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1995
                            Top =9111
                            Width =915
                            Height =300
                            Name ="Label19"
                            Caption ="جمع مبالغ:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5974
                    Top =9111
                    Width =8394
                    Height =300
                    TabIndex =20
                    BackColor =13434828
                    ForeColor =16711680
                    Name ="Text10"
                    ControlSource ="=[PGET_LST_SUB].[Form]![THES]"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14374
                            Top =9111
                            Width =825
                            Height =300
                            BackColor =65535
                            Name ="Label11"
                            Caption ="به حساب:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5974
                    Top =9461
                    Width =8394
                    Height =300
                    TabIndex =21
                    BackColor =13434828
                    ForeColor =16711680
                    Name ="Text8"
                    ControlSource ="=[PGET_LST_SUB].[Form]![FHES]"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14381
                            Top =9461
                            Width =825
                            Height =300
                            BackColor =10079487
                            Name ="Label9"
                            Caption ="از حساب:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3010
                    Top =9461
                    Width =2796
                    Height =300
                    TabIndex =22
                    BackColor =13408767
                    Name ="MANDB"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3005
                    Top =9111
                    Width =2796
                    Height =300
                    TabIndex =23
                    BackColor =13408767
                    Name ="MANDS"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =14021
                    Top =9861
                    Height =332
                    TabIndex =24
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14276
                            Top =9861
                            Width =885
                            Height =332
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="تنظيم كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =10674
                    Top =9861
                    Width =290
                    Height =332
                    TabIndex =25
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10958
                            Top =9861
                            Width =975
                            Height =332
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="مدير مالي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =7595
                    Top =9861
                    Height =227
                    TabIndex =26
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =7849
                            Top =9861
                            Width =960
                            Height =332
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
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
                    Left =20
                    Top =9861
                    Width =3660
                    Height =332
                    TabIndex =27
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="78;\"Controller\";106;\"ShirzadForoosh\";71;\"shirzad\";87;\"خانم دهقان-آزمايشگا"
                        "ه\";90;\"آباديس\";86;\"آقاي سجاد راستي\";42;\"آقاي دامكي\";95;\"آقاي حبيبي\";72;"
                        "\"توليد\";73;\"اقاي  رحيمي\";69;\"اقاي مظفري\";105;\"MsAsari\";104;\"f.heydari\""
                        ";101;\"kontorol2\";107;\"mdehghan\";100;\"mehrniya\";102;\"mina mehrnia\";2;\"مد"
                        "ير سيستم\";4;\"مدير عامل\";68;\"مدير شركت\";1;\"¬Administer\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3740
                            Top =9861
                            Width =720
                            Height =332
                            Name ="LPERSONEL"
                            Caption ="ارجاع به :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
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
                    Left =12140
                    Top =9861
                    Width =1875
                    TabIndex =28
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="SGN1usid"
                    ControlSource ="sgn1usid"
                    RowSourceType ="Value List"
                    RowSource ="78;\"Controller\";106;\"ShirzadForoosh\";71;\"shirzad\";87;\"خانم دهقان-آزمايشگا"
                        "ه\";90;\"آباديس\";86;\"آقاي سجاد راستي\";42;\"آقاي دامكي\";95;\"آقاي حبيبي\";72;"
                        "\"توليد\";73;\"اقاي  رحيمي\";69;\"اقاي مظفري\";105;\"MsAsari\";104;\"f.heydari\""
                        ";101;\"kontorol2\";107;\"mdehghan\";100;\"mehrniya\";102;\"mina mehrnia\";2;\"مد"
                        "ير سيستم\";4;\"مدير عامل\";68;\"مدير شركت\";1;\"¬Administer\""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8864
                    Top =9861
                    Width =1875
                    TabIndex =29
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    RowSourceType ="Value List"
                    RowSource ="78;\"Controller\";106;\"ShirzadForoosh\";71;\"shirzad\";87;\"خانم دهقان-آزمايشگا"
                        "ه\";90;\"آباديس\";86;\"آقاي سجاد راستي\";42;\"آقاي دامكي\";95;\"آقاي حبيبي\";72;"
                        "\"توليد\";73;\"اقاي  رحيمي\";69;\"اقاي مظفري\";105;\"MsAsari\";104;\"f.heydari\""
                        ";101;\"kontorol2\";107;\"mdehghan\";100;\"mehrniya\";102;\"mina mehrnia\";2;\"مد"
                        "ير سيستم\";4;\"مدير عامل\";68;\"مدير شركت\";1;\"¬Administer\""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5745
                    Top =9861
                    Width =1875
                    TabIndex =30
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    RowSourceType ="Value List"
                    RowSource ="78;\"Controller\";106;\"ShirzadForoosh\";71;\"shirzad\";87;\"خانم دهقان-آزمايشگا"
                        "ه\";90;\"آباديس\";86;\"آقاي سجاد راستي\";42;\"آقاي دامكي\";95;\"آقاي حبيبي\";72;"
                        "\"توليد\";73;\"اقاي  رحيمي\";69;\"اقاي مظفري\";105;\"MsAsari\";104;\"f.heydari\""
                        ";101;\"kontorol2\";107;\"mdehghan\";100;\"mehrniya\";102;\"mina mehrnia\";2;\"مد"
                        "ير سيستم\";4;\"مدير عامل\";68;\"مدير شركت\";1;\"¬Administer\""
                    ColumnWidths ="0"
                End
            End
        End
    End
End
CodeBehindForm
' See "PGET_HED.cls"
