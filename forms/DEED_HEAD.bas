Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15141
    RowHeight =300
    ItemSuffix =16
    Left =4125
    Top =15
    Right =19260
    Bottom =9705
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xc5da682f53d7e440
    End
    RecordSource ="DEED_HEAD_Q"
    Caption ="سند حسابداري"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnGotFocus ="[Event Procedure]"
    OnLostFocus ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            CanGrow = NotDefault
            Height =9694
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12361
                    Top =105
                    Height =315
                    ColumnWidth =1005
                    ColumnOrder =0
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="N_S"
                    ControlSource ="N_S"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14150
                            Top =112
                            Width =990
                            Height =300
                            Name ="Label1"
                            Caption ="شماره سند:"
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
                    Left =12361
                    Top =445
                    Height =315
                    ColumnWidth =1020
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    ValidationRule ="Is Not Null And >10101"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14150
                            Top =442
                            Width =825
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ سند:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3276
                    Top =800
                    Width =10785
                    Height =315
                    ColumnWidth =6390
                    ColumnOrder =11
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14149
                            Top =802
                            Width =885
                            Height =300
                            Name ="Label5"
                            Caption ="شرح  سند:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =302
                    Top =349
                    Width =366
                    Height =300
                    ColumnWidth =825
                    ColumnOrder =10
                    TabIndex =3
                    ForeColor =128
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =681
                            Top =356
                            Width =705
                            Height =300
                            ForeColor =128
                            Name ="Label7"
                            Caption ="نوع سند:"
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =117
                    Top =1190
                    Width =14898
                    Height =7190
                    TabIndex =4
                    Name ="Child14"
                    SourceObject ="Form.DEED_DTL_sub"
                    LinkChildFields ="N_S"
                    LinkMasterFields ="N_S"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    Left =282
                    Top =60
                    Width =1806
                    Height =300
                    ColumnOrder =9
                    TabIndex =5
                    ForeColor =128
                    Name ="وضعيت سند"
                    ControlSource ="=IIf([GHATEI]<>0,\"سند قطعي ميباشد\",\"سند قطعي نيست\")"
                    EventProcPrefix ="وضعيت_سند"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7937
                    Top =283
                    Width =951
                    ColumnWidth =735
                    ColumnOrder =8
                    TabIndex =6
                    Name ="GHATEI"
                    ControlSource ="GHATEI"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2130
                    Top =8850
                    Height =315
                    ColumnWidth =900
                    ColumnOrder =7
                    TabIndex =7
                    BackColor =255
                    ForeColor =16777215
                    Name ="SSBED"
                    ControlSource ="=[Child14].[Form]![SBED]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =8844
                            Width =1335
                            Height =240
                            Name ="Label6"
                            Caption ="جمع سند:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =405
                    Top =8850
                    Height =315
                    ColumnOrder =6
                    TabIndex =8
                    BackColor =255
                    ForeColor =16777215
                    Name ="SSBES"
                    ControlSource ="=[Child14].[Form]![SBES]"
                    Format ="#,###"
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =141
                    TextFontCharSet =0
                    Left =2266
                    Top =60
                    Width =960
                    Height =465
                    FontWeight =700
                    TabIndex =9
                    Name ="Command22"
                    Caption ="&چاپ سند"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9255
                    Top =60
                    Height =315
                    ColumnWidth =990
                    ColumnOrder =2
                    TabIndex =10
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="BASE"
                    ControlSource ="BASE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11040
                            Top =60
                            Width =915
                            Height =300
                            Name ="Label24"
                            Caption ="شماره مبنا:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6966
                    Top =56
                    Width =351
                    Height =255
                    ColumnOrder =5
                    TabIndex =11
                    Name ="DE"
                    DefaultValue ="False"
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5438
                    Top =8505
                    Width =8574
                    Height =300
                    ColumnOrder =4
                    TabIndex =12
                    BackColor =8421376
                    Name ="Text8"
                    ControlSource ="=getkolname([Child14].[Form]![HES_K]) & \" - \" & GETMOINNAME([Child14].[Form]!["
                        "HES_K],[Child14].[Form]![HES_M]) & \"... \" & GETHESNAME([Child14].[Form]![hes])"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14085
                            Top =8505
                            Width =990
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label9"
                            Caption ="شرح حساب:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5438
                    Top =8850
                    Width =8575
                    Height =300
                    ColumnOrder =12
                    TabIndex =13
                    BackColor =8421376
                    Name ="Text10"
                    ControlSource ="=[Child14].[Form]![SHARH]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14078
                            Top =8850
                            Width =990
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label11"
                            Caption ="شرح رديف:"
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =225
                    Top =741
                    Width =1686
                    Height =345
                    ColumnWidth =1740
                    ColumnOrder =3
                    TabIndex =14
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =1935
                            Top =741
                            Width =465
                            Height =300
                            Name ="Label2"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4247
                    Top =226
                    Width =340
                    Height =218
                    TabIndex =15
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4699
                            Top =158
                            Width =795
                            Height =345
                            BackColor =6723891
                            Name ="prnl"
                            Caption ="تائيد سند"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =5685
                    Top =60
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =16
                    Name ="ESLAH"
                    Caption ="اصلاح سند"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =3235
                    Top =56
                    Width =960
                    Height =465
                    FontWeight =700
                    TabIndex =17
                    Name ="Command3"
                    Caption ="&چاپ سند2"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2130
                    Top =8430
                    Height =315
                    TabIndex =18
                    BackColor =10092543
                    ForeColor =8388608
                    Name ="bedt"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =8460
                            Width =1335
                            Height =240
                            Name ="Label33"
                            Caption ="جمع تا سطر سند:"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =405
                    Top =8430
                    Height =315
                    TabIndex =19
                    BackColor =10092543
                    ForeColor =8388608
                    Name ="best"
                    Format ="#,###"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =14001
                    Top =9297
                    Height =332
                    TabIndex =20
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =14256
                            Top =9297
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
                    Left =10654
                    Top =9297
                    Width =290
                    Height =332
                    TabIndex =21
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =10938
                            Top =9297
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
                    Left =7575
                    Top =9297
                    Height =227
                    TabIndex =22
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7829
                            Top =9297
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
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =210
                    Top =9240
                    Width =3660
                    Height =332
                    TabIndex =23
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="20;\"Negahban\";14;\"paydar\";19;\"shakeri\";11;\"رنجبر\";7;\"alireza\";18;\"mir"
                        "tabaei\";9;\"milad\";5;\"عليرضا\";2;\"مدير سيستم\";4;\"مدير سيستم نجمه معقول\";3"
                        ";\"نسرين معقول\";1;\"¬Administer\";10;\"شهاب\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3930
                            Top =9240
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
                    OverlapFlags =119
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =12120
                    Top =9297
                    Width =1875
                    TabIndex =24
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="SGN1usid"
                    ControlSource ="sgn1usid"
                    RowSourceType ="Value List"
                    RowSource ="20;\"Negahban\";14;\"paydar\";19;\"shakeri\";11;\"رنجبر\";7;\"alireza\";18;\"mir"
                        "tabaei\";9;\"milad\";5;\"عليرضا\";2;\"مدير سيستم\";4;\"مدير سيستم نجمه معقول\";3"
                        ";\"نسرين معقول\";1;\"¬Administer\";10;\"شهاب\""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8844
                    Top =9297
                    Width =1875
                    TabIndex =25
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    RowSourceType ="Value List"
                    RowSource ="20;\"Negahban\";14;\"paydar\";19;\"shakeri\";11;\"رنجبر\";7;\"alireza\";18;\"mir"
                        "tabaei\";9;\"milad\";5;\"عليرضا\";2;\"مدير سيستم\";4;\"مدير سيستم نجمه معقول\";3"
                        ";\"نسرين معقول\";1;\"¬Administer\";10;\"شهاب\""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5725
                    Top =9297
                    Width =1875
                    TabIndex =26
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    RowSourceType ="Value List"
                    RowSource ="20;\"Negahban\";14;\"paydar\";19;\"shakeri\";11;\"رنجبر\";7;\"alireza\";18;\"mir"
                        "tabaei\";9;\"milad\";5;\"عليرضا\";2;\"مدير سيستم\";4;\"مدير سيستم نجمه معقول\";3"
                        ";\"نسرين معقول\";1;\"¬Administer\";10;\"شهاب\""
                    ColumnWidths ="0"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1303
                    Top =453
                    ColumnWidth =2070
                    TabIndex =27
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4275
                            Top =450
                            Width =735
                            Height =240
                            Name ="Label10"
                            Caption ="زمان ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1303
                    Top =793
                    TabIndex =28
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4305
                            Top =795
                            Width =585
                            Height =240
                            Name ="Label12"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9269
                    Top =450
                    Height =315
                    TabIndex =29
                    BackColor =10092543
                    ForeColor =8388608
                    Name ="BAYEG"
                    ControlSource ="BAYEG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11055
                            Top =450
                            Width =1110
                            Height =240
                            Name ="Label15"
                            Caption ="شماره بايگاني:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DEED_HEAD.cls"
