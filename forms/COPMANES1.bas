Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9123
    RowHeight =285
    ItemSuffix =51
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x7f6b76de71c9e540
    End
    RecordSource ="COPMANES"
    Caption ="مشتريان"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
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
        Begin Section
            CanGrow = NotDefault
            Height =14225
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =57
                    Width =2490
                    ColumnWidth =1395
                    Name ="COMPANY_NAME"
                    ControlSource ="COMPANY_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =57
                            Width =780
                            Height =240
                            Name ="IS1"
                            Caption ="نام شركت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =397
                    Width =2490
                    TabIndex =1
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =397
                            Width =405
                            Height =240
                            Name ="IS2"
                            Caption ="شهر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =737
                    Width =2490
                    TabIndex =2
                    Name ="MANAGER"
                    ControlSource ="MANAGER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =737
                            Width =780
                            Height =240
                            Name ="IS3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =1077
                    Width =2490
                    TabIndex =3
                    Name ="FACT_TEL"
                    ControlSource ="FACT_TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1077
                            Width =900
                            Height =240
                            Name ="IS4"
                            Caption ="تلفن كارخانه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =1418
                    Width =2490
                    TabIndex =4
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1418
                            Width =1245
                            Height =240
                            Name ="IS5"
                            Caption ="موبايل مدير عامل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =1758
                    TabIndex =5
                    Name ="PERNUM"
                    ControlSource ="PERNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1758
                            Width =930
                            Height =240
                            Name ="IS6"
                            Caption ="تعداد پرسنل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =2438
                    Width =2490
                    TabIndex =6
                    Name ="PRODUCTS"
                    ControlSource ="PRODUCTS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2438
                            Width =720
                            Height =240
                            Name ="IS8"
                            Caption ="محصولات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =2778
                    Width =2490
                    TabIndex =7
                    Name ="ADDR"
                    ControlSource ="ADDR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2778
                            Width =450
                            Height =240
                            Name ="IS9"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3118
                    Width =2490
                    TabIndex =8
                    Name ="ACCOUNTANT"
                    ControlSource ="ACCOUNTANT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =3118
                            Width =645
                            Height =240
                            Name ="IS10"
                            Caption ="حسابدار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3458
                    Width =2490
                    TabIndex =9
                    Name ="SOFTWARE"
                    ControlSource ="SOFTWARE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =3458
                            Width =645
                            Height =240
                            Name ="IS11"
                            Caption ="نرم افزار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3799
                    Width =2490
                    TabIndex =10
                    Name ="ESP_PERSON"
                    ControlSource ="ESP_PERSON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =3799
                            Width =1230
                            Height =240
                            Name ="IS12"
                            Caption ="اشخاص ذي نفوذ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =4139
                    Width =2490
                    TabIndex =11
                    Name ="REAGENT"
                    ControlSource ="REAGENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =4139
                            Width =480
                            Height =240
                            Name ="IS13"
                            Caption ="معرف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =4819
                    Width =2490
                    TabIndex =12
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =4819
                            Width =825
                            Height =240
                            Name ="IS15"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1814
                    Top =4479
                    TabIndex =13
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000000000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x31000000310030000000
                    End
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"شروع ارتباط\";2;\"درحال مذاكره\";3;\"قرار ملاقات\";4;\"ملاقات شده \";5;\"در "
                        "شرف قرار داد\";6;\"راكد\";7;\"قرارداد بسته شد\";8;\"روي ميز\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =4479
                            Width =585
                            Height =240
                            Name ="IS14"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1867
                    Top =2098
                    Width =2490
                    TabIndex =14
                    Name ="STATUS_FACT"
                    ControlSource ="STATUS_FACT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT COPMANES.STATUS_FACT FROM COPMANES GROUP BY COPMANES.STATUS_FACT; "
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =53
                            Top =2098
                            Width =1095
                            Height =240
                            Name ="IS7"
                            Caption ="وضعيت كارخانه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4758
                    Top =2948
                    TabIndex =15
                    Name ="date_sabt"
                    ControlSource ="date_sabt"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7143
                            Top =2955
                            Width =915
                            Height =405
                            Name ="Label31"
                            Caption ="تاريخ و ساعت ثبت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =623
                    Top =5953
                    TabIndex =16
                    Name ="id"
                    ControlSource ="id"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3772
                            Top =5959
                            Width =525
                            Height =240
                            Name ="Label36"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =623
                    Top =6319
                    TabIndex =17
                    Name ="dt"
                    ControlSource ="dt"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3772
                            Top =6325
                            Width =525
                            Height =240
                            Name ="Label38"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =5386
                    Width =2196
                    TabIndex =18
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="=ucurrentuser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3939
                            Top =5386
                            Width =1035
                            Height =240
                            Name ="IS16"
                            Caption ="نام كارشناس"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =5746
                    Width =2196
                    TabIndex =19
                    Name ="userid"
                    ControlSource ="userid"
                    DefaultValue ="=[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3939
                            Top =5746
                            Width =1035
                            Height =240
                            Name ="Label40"
                            Caption ="کد كارشناس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5268
                    Top =3799
                    TabIndex =20
                    Name ="Longitude"
                    ControlSource ="Longitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7823
                            Top =3799
                            Width =1110
                            Height =240
                            Name ="Label41"
                            Caption ="طول جغرافيايي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5268
                    Top =4139
                    TabIndex =21
                    Name ="Latitude"
                    ControlSource ="Latitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7950
                            Top =4139
                            Width =720
                            Height =240
                            Name ="Label42"
                            Caption ="عرض جغرافيايي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5177
                    Top =1871
                    Width =2661
                    Height =300
                    TabIndex =22
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =7945
                            Top =1876
                            Width =705
                            Height =240
                            Name ="Label45"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5155
                    Top =2272
                    Width =2661
                    Height =300
                    TabIndex =23
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =7923
                            Top =2266
                            Width =990
                            Height =240
                            Name ="Label46"
                            Caption ="کد شهرستان"
                        End
                    End
                End
            End
        End
    End
End
