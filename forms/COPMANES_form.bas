Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9581
    RowHeight =330
    ItemSuffix =48
    Left =465
    Top =1935
    Right =15075
    Bottom =9720
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x92c82a1ddceae540
    End
    RecordSource ="SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN ev"
        "entscount ON COPMANES.id = eventscount.idc where  COPMANES.userid=2 or COPMANES."
        "userid=5 or COPMANES.userid=6 or COPMANES.userid=7 ORDER BY COPMANES.id "
    Caption ="ليست مشتريان"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =8215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2271
                    Top =113
                    Width =2490
                    ColumnWidth =2910
                    ColumnOrder =2
                    Name ="COMPANY_NAME"
                    ControlSource ="COMPANY_NAME"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =113
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
                    Left =2271
                    Top =453
                    Width =2490
                    ColumnWidth =500
                    ColumnOrder =17
                    TabIndex =1
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =453
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
                    Left =2271
                    Top =793
                    Width =2490
                    ColumnWidth =3075
                    ColumnOrder =3
                    TabIndex =2
                    Name ="MANAGER"
                    ControlSource ="MANAGER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =793
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
                    Left =2271
                    Top =1133
                    Width =2490
                    ColumnWidth =2460
                    ColumnOrder =4
                    TabIndex =3
                    Name ="FACT_TEL"
                    ControlSource ="FACT_TEL"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =1133
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
                    Left =2271
                    Top =1474
                    Width =2490
                    ColumnWidth =1350
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =1474
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
                    Left =2271
                    Top =1814
                    ColumnWidth =1035
                    ColumnOrder =8
                    TabIndex =5
                    Name ="PERNUM"
                    ControlSource ="PERNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =1814
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
                    Left =2271
                    Top =2494
                    Width =2490
                    ColumnWidth =2505
                    ColumnOrder =10
                    TabIndex =7
                    Name ="PRODUCTS"
                    ControlSource ="PRODUCTS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =2494
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
                    Left =2271
                    Top =2834
                    Width =2490
                    ColumnWidth =5040
                    ColumnOrder =14
                    TabIndex =8
                    Name ="ADDR"
                    ControlSource ="ADDR"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =2834
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
                    Left =2271
                    Top =3174
                    Width =2490
                    ColumnWidth =1275
                    ColumnOrder =11
                    TabIndex =9
                    Name ="ACCOUNTANT"
                    ControlSource ="ACCOUNTANT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =3174
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
                    Left =2271
                    Top =3514
                    Width =2490
                    ColumnWidth =4095
                    ColumnOrder =12
                    TabIndex =10
                    Name ="SOFTWARE"
                    ControlSource ="SOFTWARE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =3514
                            Width =645
                            Height =240
                            Name ="IS11"
                            Caption ="سابقه کار"
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
                    Left =2271
                    Top =3855
                    Width =2490
                    ColumnWidth =3855
                    ColumnOrder =13
                    TabIndex =11
                    Name ="ESP_PERSON"
                    ControlSource ="ESP_PERSON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =3855
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
                    Left =2271
                    Top =4195
                    Width =2490
                    ColumnWidth =1605
                    ColumnOrder =15
                    TabIndex =12
                    Name ="REAGENT"
                    ControlSource ="REAGENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =4195
                            Width =480
                            Height =240
                            Name ="IS13"
                            Caption ="معرف"
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
                    Left =2271
                    Top =4875
                    Width =2490
                    ColumnWidth =12690
                    ColumnOrder =16
                    TabIndex =14
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =4875
                            Width =825
                            Height =240
                            Name ="IS15"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =1
                    Left =5952
                    Top =113
                    Width =3626
                    Height =1590
                    TabIndex =15
                    Name ="FEVENTS"
                    SourceObject ="Form.FEVENTS"
                    LinkChildFields ="idc"
                    LinkMasterFields ="id"
                End
                Begin ComboBox
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
                    Left =2271
                    Top =4535
                    ColumnWidth =1935
                    ColumnOrder =6
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
                    RowSource ="1;\"سلامتي مديران\";2;\"سلامتي پرسنل\";3;\"سلامتي ما\";4;\"تماس مجدد\";5;\"تماس "
                        "اونها\";6;\"تماس بگير\";7;\"قطع ارتباط\";8;\"ارتباط مجدد\";9;\"ارتباط مرتبط\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =457
                            Top =4535
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
                    Left =2324
                    Top =2154
                    Width =2490
                    ColumnWidth =1200
                    ColumnOrder =9
                    TabIndex =6
                    Name ="STATUS_FACT"
                    ControlSource ="STATUS_FACT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT COPMANES.STATUS_FACT FROM COPMANES GROUP BY COPMANES.STATUS_FACT; "
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =2154
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
                    Left =5215
                    Top =3004
                    ColumnWidth =2085
                    ColumnOrder =18
                    TabIndex =16
                    Name ="date_sabt"
                    ControlSource ="date_sabt"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7600
                            Top =3011
                            Width =915
                            Height =405
                            Name ="Label31"
                            Caption ="تاريخ و ساعت ثبت"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    OLETypeAllowed =0
                    Left =2838
                    Top =7483
                    Width =1821
                    Height =330
                    ColumnWidth =675
                    ColumnOrder =19
                    TabIndex =18
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1139
                            Top =7485
                            Width =570
                            Height =240
                            Name ="Label33"
                            Caption ="ضميمه"
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
                    Left =1080
                    Top =6009
                    ColumnWidth =630
                    ColumnOrder =20
                    TabIndex =19
                    Name ="id"
                    ControlSource ="id"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4229
                            Top =6015
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
                    Left =1080
                    Top =6375
                    ColumnWidth =1020
                    ColumnOrder =0
                    TabIndex =20
                    Name ="dt"
                    ControlSource ="dt"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4229
                            Top =6381
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
                    Left =1534
                    Top =5442
                    Width =2196
                    ColumnWidth =1335
                    ColumnOrder =7
                    TabIndex =17
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="=ucurrentuser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4396
                            Top =5442
                            Width =1035
                            Height =240
                            Name ="IS16"
                            Caption ="نام كارشناس"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1534
                    Top =5802
                    Width =2196
                    ColumnOrder =23
                    TabIndex =21
                    Name ="userid"
                    ControlSource ="userid"
                    DefaultValue ="=[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4396
                            Top =5802
                            Width =1035
                            Height =240
                            Name ="Label40"
                            Caption ="کد كارشناس"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5725
                    Top =3855
                    ColumnWidth =0
                    ColumnOrder =22
                    TabIndex =22
                    Name ="Longitude"
                    ControlSource ="Longitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8280
                            Top =3855
                            Width =1110
                            Height =240
                            Name ="Label41"
                            Caption ="طول جغرافيايي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5725
                    Top =4195
                    ColumnWidth =0
                    ColumnOrder =24
                    TabIndex =23
                    Name ="Latitude"
                    ControlSource ="Latitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8407
                            Top =4195
                            Width =720
                            Height =240
                            Name ="Label42"
                            Caption ="عرض جغرافيايي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
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
                    Left =5634
                    Top =1927
                    Width =2661
                    Height =300
                    ColumnWidth =0
                    ColumnOrder =21
                    TabIndex =24
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =8402
                            Top =1932
                            Width =705
                            Height =240
                            Name ="Label45"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
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
                    Left =5612
                    Top =2328
                    Width =2661
                    Height =300
                    ColumnWidth =0
                    ColumnOrder =25
                    TabIndex =25
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =8380
                            Top =2322
                            Width =990
                            Height =240
                            Name ="Label46"
                            Caption ="کد شهرستان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4591
                    Top =4762
                    ColumnWidth =500
                    ColumnOrder =1
                    TabIndex =26
                    Name ="idcn"
                    ControlSource ="idcn"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7665
                            Top =4755
                            Width =495
                            Height =240
                            Name ="Label47"
                            Caption ="گردش"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "COPMANES_form.cls"
