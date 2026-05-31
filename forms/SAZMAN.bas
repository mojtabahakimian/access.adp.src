Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =13230
    DatasheetFontHeight =10
    ItemSuffix =438
    Left =2925
    Top =3165
    Right =13530
    Bottom =9720
    HelpContextId =100107
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x407339af3268e240
    End
    RecordSource ="SAZMAN"
    Caption ="گزينه هاي سيستم"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =6576
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Width =13230
                    Height =5895
                    FontWeight =700
                    TabFixedWidth =1247
                    TabFixedHeight =397
                    Name ="TabCtl137"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="Page138"
                            Caption ="عمومي"
                            Begin
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =285
                                    Top =735
                                    Width =8347
                                    Height =300
                                    Name ="WIDTH_D"
                                    ControlSource ="WIDTH_D"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8744
                                            Top =740
                                            Width =1245
                                            Height =285
                                            Name ="Label37"
                                            Caption ="متن بالاي فاكتور:"
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
                                    Left =285
                                    Top =1075
                                    Width =8347
                                    Height =300
                                    TabIndex =1
                                    Name ="HIGH_D"
                                    ControlSource ="HIGH_D"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8744
                                            Top =1070
                                            Width =1185
                                            Height =285
                                            Name ="Label39"
                                            Caption ="متن پائين فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    Left =285
                                    Top =1430
                                    Width =8347
                                    Height =300
                                    TabIndex =2
                                    Name ="BACKPATH"
                                    ControlSource ="BACKPATH"
                                    DefaultValue ="\"D:\\DENABACK\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =8744
                                            Top =1425
                                            Width =1470
                                            Height =240
                                            Name ="Label97"
                                            Caption ="مسير فايل تصاوير:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =450
                                    Top =3165
                                    Width =8976
                                    Height =300
                                    TabIndex =3
                                    Name ="TFADDRESS"
                                    ControlSource ="TFADDRESS"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =9507
                                            Top =3165
                                            Width =615
                                            Height =285
                                            Name ="Label114"
                                            Caption ="آدرس:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =347
                                    Top =2994
                                    Width =9973
                                    Height =1021
                                    Name ="Box116"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    Left =5297
                                    Top =2709
                                    Width =885
                                    Height =255
                                    BackColor =-2147483633
                                    Name ="Label117"
                                    Caption ="آدرس ، تلفن"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6705
                                    Top =3525
                                    Width =2706
                                    Height =300
                                    TabIndex =4
                                    Name ="TFTEL"
                                    ControlSource ="TFTEL"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =9537
                                            Top =3525
                                            Width =525
                                            Height =285
                                            Name ="Label115"
                                            Caption ="تلفن:"
                                        End
                                    End
                                End
                                Begin BoundObjectFrame
                                    SizeMode =1
                                    OverlapFlags =215
                                    OLETypeAllowed =1
                                    Left =289
                                    Top =1862
                                    Width =10030
                                    Height =860
                                    TabIndex =5
                                    Name ="OLEBound157"
                                    ControlSource ="EMZA"
                                    OnDblClick ="[Event Procedure]"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    Left =6746
                                    Top =4244
                                    Width =2706
                                    Height =300
                                    TabIndex =6
                                    Name ="SERVERNAM"
                                    ControlSource ="SERVERNAM"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =9573
                                            Top =4243
                                            Width =750
                                            Height =285
                                            Name ="Label313"
                                            Caption ="نام سرور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =349
                                    Top =4697
                                    Width =4821
                                    Height =255
                                    TabIndex =7
                                    Name ="OPTIONSS"
                                    ControlSource ="OPTIONSS"
                                    OnDblClick ="[Event Procedure]"
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    Left =5272
                                    Top =4640
                                    Width =1290
                                    Height =345
                                    FontWeight =700
                                    TabIndex =8
                                    Name ="Command322"
                                    Caption ="تنظيمات بيشتر"
                                    OnClick ="[Event Procedure]"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="Page139"
                            Caption ="اسناد وحساب"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =225
                                    Top =779
                                    Width =2717
                                    Height =1318
                                    Name ="Frame98"
                                    ControlSource ="SANAD"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1594
                                            Top =665
                                            Width =675
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label99"
                                            Caption ="سند زدن"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2592
                                            Top =1032
                                            OptionValue =1
                                            Name ="Option101"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1115
                                                    Top =975
                                                    Width =1425
                                                    Height =255
                                                    Name ="Label102"
                                                    Caption ="در هر زمان اتوماتيك"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2592
                                            Top =1372
                                            OptionValue =2
                                            Name ="Option103"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =335
                                                    Top =1315
                                                    Width =2205
                                                    Height =285
                                                    Name ="Label104"
                                                    Caption ="در زمان تائيد توسط حسابداري"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2577
                                            Top =1703
                                            OptionValue =3
                                            Name ="Option149"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =980
                                                    Top =1656
                                                    Width =1560
                                                    Height =285
                                                    Name ="Label150"
                                                    Caption ="هيچكدام"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =3515
                                    Top =792
                                    Width =2579
                                    Height =943
                                    TabIndex =1
                                    Name ="Frame124"
                                    ControlSource ="PERSON"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3620
                                            Top =630
                                            Width =2385
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label125"
                                            Caption ="اشخاص مشخص شده  در حسابها"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5411
                                            Top =1030
                                            OptionValue =1
                                            Name ="Option126"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4032
                                                    Top =973
                                                    Width =1410
                                                    Height =285
                                                    Name ="Label127"
                                                    Caption ="نمايش داده شوند"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5411
                                            Top =1360
                                            OptionValue =2
                                            Name ="Option128"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4017
                                                    Top =1313
                                                    Width =1425
                                                    Height =285
                                                    Name ="Label129"
                                                    Caption ="نمايش داده نشوند"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =3485
                                    Top =2472
                                    Width =2620
                                    Height =1273
                                    TabIndex =2
                                    Name ="TKHF"
                                    ControlSource ="TKHF"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =4168
                                            Top =2310
                                            Width =1215
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label152"
                                            Caption ="تخفيفات فاكتورها"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5411
                                            Top =2710
                                            OptionValue =2
                                            Name ="Option153"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3747
                                                    Top =2674
                                                    Width =1695
                                                    Height =285
                                                    Name ="Label154"
                                                    Caption ="مصوب عمل شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5411
                                            Top =3040
                                            OptionValue =1
                                            Name ="Option155"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3485
                                                    Top =3015
                                                    Width =1950
                                                    Height =285
                                                    Name ="Label156"
                                                    Caption ="مبلغ كلي پشت فاكتور"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5404
                                            Top =3386
                                            OptionValue =3
                                            Name ="Option300"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3635
                                                    Top =3360
                                                    Width =1800
                                                    Height =285
                                                    Name ="Label301"
                                                    Caption ="به ريز كالا عمل شود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =3515
                                    Top =4427
                                    Width =621
                                    Height =255
                                    TabIndex =3
                                    Name ="TFTPAGE"
                                    ControlSource ="TFTPAGE"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =4195
                                            Top =4426
                                            Width =2325
                                            Height =240
                                            Name ="Label219"
                                            Caption ="فاصله از پائين كاغذ در چاپ سند :"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =275
                                    Top =2443
                                    Width =2729
                                    Height =1857
                                    TabIndex =4
                                    Name ="DAFT_ASN"
                                    SourceObject ="Form.FDAFT_ASN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =847
                                            Top =2205
                                            Width =1425
                                            Height =240
                                            Name ="DAFT_ASN Label"
                                            Caption ="دفتر اسناد دريافتني"
                                            EventProcPrefix ="DAFT_ASN_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7320
                                    Top =2052
                                    Width =621
                                    Height =450
                                    TabIndex =5
                                    Name ="Text281"
                                    ControlSource ="CPI"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =8015
                                            Top =2051
                                            Width =2310
                                            Height =435
                                            Name ="Label282"
                                            Caption ="تعداد تاريخهاي قابل برگشت ومشاهده در فاكتور فروش:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =9224
                                    Top =714
                                    Width =1139
                                    Height =227
                                    TabIndex =6
                                    Name ="SIGN"
                                    ControlSource ="SIGN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7377
                                            Top =690
                                            Width =2700
                                            Height =240
                                            Name ="Label311"
                                            Caption ="گردش كاري و فرمها با امضاء عمل شود"
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
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =275
                                    Top =4380
                                    Height =300
                                    TabIndex =7
                                    BackColor =13434828
                                    Name ="HESMBAA"
                                    ControlSource ="HESMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2045
                                            Top =4380
                                            Width =1350
                                            Height =285
                                            Name ="Label320"
                                            Caption ="ماليات:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =9219
                                    Top =1108
                                    Width =1139
                                    Height =227
                                    TabIndex =8
                                    Name ="CTL_DT"
                                    ControlSource ="CTL_DT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7372
                                            Top =1084
                                            Width =2700
                                            Height =240
                                            Name ="Label322"
                                            Caption ="تاريخ فقط مربوط به همين سال باشد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6305
                                    Top =4928
                                    TabIndex =9
                                    Name ="SAGHF"
                                    ControlSource ="SAGHF"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =2282
                                            Top =4875
                                            Width =3915
                                            Height =255
                                            Name ="Label294"
                                            Caption ="سقف اعتبار اشخاص بر اساس مانده  حساب  كنترل گردد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6303
                                    Top =5273
                                    TabIndex =10
                                    Name ="SAGHF2"
                                    ControlSource ="SAGHF2"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =2265
                                            Top =5220
                                            Width =3930
                                            Height =255
                                            Name ="Label315"
                                            Caption ="سقف اعتبار اشخاص بر اساس مانده   اسناد   كنترل گردد"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =7320
                                    Top =2732
                                    Width =561
                                    Height =255
                                    TabIndex =11
                                    Name ="ARSESH"
                                    ControlSource ="ARSESH"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    InputMask ="#\\%"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =7925
                                            Top =2731
                                            Width =1980
                                            Height =240
                                            Name ="Label317"
                                            Caption ="درصد ماليات بر ارزش افزوده:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    TripleState = NotDefault
                                    OverlapFlags =247
                                    Left =10150
                                    Top =1482
                                    Width =305
                                    TabIndex =12
                                    Name ="SNDKH"
                                    ControlSource ="SNDKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =8632
                                            Top =1429
                                            Width =1455
                                            Height =240
                                            Name ="Label372"
                                            Caption ="سند ها روزانه باشد"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =525
                            Width =13035
                            Height =5235
                            Name ="Page140"
                            Caption ="فاكتور و انبار"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =139
                                    Top =3057
                                    Width =2054
                                    Height =943
                                    Name ="Frame118"
                                    ControlSource ="KALA"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =332
                                            Top =2899
                                            Width =1710
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label119"
                                            Caption =" كالاها با موجودي صفر"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =1705
                                            Top =3295
                                            OptionValue =1
                                            Name ="Option120"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =273
                                                    Top =3247
                                                    Width =1365
                                                    Height =285
                                                    Name ="Label121"
                                                    Caption ="نمايش داده شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =1705
                                            Top =3625
                                            OptionValue =2
                                            Name ="Option122"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =213
                                                    Top =3587
                                                    Width =1425
                                                    Height =285
                                                    Name ="Label123"
                                                    Caption ="نمايش داده نشود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =5119
                                    Top =792
                                    Width =1755
                                    Height =943
                                    TabIndex =1
                                    Name ="Frame130"
                                    ControlSource ="WAR"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =5230
                                            Top =630
                                            Width =1425
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label131"
                                            Caption ="در زمان تغيير مقادير"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6467
                                            Top =1030
                                            OptionValue =1
                                            Name ="Option132"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5164
                                                    Top =973
                                                    Width =1305
                                                    Height =285
                                                    Name ="Label133"
                                                    Caption ="اخطار داده شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6467
                                            Top =1360
                                            OptionValue =2
                                            Name ="Option134"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5119
                                                    Top =1370
                                                    Width =1350
                                                    Height =285
                                                    Name ="Label135"
                                                    Caption ="اخطار داده نشود"
                                                End
                                            End
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
                                    Left =6933
                                    Top =1487
                                    Width =1251
                                    Height =300
                                    TabIndex =2
                                    Name ="DIG"
                                    ControlSource ="DIG"
                                    ValidationRule ="<14"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8220
                                            Top =1485
                                            Width =2190
                                            Height =285
                                            Name ="Label136"
                                            Caption ="تعداد اعشار در موجودي انبار:"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =2344
                                    Top =3057
                                    Width =2417
                                    Height =943
                                    TabIndex =3
                                    Name ="Frame141"
                                    ControlSource ="LST"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =2486
                                            Top =2899
                                            Width =2025
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label142"
                                            Caption ="در زمان ورود كالاي جديد"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =4315
                                            Top =3295
                                            OptionValue =1
                                            Name ="Option143"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2406
                                                    Top =3251
                                                    Width =1950
                                                    Height =285
                                                    Name ="Label144"
                                                    Caption ="ليست بطوراتوماتيك باز شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =4315
                                            Top =3625
                                            OptionValue =2
                                            Name ="Option145"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3231
                                                    Top =3591
                                                    Width =1125
                                                    Height =285
                                                    Name ="Label146"
                                                    Caption ="ليست باز نشود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =6625
                                    Top =2406
                                    Width =245
                                    TabIndex =4
                                    Name ="MAND"
                                    ControlSource ="MAND"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =3004
                                            Top =2355
                                            Width =3525
                                            Height =255
                                            Name ="Label183"
                                            Caption ="مانده حساب  اشخاص روي فاكتور نمايش داده شود"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10249
                                    Top =859
                                    TabIndex =5
                                    Name ="MOJU"
                                    ControlSource ="MOJU"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =7200
                                            Top =795
                                            Width =2925
                                            Height =255
                                            Name ="Label184"
                                            Caption ="موجودي انبار در زمان خروج  كنترل شود"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    DecimalPlaces =0
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =3575
                                    Top =4322
                                    Width =2031
                                    Height =313
                                    TabIndex =6
                                    Name ="DEFANB"
                                    ControlSource ="DEFANB"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                                    ColumnWidths ="0"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =5690
                                            Top =4320
                                            Width =1185
                                            Height =240
                                            Name ="Label254"
                                            Caption ="پيش فرض انبار :"
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
                                    DecimalPlaces =0
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =64
                                    Top =4323
                                    Width =1536
                                    Height =313
                                    TabIndex =7
                                    Name ="DEFTKH"
                                    ControlSource ="DEFTKH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                                    ColumnWidths ="0"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =1688
                                            Top =4326
                                            Width =1755
                                            Height =240
                                            Name ="Label256"
                                            Caption ="پيش فرض نوع مشتري :"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =3049
                                    Top =792
                                    Width =1725
                                    Height =943
                                    TabIndex =8
                                    Name ="FRUP"
                                    ControlSource ="FRUP"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =3334
                                            Top =630
                                            Width =1140
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label261"
                                            Caption ="ستون مبلغ "
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =4395
                                            Top =1030
                                            OptionValue =0
                                            Name ="Option262"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3167
                                                    Top =973
                                                    Width =1140
                                                    Height =285
                                                    Name ="Label263"
                                                    Caption ="قابل تغيير باشد"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =4395
                                            Top =1360
                                            Name ="Option264"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3109
                                                    Top =1365
                                                    Width =1200
                                                    Height =285
                                                    Name ="Label265"
                                                    Caption ="قابل تغيير نباشد"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =4929
                                    Top =3061
                                    Width =1890
                                    Height =943
                                    TabIndex =9
                                    Name ="UPDDATE"
                                    ControlSource ="UPDDATE"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =5199
                                            Top =2895
                                            Width =1455
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label267"
                                            Caption ="تاريخ فاكتور"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =6450
                                            Top =3299
                                            Name ="Option268"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5267
                                                    Top =3242
                                                    Width =1080
                                                    Height =285
                                                    Name ="Label269"
                                                    Caption ="قابل تغيير باشد"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =6450
                                            Top =3629
                                            OptionValue =0
                                            Name ="Option270"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4929
                                                    Top =3645
                                                    Width =1425
                                                    Height =285
                                                    Name ="Label271"
                                                    Caption ="قابل تغيير نباشد"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6591
                                    Top =2064
                                    Width =224
                                    Height =257
                                    TabIndex =10
                                    Name ="CODEVIEW"
                                    ControlSource ="CODEVIEW"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =3349
                                            Top =2040
                                            Width =3180
                                            Height =240
                                            Name ="Label308"
                                            Caption ="كد كالاها در زمان چاپ نمايش  داده شوند"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =5677
                                    Top =4852
                                    Width =1139
                                    Height =227
                                    TabIndex =11
                                    Name ="TRANSF"
                                    ControlSource ="TRANSF"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2884
                                            Top =4830
                                            Width =3645
                                            Height =240
                                            Name ="Label326"
                                            Caption ="سابقه تغييرات فاكتورها  بعد از قفل شدن ذخيره گردد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =9307
                                    Top =2721
                                    Width =1139
                                    TabIndex =12
                                    Name ="LOCKFAP"
                                    ControlSource ="LOCKFAP"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =8299
                                            Top =2670
                                            Width =1860
                                            Height =240
                                            Name ="Label328"
                                            Caption ="فاكتور بعداز چاپ قفل گردد"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =60
                                    Top =778
                                    Width =2857
                                    Height =2126
                                    TabIndex =13
                                    Name ="Frame105"
                                    ControlSource ="GHAYM"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =968
                                            Top =633
                                            Width =1695
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label106"
                                            Caption ="نمايش قيمت پيش فرض"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2562
                                            Top =1021
                                            OptionValue =1
                                            Name ="Option108"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =945
                                                    Top =978
                                                    Width =1530
                                                    Height =285
                                                    Name ="Label109"
                                                    Caption ="آخرين قيمت فروش"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2563
                                            Top =1351
                                            Width =261
                                            OptionValue =2
                                            Name ="Option110"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1230
                                                    Top =1318
                                                    Width =1245
                                                    Height =285
                                                    Name ="Label111"
                                                    Caption ="قيمت تعريف اوليه"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2563
                                            Top =1651
                                            Width =261
                                            OptionValue =4
                                            Name ="Option317"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =120
                                                    Top =1623
                                                    Width =2355
                                                    Height =285
                                                    Name ="Label318"
                                                    Caption ="آخرين قيمت فروش به اين مشتري"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2563
                                            Top =1965
                                            Width =261
                                            OptionValue =5
                                            Name ="Option320"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =60
                                                    Top =1932
                                                    Width =2400
                                                    Height =285
                                                    Name ="Label321"
                                                    Caption ="قيمت مصوب"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2563
                                            Top =2299
                                            Width =261
                                            OptionValue =3
                                            Name ="Option112"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1845
                                                    Top =2277
                                                    Width =645
                                                    Height =285
                                                    Name ="Label113"
                                                    Caption ="هيچكدام"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =2608
                                            Top =2632
                                            Width =216
                                            OptionValue =7
                                            Name ="elamgh"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1309
                                                    Top =2610
                                                    Width =1185
                                                    Height =285
                                                    Name ="Label376"
                                                    Caption ="اعلاميه قيمت"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6934
                                    Top =1876
                                    Width =1251
                                    Height =255
                                    TabIndex =14
                                    Name ="TFCODE_E"
                                    ControlSource ="TFCODE_E"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =8214
                                            Top =1875
                                            Width =2190
                                            Height =240
                                            Name ="Label347"
                                            Caption ="برچسب خدمات در چاپ فاكتور :"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10257
                                    Top =1193
                                    TabIndex =15
                                    Name ="RMOG"
                                    ControlSource ="RMOG"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =7311
                                            Top =1132
                                            Width =2820
                                            Height =255
                                            Name ="Label350"
                                            Caption ="موجودي بصورت قطعي كنترل شود"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =10203
                                    Top =2364
                                    Width =220
                                    Height =249
                                    TabIndex =16
                                    Name ="SERFACB"
                                    ControlSource ="SERFACB"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =6863
                                            Top =2334
                                            Width =3300
                                            Height =240
                                            Name ="Label371"
                                            Caption ="سفارشات تبلت بصورت فاكتور فروش ثبت شوند"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =510
                            Width =12960
                            Height =5250
                            Name ="Page159"
                            Caption ="حقوق"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =960
                                    Name ="HEZA"
                                    ControlSource ="HEZA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =930
                                            Width =675
                                            Height =285
                                            Name ="Label160"
                                            Caption ="اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =1300
                                    TabIndex =1
                                    Name ="HPAD"
                                    ControlSource ="HPAD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =1270
                                            Width =450
                                            Height =285
                                            Name ="Label161"
                                            Caption ="پاداش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =1640
                                    TabIndex =2
                                    Name ="HOLA"
                                    ControlSource ="HOLA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =1610
                                            Width =630
                                            Height =285
                                            Name ="Label162"
                                            Caption ="حق اولاد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =1962
                                    TabIndex =3
                                    Name ="HKHA"
                                    ControlSource ="HKHA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =1935
                                            Width =1170
                                            Height =285
                                            Name ="Label163"
                                            Caption ="خواربار و مسكن"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =2280
                                    Width =261
                                    TabIndex =4
                                    Name ="HNAH"
                                    ControlSource ="HNAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =2250
                                            Width =615
                                            Height =285
                                            Name ="Label181"
                                            Caption ="حق نهار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =8746
                                    Top =2598
                                    TabIndex =5
                                    Name ="HJAZ"
                                    ControlSource ="HJAZ"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =2568
                                            Width =645
                                            Height =285
                                            Name ="Label164"
                                            Caption ="حق جذب"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =8505
                                    Top =683
                                    Width =1772
                                    Height =4313
                                    Name ="Box176"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =8610
                                    Top =570
                                    Width =1545
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label177"
                                    Caption ="مشمول بيمه و ماليات"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =2909
                                    TabIndex =6
                                    Name ="HRAN"
                                    ControlSource ="HRAN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =2879
                                            Width =555
                                            Height =285
                                            Name ="Label165"
                                            Caption ="راندمان"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =3221
                                    TabIndex =7
                                    Name ="HCON"
                                    ControlSource ="HCON"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =3198
                                            Width =975
                                            Height =285
                                            Name ="Label167"
                                            Caption ="شرايط محيط"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =3540
                                    TabIndex =8
                                    Name ="HSAY"
                                    ControlSource ="HSAY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =3510
                                            Width =360
                                            Height =285
                                            Name ="Label166"
                                            Caption ="ساير"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =3885
                                    TabIndex =9
                                    Name ="HSHI"
                                    ControlSource ="HSHI"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =3855
                                            Width =840
                                            Height =285
                                            Name ="Label258"
                                            Caption ="حق شيفت"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =442
                                    Top =960
                                    Width =1911
                                    Height =300
                                    TabIndex =10
                                    BackColor =13434828
                                    Name ="HAZEDAR"
                                    ControlSource ="HAZEDAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2417
                                            Top =960
                                            Width =1425
                                            Height =285
                                            Name ="Label168"
                                            Caption ="هزينه حقوق اداري:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =442
                                    Top =1305
                                    Width =1911
                                    Height =300
                                    TabIndex =11
                                    BackColor =13434828
                                    Name ="HAZTOLID"
                                    ControlSource ="HAZTOLID"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2417
                                            Top =1305
                                            Width =1425
                                            Height =285
                                            Name ="Label248"
                                            Caption ="هزينه حقوق توليد:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =442
                                    Top =1650
                                    Width =1911
                                    Height =300
                                    TabIndex =12
                                    BackColor =13434828
                                    Name ="HAZFROOSH"
                                    ControlSource ="HAZFROOSH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =1650
                                            Width =1455
                                            Height =285
                                            Name ="Label250"
                                            Caption ="هزينه حقوق فروش:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =322
                                    Top =660
                                    Width =4689
                                    Height =4213
                                    Name ="Box178"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =1287
                                    Top =510
                                    Width =1875
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label179"
                                    Caption ="سر فصل حسابهاي حقوق"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =2005
                                    Width =1911
                                    Height =300
                                    TabIndex =13
                                    BackColor =13434828
                                    Name ="HAZKHADAMAT"
                                    ControlSource ="HAZKHADAMAT"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =2010
                                            Width =1500
                                            Height =285
                                            Name ="Label252"
                                            Caption ="هزينه حقوق خدمات:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =2355
                                    Width =1911
                                    Height =300
                                    TabIndex =14
                                    BackColor =13434828
                                    Name ="EDABIM"
                                    ControlSource ="EDABIM"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY NAME"
                                    ColumnWidths ="1701;3402"
                                    BeforeUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =2389
                                            Width =855
                                            Height =285
                                            Name ="Label169"
                                            Caption ="اداره  بيمه:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =2704
                                    Width =1911
                                    Height =300
                                    TabIndex =15
                                    BackColor =13434828
                                    Name ="HAZBIM"
                                    ControlSource ="HAZBIM"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =2700
                                            Width =1260
                                            Height =285
                                            Name ="Label170"
                                            Caption ="هزينه 23% بيمه:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =3045
                                    Width =1911
                                    Height =300
                                    TabIndex =16
                                    BackColor =13434828
                                    Name ="BESHO"
                                    ControlSource ="BESHO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2421
                                            Top =3045
                                            Width =1320
                                            Height =285
                                            Name ="Label171"
                                            Caption ="بستانكاران حقوق:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =3390
                                    Width =1911
                                    Height =300
                                    TabIndex =17
                                    BackColor =13434828
                                    Name ="BEDMOS"
                                    ControlSource ="BEDMOS"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =3390
                                            Width =1815
                                            Height =285
                                            Name ="Label172"
                                            Caption ="بدهكاران مساعده وغيره:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =3730
                                    Width =1911
                                    Height =300
                                    TabIndex =18
                                    BackColor =13434828
                                    Name ="PARDAKH"
                                    ControlSource ="PARDAKH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    BeforeUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2416
                                            Top =3730
                                            Width =1020
                                            Height =285
                                            Name ="Label173"
                                            Caption ="محل پرداخت:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =4080
                                    Width =1911
                                    Height =300
                                    TabIndex =19
                                    BackColor =13434828
                                    Name ="HAZMALI"
                                    ControlSource ="HAZMALI"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY NAME"
                                    ColumnWidths ="1701;3402"
                                    BeforeUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =4080
                                            Width =1710
                                            Height =285
                                            Name ="Label174"
                                            Caption ="اداره دارائي بابت ماليات:"
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
                                    ListWidth =4536
                                    Left =442
                                    Top =4425
                                    Width =1911
                                    Height =300
                                    TabIndex =20
                                    BackColor =13434828
                                    Name ="PSANDHES"
                                    ControlSource ="PSANDHES"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2422
                                            Top =4425
                                            Width =1710
                                            Height =285
                                            Name ="Label276"
                                            Caption ="صندوق  پرسنل:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =2241
                                    Top =5453
                                    Width =1671
                                    Height =300
                                    TabIndex =21
                                    Name ="SAGHFH"
                                    ControlSource ="SAGHFH"
                                    Format ="#,###"
                                    OnDblClick ="[Event Procedure]"
                                    OnKeyPress ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3975
                                            Top =5449
                                            Width =2160
                                            Height =285
                                            Name ="Label182"
                                            Caption ="حد اقل حقوق مشمول ماليات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =6406
                                    Top =5333
                                    Width =876
                                    Height =300
                                    TabIndex =22
                                    Name ="SANAVP"
                                    ControlSource ="SANAVP"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =7345
                                            Top =5329
                                            Width =1185
                                            Height =240
                                            Name ="Label278"
                                            Caption ="حق سنوات پايه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1390
                                    Top =4942
                                    Width =876
                                    Height =300
                                    TabIndex =23
                                    Name ="BON"
                                    ControlSource ="BON"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =2329
                                            Top =4938
                                            Width =1185
                                            Height =240
                                            Name ="Label285"
                                            Caption ="بن:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =2411
                                    Top =4938
                                    Width =1590
                                    Height =405
                                    TabIndex =24
                                    Name ="Command309"
                                    Caption ="جدول ماليات  حقوق"
                                    OnClick ="[Event Procedure]"
                                    ObjectPalette = Begin
                                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                                        0xc0c0c00080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                                        0xffffff0000000000
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =4225
                                    Top =5082
                                    Width =396
                                    Height =189
                                    TabIndex =25
                                    Name ="Check315"
                                    ControlSource ="TENDAR"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            Left =4676
                                            Top =5052
                                            Width =915
                                            Height =240
                                            Name ="Label316"
                                            Caption ="ده درصدي"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =225
                                    Top =5310
                                    Width =1266
                                    Height =336
                                    TabIndex =26
                                    Name ="Command314"
                                    Caption ="بيشتر"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =4225
                                    TabIndex =27
                                    Name ="HBON"
                                    ControlSource ="HBON"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =4195
                                            Width =780
                                            Height =285
                                            Name ="Label375"
                                            Caption ="بن"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =8746
                                    Top =4575
                                    TabIndex =28
                                    Name ="HTAHOL"
                                    ControlSource ="HTAHOL"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =8976
                                            Top =4545
                                            Width =780
                                            Height =285
                                            Name ="Label437"
                                            Caption ="حق تاهل"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =525
                            Width =12975
                            Height =5235
                            Name ="sanati"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =1065
                                    Name ="SA_HOGH"
                                    ControlSource ="SA_HOGH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =1035
                                            Width =495
                                            Height =240
                                            Name ="Label233"
                                            Caption ="حقوق"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =1405
                                    TabIndex =1
                                    Name ="SA_40EZ"
                                    ControlSource ="SA_40EZ"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =1375
                                            Width =960
                                            Height =240
                                            Name ="Label234"
                                            Caption ="40%اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =1745
                                    TabIndex =2
                                    Name ="SA_EZAF"
                                    ControlSource ="SA_EZAF"
                                    AfterUpdate ="[Event Procedure]"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =1715
                                            Width =690
                                            Height =240
                                            Name ="Label235"
                                            Caption ="اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =2085
                                    TabIndex =3
                                    Name ="SA_PADA"
                                    ControlSource ="SA_PADA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =2055
                                            Width =480
                                            Height =240
                                            Name ="Label236"
                                            Caption ="پاداش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =2426
                                    TabIndex =4
                                    Name ="SA_HOLA"
                                    ControlSource ="SA_HOLA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =2396
                                            Width =675
                                            Height =240
                                            Name ="Label237"
                                            Caption ="حق اولاد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7254
                                    Top =2766
                                    TabIndex =5
                                    Name ="SA_KHAR"
                                    ControlSource ="SA_KHAR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =7484
                                            Top =2736
                                            Width =1155
                                            Height =240
                                            Name ="Label238"
                                            Caption ="خواربار و مسكن"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5499
                                    Top =1050
                                    TabIndex =6
                                    Name ="SA_NAHA"
                                    ControlSource ="SA_NAHA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5729
                                            Top =1020
                                            Width =615
                                            Height =240
                                            Name ="Label239"
                                            Caption ="حق نهار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5490
                                    Top =1392
                                    TabIndex =7
                                    Name ="SA_JAZB"
                                    ControlSource ="SA_JAZB"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5720
                                            Top =1362
                                            Width =705
                                            Height =240
                                            Name ="Label240"
                                            Caption ="حق جذب"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5490
                                    Top =1732
                                    TabIndex =8
                                    Name ="SA_RAND"
                                    ControlSource ="SA_RAND"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5720
                                            Top =1702
                                            Width =555
                                            Height =240
                                            Name ="Label241"
                                            Caption ="راندمان"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5490
                                    Top =2072
                                    TabIndex =9
                                    Name ="SA_COND"
                                    ControlSource ="SA_COND"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5720
                                            Top =2042
                                            Width =945
                                            Height =240
                                            Name ="Label242"
                                            Caption ="شرايط محيط"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5490
                                    Top =2412
                                    TabIndex =10
                                    Name ="SA_SAYE"
                                    ControlSource ="SA_SAYE"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5720
                                            Top =2382
                                            Width =405
                                            Height =240
                                            Name ="Label243"
                                            Caption ="ساير"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =5490
                                    Top =2753
                                    TabIndex =11
                                    Name ="SA_23BI"
                                    ControlSource ="SA_23BI"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5720
                                            Top =2723
                                            Width =645
                                            Height =240
                                            Name ="Label244"
                                            Caption ="23%بيمه"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =5265
                                    Top =683
                                    Width =3515
                                    Height =2561
                                    Name ="Box245"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =5325
                                    Top =570
                                    Width =3285
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label246"
                                    Caption ="عواملي كه در نرخ يك ساعت كار بايد حذف شوند"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =322
                                    Top =656
                                    TabIndex =12
                                    Name ="ECONM"
                                    ControlSource ="ECONM"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =557
                                            Top =570
                                            Width =2370
                                            Height =300
                                            Name ="Label259"
                                            Caption ="مواد اوليه در زمان توليد خارج  شود"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =5220
                                    Top =3454
                                    Width =3571
                                    Height =1078
                                    TabIndex =13
                                    Name ="FINALS"
                                    ControlSource ="FINALS"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =5742
                                            Top =3285
                                            Width =2505
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label273"
                                            Caption ="خروج مواد به كالاي درجريان "
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =8350
                                            Top =3732
                                            Width =245
                                            Name ="Option274"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =6582
                                                    Top =3675
                                                    Width =1650
                                                    Height =285
                                                    Name ="Label275"
                                                    Caption ="بصورت كلي انجام شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =8350
                                            Top =4062
                                            Width =335
                                            OptionValue =0
                                            Name ="Option276"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5415
                                                    Top =4072
                                                    Width =2820
                                                    Height =285
                                                    Name ="Label277"
                                                    Caption =" به معين كالاي ساخته شده منتقل شود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =225
                                    Top =926
                                    Width =275
                                    TabIndex =14
                                    Name ="SANAT"
                                    ControlSource ="SANAT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =565
                                            Top =915
                                            Width =2370
                                            Height =285
                                            Name ="Label303"
                                            Caption ="ارتيكل اسناد صنعتي صادر شود"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =1572
                                    Width =1086
                                    TabIndex =15
                                    Name ="STFR"
                                    ControlSource ="STFR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =1575
                                            Width =930
                                            Height =240
                                            Name ="Label323"
                                            Caption ="فاكتور فروش"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =1912
                                    Width =1086
                                    TabIndex =16
                                    Name ="STKH"
                                    ControlSource ="STKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =1915
                                            Width =885
                                            Height =240
                                            Name ="Label324"
                                            Caption ="فاكتور خريد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =2252
                                    Width =1086
                                    TabIndex =17
                                    Name ="STHFR"
                                    ControlSource ="STHFR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =2245
                                            Width =960
                                            Height =240
                                            Name ="Label325"
                                            Caption ="حواله فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =2593
                                    Width =1086
                                    TabIndex =18
                                    Name ="STHKH"
                                    ControlSource ="STHKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =2593
                                            Width =900
                                            Height =240
                                            Name ="Label327"
                                            Caption ="رسيد انبار:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =2933
                                    Width =1086
                                    TabIndex =19
                                    Name ="STENT"
                                    ControlSource ="STENT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =2933
                                            Width =885
                                            Height =240
                                            Name ="Label329"
                                            Caption ="انتقالي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =3273
                                    Width =1086
                                    TabIndex =20
                                    Name ="STKHS"
                                    ControlSource ="STKHS"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =3273
                                            Width =885
                                            Height =240
                                            Name ="Label330"
                                            Caption ="خروج ساير:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =3613
                                    Width =1086
                                    TabIndex =21
                                    Name ="STKHH"
                                    ControlSource ="STKHH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =3613
                                            Width =900
                                            Height =240
                                            Name ="Label331"
                                            Caption ="خروج:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =3953
                                    Width =1086
                                    TabIndex =22
                                    Name ="STTOL"
                                    ControlSource ="STTOL"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =3953
                                            Width =885
                                            Height =240
                                            Name ="Label332"
                                            Caption ="توليد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =4293
                                    Width =1086
                                    TabIndex =23
                                    Name ="STFRB"
                                    ControlSource ="STFRB"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =4300
                                            Width =1125
                                            Height =240
                                            Name ="Label333"
                                            Caption ="برگشت فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =4634
                                    Width =1086
                                    TabIndex =24
                                    Name ="STBKH"
                                    ControlSource ="STBKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =4630
                                            Width =1020
                                            Height =240
                                            Name ="Label334"
                                            Caption ="برگشت خريد:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =120
                                    Top =1410
                                    Width =4866
                                    Height =4145
                                    Name ="Box341"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    Left =1020
                                    Top =1260
                                    Width =1035
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Label342"
                                    Caption ="شماره شروع "
                                    ControlTipText ="قبل از صدور سند در هر بخش ميتوانيد شماره شروع سند را در ذيل مشخص كنيد"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =4935
                                    Width =1086
                                    TabIndex =25
                                    Name ="STMO"
                                    ControlSource ="STMO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1815
                                            Top =4935
                                            Width =1155
                                            Height =240
                                            Name ="Label344"
                                            Caption ="رسيد مستقيم:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =388
                                    Top =5224
                                    Width =1086
                                    TabIndex =26
                                    Name ="STKHA"
                                    ControlSource ="STKHA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1815
                                            Top =5220
                                            Width =1035
                                            Height =240
                                            Name ="Label346"
                                            Caption ="فاكتور خدمات:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="Page283"
                            Caption ="ISO"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3631
                                    Top =970
                                    Width =1191
                                    Height =330
                                    Name ="ISO_FROOSH"
                                    ControlSource ="ISO_FROOSH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4882
                                            Top =963
                                            Width =1260
                                            Height =330
                                            Name ="Label287"
                                            Caption ="حواله انبار فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3631
                                    Top =1387
                                    Width =1191
                                    Height =330
                                    TabIndex =1
                                    Name ="ISO_KHAREED"
                                    ControlSource ="ISO_KHAREED"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4882
                                            Top =1380
                                            Width =1260
                                            Height =330
                                            Name ="Label289"
                                            Caption ="رسيد انبار خريد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3631
                                    Top =1804
                                    Width =1191
                                    Height =330
                                    TabIndex =2
                                    Name ="ISO_MAVAD"
                                    ControlSource ="ISO_MAVAD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4882
                                            Top =1800
                                            Width =1590
                                            Height =345
                                            Name ="Label291"
                                            Caption ="حواله خروج مواد اوليه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3631
                                    Top =2218
                                    Width =1191
                                    Height =330
                                    TabIndex =3
                                    Name ="ISO_TOLID"
                                    ControlSource ="ISO_TOLID"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4882
                                            Top =2214
                                            Width =1590
                                            Height =345
                                            Name ="Label293"
                                            Caption ="رسيد انبار  آمار توليد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3631
                                    Top =2632
                                    Width =1191
                                    Height =330
                                    TabIndex =4
                                    Name ="ISO_MAVADSAYER"
                                    ControlSource ="ISO_MAVADSAYER"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4882
                                            Top =2628
                                            Width =1590
                                            Height =345
                                            Name ="Label295"
                                            Caption ="حواله خروج ساير مواد:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =3185
                                    Top =737
                                    Width =3401
                                    Height =2727
                                    Name ="Box296"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =5272
                                    Top =623
                                    Width =1125
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Label297"
                                    Caption ="كدينگ فرمهاي "
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3630
                                    Top =3019
                                    Width =1191
                                    Height =330
                                    TabIndex =5
                                    Name ="ISO_DTOLID"
                                    ControlSource ="ISO_DTOLID"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4881
                                            Top =3015
                                            Width =1590
                                            Height =345
                                            Name ="Label362"
                                            Caption ="درخواست توليد:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="SMS"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =870
                                    Width =1956
                                    BackColor =13434828
                                    Name ="SMS_USERNAME"
                                    ControlSource ="SMS_USERNAME"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =870
                                            Width =945
                                            Height =240
                                            Name ="Label337"
                                            Caption ="USERNAME:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =1210
                                    Width =1956
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="SMS_PASSWORD"
                                    ControlSource ="SMS_PASSWORD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =1210
                                            Width =990
                                            Height =240
                                            Name ="Label338"
                                            Caption ="PASSWORD:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =1550
                                    Width =5076
                                    TabIndex =2
                                    BackColor =13434828
                                    Name ="SMS_LIBKEY"
                                    ControlSource ="SMS_LIBKEY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =1550
                                            Width =630
                                            Height =240
                                            Name ="Label339"
                                            Caption ="LIBKEY:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =1890
                                    Width =1956
                                    TabIndex =3
                                    BackColor =13434828
                                    Name ="SMS_TSMSHOST"
                                    ControlSource ="SMS_TSMSHOST"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =1890
                                            Width =930
                                            Height =240
                                            Name ="Label340"
                                            Caption ="TSMSHOST:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =2598
                                    Width =1956
                                    TabIndex =4
                                    BackColor =10092543
                                    Name ="SMS_ProxyUserName"
                                    ControlSource ="SMS_ProxyUserName"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =2598
                                            Width =1290
                                            Height =240
                                            Name ="Label341"
                                            Caption ="ProxyUserName:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =2938
                                    Width =1956
                                    TabIndex =5
                                    BackColor =10092543
                                    Name ="SMS_ProxyPassword"
                                    ControlSource ="SMS_ProxyPassword"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =2938
                                            Width =1245
                                            Height =240
                                            Name ="Label343"
                                            Caption ="ProxyPassword:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =3278
                                    Width =1956
                                    TabIndex =6
                                    BackColor =10092543
                                    Name ="SMS_ProxyServer"
                                    ControlSource ="SMS_ProxyServer"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =3278
                                            Width =1035
                                            Height =240
                                            Name ="Label345"
                                            Caption ="ProxyServer:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =1
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =3618
                                    Width =1956
                                    TabIndex =7
                                    BackColor =10092543
                                    Name ="SMS_ProxyPort"
                                    ControlSource ="SMS_ProxyPort"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =220
                                            Top =3618
                                            Width =855
                                            Height =240
                                            Name ="Label352"
                                            Caption ="ProxyPort:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =1666
                                    Top =4390
                                    Width =1956
                                    TabIndex =8
                                    Name ="SMS_FirewallPassword"
                                    ControlSource ="SMS_FirewallPassword"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            Left =220
                                            Top =4390
                                            Width =1365
                                            Height =240
                                            Name ="Label349"
                                            Caption ="FirewallPassword:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =165
                                    Top =701
                                    Width =10083
                                    Height =1581
                                    Name ="Box353"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    Left =4600
                                    Top =570
                                    Width =2055
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label355"
                                    Caption ="تنظيمات اعتبار پيامك"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =165
                                    Top =2525
                                    Width =10083
                                    Height =1401
                                    Name ="Box357"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    Left =4600
                                    Top =2400
                                    Width =2055
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label358"
                                    Caption ="تنظيمات شبكه داخلي"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =9805
                                    Top =4103
                                    TabIndex =9
                                    Name ="DSMS"
                                    ControlSource ="DSMS"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6697
                                            Top =4050
                                            Width =3000
                                            Height =240
                                            Name ="Label361"
                                            Caption ="ارسال پيامك فقط از طريق سرور انجام  شود"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    ReadingOrder =2
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =3400
                                    Top =4386
                                    Width =5436
                                    Height =285
                                    TabIndex =10
                                    Name ="SMS_OWNER"
                                    ControlSource ="SMS_OWNER"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =8902
                                            Top =4380
                                            Width =1290
                                            Height =240
                                            Name ="Label364"
                                            Caption ="عنوان پيامك:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =3328
                                    Top =4088
                                    TabIndex =11
                                    Name ="PRMFR"
                                    ControlSource ="PRMFR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =220
                                            Top =4035
                                            Width =3000
                                            Height =240
                                            Name ="Label366"
                                            Caption ="ارسال پيامك با تاييد كاربر باشد"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =3742
                                    Top =4960
                                    Width =2475
                                    Height =345
                                    FontWeight =700
                                    TabIndex =12
                                    Name ="Command367"
                                    Caption ="تعريف فرمت ارسال پيامك"
                                    OnClick ="[Event Procedure]"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2845
                                    Top =5006
                                    TabIndex =13
                                    Name ="SMSACT"
                                    ControlSource ="SMSACT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1072
                                            Top =4953
                                            Width =1665
                                            Height =240
                                            Name ="Label369"
                                            Caption ="ارسال پيامك فعال باشد"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =525
                            Width =12975
                            Height =5235
                            Name ="crm"
                            Caption ="تنظيماتCRM"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =1096
                                    Width =1791
                                    Height =255
                                    Name ="IT1"
                                    ControlSource ="IT1"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =1096
                                            Width =765
                                            Height =240
                                            Name ="Label40"
                                            Caption ="1-آيتم اول"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =1426
                                    Width =1791
                                    Height =255
                                    Name ="IT2"
                                    ControlSource ="IT2"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =1426
                                            Width =795
                                            Height =240
                                            Name ="Label41"
                                            Caption ="2-آيتم دوم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =1757
                                    Width =1791
                                    Height =255
                                    Name ="IT3"
                                    ControlSource ="IT3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =1757
                                            Width =885
                                            Height =240
                                            Name ="Label42"
                                            Caption ="3-آيتم سوم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =2088
                                    Width =1791
                                    Height =255
                                    Name ="IT4"
                                    ControlSource ="IT4"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =2088
                                            Width =945
                                            Height =240
                                            Name ="Label43"
                                            Caption ="4-آيتم چهارم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =2418
                                    Width =1791
                                    Height =255
                                    Name ="IT5"
                                    ControlSource ="IT5"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =2418
                                            Width =870
                                            Height =240
                                            Name ="Label44"
                                            Caption ="5-آيتم پنجم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =2749
                                    Width =1791
                                    Height =255
                                    Name ="IT6"
                                    ControlSource ="IT6"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =2749
                                            Width =1005
                                            Height =240
                                            Name ="Label45"
                                            Caption ="6-آيتم ششم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =3080
                                    Width =1791
                                    Height =255
                                    Name ="IT7"
                                    ControlSource ="IT7"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =3080
                                            Width =930
                                            Height =240
                                            Name ="Label46"
                                            Caption ="7-آيتم هفتم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =3411
                                    Width =1791
                                    Height =255
                                    Name ="IT8"
                                    ControlSource ="IT8"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =3411
                                            Width =1020
                                            Height =240
                                            Name ="Label47"
                                            Caption ="8-آيتم هشتم"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =7115
                                    Top =3741
                                    Width =1791
                                    Height =255
                                    Name ="IT9"
                                    ControlSource ="IT9"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =9119
                                            Top =3741
                                            Width =795
                                            Height =240
                                            Name ="Label48"
                                            Caption ="9-آيتم نهم"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =6945
                                    Top =769
                                    Width =3311
                                    Height =3307
                                    Name ="Box49"
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =7815
                                    Top =634
                                    Width =1245
                                    Height =240
                                    BackColor =-2147483633
                                    BorderColor =12632256
                                    Name ="Label50"
                                    Caption ="عناوين آيتمها"
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =1099
                                    Width =1791
                                    Height =255
                                    Name ="IS1"
                                    ControlSource ="IS1"
                                    DefaultValue ="=\"نام شركت\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =1099
                                            Width =765
                                            Height =240
                                            Name ="Label52"
                                            Caption ="نام شركت"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =1429
                                    Width =1791
                                    Height =255
                                    Name ="IS2"
                                    ControlSource ="IS2"
                                    DefaultValue ="\"شهر\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =1429
                                            Width =795
                                            Height =240
                                            Name ="Label54"
                                            Caption ="شهر"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =1760
                                    Width =1791
                                    Height =255
                                    Name ="IS3"
                                    ControlSource ="IS3"
                                    DefaultValue ="\"مدير عامل\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =1760
                                            Width =885
                                            Height =240
                                            Name ="Label56"
                                            Caption ="مدير عامل"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =2091
                                    Width =1791
                                    Height =255
                                    Name ="IS4"
                                    ControlSource ="IS4"
                                    DefaultValue ="\"تلفن كارخانه\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =2091
                                            Width =945
                                            Height =240
                                            Name ="Label58"
                                            Caption ="تلفن كارخانه"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =2421
                                    Width =1791
                                    Height =255
                                    Name ="IS5"
                                    ControlSource ="IS5"
                                    DefaultValue ="\"موبايل مدير عامل\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =2421
                                            Width =1245
                                            Height =240
                                            Name ="Label60"
                                            Caption ="موبايل مدير عامل"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =2752
                                    Width =1791
                                    Height =255
                                    Name ="IS6"
                                    ControlSource ="IS6"
                                    DefaultValue ="\"تعداد پرسنل\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =2752
                                            Width =1005
                                            Height =240
                                            Name ="Label62"
                                            Caption ="تعداد پرسنل"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =3083
                                    Width =1791
                                    Height =255
                                    Name ="IS7"
                                    ControlSource ="IS7"
                                    DefaultValue ="\"وضعيت كارخانه\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =3083
                                            Width =1095
                                            Height =240
                                            Name ="Label64"
                                            Caption ="وضعيت كارخانه"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =3414
                                    Width =1791
                                    Height =255
                                    Name ="IS8"
                                    ControlSource ="IS8"
                                    DefaultValue ="\"محصولات\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =3414
                                            Width =1020
                                            Height =240
                                            Name ="Label66"
                                            Caption ="محصولات"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3540
                                    Top =3744
                                    Width =1791
                                    Height =255
                                    Name ="IS9"
                                    ControlSource ="IS9"
                                    DefaultValue ="\"آدرس\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =5544
                                            Top =3744
                                            Width =795
                                            Height =240
                                            Name ="Label68"
                                            Caption ="آدرس"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =120
                                    Top =769
                                    Width =6701
                                    Height =3307
                                    Name ="Box69"
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2452
                                    Top =630
                                    Width =1275
                                    Height =240
                                    BackColor =-2147483633
                                    BorderColor =12632256
                                    Name ="Label70"
                                    Caption ="عناوين ستونها"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =1054
                                    Width =1791
                                    Height =255
                                    Name ="IS10"
                                    ControlSource ="IS10"
                                    DefaultValue ="\"حسابدار\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =1054
                                            Width =765
                                            Height =240
                                            Name ="Label72"
                                            Caption ="حسابدار"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =1384
                                    Width =1791
                                    Height =255
                                    Name ="IS11"
                                    ControlSource ="IS11"
                                    DefaultValue ="\"نرم افزار\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =1384
                                            Width =795
                                            Height =240
                                            Name ="Label74"
                                            Caption ="نرم افزار"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =1715
                                    Width =1791
                                    Height =255
                                    Name ="IS12"
                                    ControlSource ="IS12"
                                    DefaultValue ="\"اشخاص ذي نفوذ\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =1715
                                            Width =1230
                                            Height =240
                                            Name ="Label76"
                                            Caption ="اشخاص ذي نفوذ"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =2046
                                    Width =1791
                                    Height =255
                                    Name ="IS13"
                                    ControlSource ="IS13"
                                    DefaultValue ="\"تلفن كارخانه\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =2046
                                            Width =945
                                            Height =240
                                            Name ="Label78"
                                            Caption ="تلفن كارخانه"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =2376
                                    Width =1791
                                    Height =255
                                    Name ="IS14"
                                    ControlSource ="IS14"
                                    DefaultValue ="\"معرف\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =2376
                                            Width =1245
                                            Height =240
                                            Name ="Label80"
                                            Caption ="معرف"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =2707
                                    Width =1791
                                    Height =255
                                    Name ="IS15"
                                    ControlSource ="IS15"
                                    DefaultValue ="\"وضعيت\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =2707
                                            Width =1005
                                            Height =240
                                            Name ="Label82"
                                            Caption ="وضعيت"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =195
                                    Top =3038
                                    Width =1791
                                    Height =255
                                    Name ="IS16"
                                    ControlSource ="IS16"
                                    DefaultValue ="\"توضيحات\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2199
                                            Top =3038
                                            Width =1095
                                            Height =240
                                            Name ="Label84"
                                            Caption ="توضيحات"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="SSM"
                            Begin
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =4701
                                    Top =735
                                    Width =846
                                    BackColor =13434828
                                    Name ="SSMTRTAKM"
                                    ControlSource ="SSMTRTAKM"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5610
                                            Top =735
                                            Width =3690
                                            Height =240
                                            Name ="Label379"
                                            Caption ="تعداد روزهاي تاخير مجاز :"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =4701
                                    Top =1075
                                    Width =846
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="SSMTRTAGM"
                                    ControlSource ="SSMTRTAGM"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5610
                                            Top =1080
                                            Width =3690
                                            Height =240
                                            Name ="Label381"
                                            Caption ="\011تعداد روزهاي تعجيل مجاز:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =4701
                                    Top =1755
                                    Width =846
                                    TabIndex =2
                                    BackColor =13434828
                                    Name ="SSMTBMON"
                                    ControlSource ="SSMTBMON"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5610
                                            Top =1755
                                            Width =3690
                                            Height =240
                                            Name ="Label385"
                                            Caption ="\011تعداد ماه براي كنترل مستر كاركت مثلا قبل از يك ماه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =4695
                                    Top =2070
                                    Width =846
                                    TabIndex =3
                                    BackColor =13434828
                                    Name ="SSMDARSAD"
                                    ControlSource ="SSMDARSAD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5604
                                            Top =2070
                                            Width =3690
                                            Height =240
                                            Name ="Label410"
                                            Caption ="\011درصد  كاهش يا افزايش تخفيف  برحسب روز:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5325
                                    Top =1425
                                    Width =397
                                    Height =227
                                    TabIndex =4
                                    Name ="Check411"
                                    ControlSource ="SSMSNDAUTO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5610
                                            Top =1395
                                            Width =3825
                                            Height =240
                                            Name ="Label383"
                                            Caption ="\011سندهاي تاخير اتوماتک در يکم هر ماه صادر شود باشد :"
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
                                    ListWidth =4536
                                    Left =396
                                    Top =793
                                    Width =1911
                                    Height =300
                                    TabIndex =5
                                    BackColor =13434828
                                    Name ="HDARKASRTAKHF"
                                    ControlSource ="HDARKASRTAKHF"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT hes, NAME FROM CUST_HESAB ORDER BY hes"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2370
                                            Top =795
                                            Width =1920
                                            Height =285
                                            Name ="Label414"
                                            Caption ="حساب درآمد  کسر تخفيف:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =525
                            Width =12960
                            Height =5235
                            Name ="MOADIAN"
                            PictureData = Begin
                                0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                0xffff0000ffffff00dadadadadadadadaadadada9adadadaddadada999adadada ,
                                0xadada99999adadaddada9999999adadaada999999999adadda99999999999ada ,
                                0xa9999999999999ad999999999999999aa9999999999999adda99999999999ada ,
                                0xada999999999adaddada9999999adadaadada99999adadaddadada999adadada ,
                                0xadadada9adadadad000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                0x0000000000000000
                            End
                            Caption =" موديان"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =165
                                    Top =735
                                    Width =11781
                                    Height =2385
                                    BackColor =13434828
                                    Name ="PRIVIATEKEY"
                                    ControlSource ="PRIVIATEKEY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =12009
                                            Top =735
                                            Width =1080
                                            Height =240
                                            Name ="Label417"
                                            Caption ="کليد خصوصي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =165
                                    Top =735
                                    Width =5316
                                    Height =2445
                                    TabIndex =1
                                    BackColor =10092543
                                    Name ="PUBLICKEY"
                                    ControlSource ="PUBLICKEY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5544
                                            Top =735
                                            Width =1005
                                            Height =240
                                            Name ="Label429"
                                            Caption ="کليد عمومي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =9300
                                    Top =4875
                                    Width =1731
                                    Height =330
                                    TabIndex =2
                                    BackColor =13434828
                                    Name ="MEMORYID"
                                    ControlSource ="MEMORYID"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =11094
                                            Top =4875
                                            Width =1110
                                            Height =240
                                            Name ="Label431"
                                            Caption ="حافظه مالياتي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =5790
                                    Top =4920
                                    Width =1731
                                    Height =330
                                    TabIndex =3
                                    BackColor =13434828
                                    Name ="MEMORYIDsand"
                                    ControlSource ="MEMORYIDsand"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =7579
                                            Top =4923
                                            Width =1605
                                            Height =240
                                            Name ="Label433"
                                            Caption ="حافظه مالياتي تستي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    KeyboardLanguage =11
                                    IMESentenceMode =3
                                    Left =168
                                    Top =3231
                                    Width =11781
                                    Height =1590
                                    TabIndex =4
                                    BackColor =10092543
                                    Name ="Dcertificate"
                                    ControlSource ="Dcertificate"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =12012
                                            Top =3231
                                            Width =1005
                                            Height =240
                                            Name ="Label435"
                                            Caption ="گواهي امضا:"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =3930
                    Top =6120
                    Width =1290
                    Height =345
                    FontWeight =700
                    TabIndex =1
                    Name ="Command147"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =208
                    TextFontCharSet =0
                    Left =2655
                    Top =6120
                    Width =1275
                    Height =345
                    FontWeight =700
                    TabIndex =2
                    Name ="Command148"
                    Caption ="&ذخيره و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1365
                    Top =6120
                    Width =1290
                    Height =345
                    FontWeight =700
                    TabIndex =3
                    Name ="Command351"
                    Caption ="اصلاح"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "SAZMAN.cls"
