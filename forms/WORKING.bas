Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
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
    Width =14626
    DatasheetFontHeight =10
    ItemSuffix =55
    Left =1365
    Top =1170
    Right =10380
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x602138f4db81e240
    End
    RecordSource ="PERSONEL"
    Caption ="مشخصات شخصي و حقوق  پرسنل"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =8325
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Top =4260
                    Width =14520
                    Height =4065
                    Name ="TabCtl39"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =4665
                            Width =14256
                            Height =3530
                            Name ="Page40"
                            Caption ="كاركرد ماهيانه"
                            Begin
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    Left =186
                                    Top =4745
                                    Width =14205
                                    Height =3450
                                    Name ="WORKING_SUB"
                                    SourceObject ="Form.WORKING_SUB"
                                    LinkChildFields ="PCODE"
                                    LinkMasterFields ="CODE"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =4665
                            Width =14257
                            Height =3525
                            Name ="Page41"
                            Caption ="مرخصي"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =232
                                    Top =4745
                                    Width =14160
                                    Height =2986
                                    Name ="Child46"
                                    SourceObject ="Form.PMORAKHSUB"
                                    LinkChildFields ="CODE"
                                    LinkMasterFields ="CODE"
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =10516
                                    Top =7787
                                    Width =2031
                                    Height =315
                                    TabIndex =1
                                    BackColor =13434828
                                    Name ="Text51"
                                    ControlSource ="=IIf(MORAKH_MAND(FARSIDate(Date()),[CODE])=110110110,\"مبهم\",MORAKH_MAND(FARSID"
                                        "ate(Date()),[CODE]))"
                                    Format ="Fixed"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =12637
                                            Top =7785
                                            Width =1725
                                            Height =300
                                            Name ="Label52"
                                            Caption ="مانده مرخصي تا امروز"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =8719
                                    Top =7787
                                    Width =1725
                                    Height =300
                                    Name ="Label54"
                                    Caption ="دقيقه"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4665
                            Width =14257
                            Height =3531
                            Name ="Page42"
                            Caption ="قراردادها"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =232
                                    Top =4745
                                    Width =14160
                                    Height =3451
                                    Name ="Child50"
                                    SourceObject ="Form.PGHARARsub"
                                    LinkChildFields ="CODE"
                                    LinkMasterFields ="CODE"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =127
                            Top =4665
                            Width =14265
                            Height =3531
                            Name ="Page43"
                            Caption ="احكام حقوق"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =127
                                    Top =4745
                                    Width =14265
                                    Height =3451
                                    Name ="Child49"
                                    SourceObject ="Form.PHOKMSUB"
                                    LinkChildFields ="CODE"
                                    LinkMasterFields ="CODE"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9748
                    Top =517
                    Width =2016
                    Height =315
                    TabIndex =1
                    Name ="BIMEH_NUM"
                    ControlSource ="BIMEH_NUM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11865
                            Top =517
                            Width =885
                            Height =300
                            Name ="Label3"
                            Caption ="شماره بيمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8959
                    Top =877
                    Width =2805
                    Height =315
                    TabIndex =2
                    Name ="PNAME"
                    ControlSource ="PNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11865
                            Top =877
                            Width =300
                            Height =300
                            Name ="Label5"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8959
                    Top =1237
                    Width =2805
                    Height =315
                    TabIndex =3
                    Name ="PFAMILY"
                    ControlSource ="PFAMILY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11865
                            Top =1237
                            Width =1020
                            Height =300
                            Name ="Label7"
                            Caption ="نام خانوادگي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8959
                    Top =1597
                    Width =2805
                    Height =315
                    TabIndex =4
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11865
                            Top =1597
                            Width =1335
                            Height =300
                            Name ="Label9"
                            Caption ="شماره شناسنامه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8959
                    Top =1957
                    Width =2805
                    Height =315
                    TabIndex =5
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11865
                            Top =1957
                            Width =660
                            Height =300
                            Name ="Label11"
                            Caption ="نام پدر"
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
                    Left =6870
                    Top =165
                    Width =4896
                    Height =315
                    TabIndex =6
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR)) AS PER, NUMBER FR"
                        "OM TDETA_HES WHERE (NUMBER = 1) AND (N_KOL = 124) ORDER BY NAME + N' ' + RTRIM(C"
                        "AST(TNUMBER AS NVARCHAR))"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11867
                            Top =165
                            Width =705
                            Height =300
                            Name ="Label1"
                            Caption ="كد معين"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =10793
                    Top =2834
                    Width =999
                    Height =743
                    TabIndex =7
                    Name ="Frame28"
                    ControlSource ="SEX"
                    DefaultValue ="-1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11497
                            Top =2904
                            Name ="Option31"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =11092
                                    Top =2876
                                    Width =435
                                    Height =300
                                    Name ="Label32"
                                    Caption ="مذكر"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11494
                            Top =3277
                            OptionValue =0
                            Name ="Option33"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =11069
                                    Top =3277
                                    Width =465
                                    Height =300
                                    Name ="Label34"
                                    Caption ="مونث"
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =11858
                    Top =2990
                    Width =765
                    Height =300
                    Name ="Label35"
                    Caption ="جنسيت"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =4475
                    Top =1245
                    TabIndex =8
                    Name ="SEX"
                    ControlSource ="SEX"
                    DefaultValue ="-1"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextFontCharSet =0
                            Left =3875
                            Top =1215
                            Width =630
                            Height =240
                            Name ="Label36"
                            Caption ="جنسيت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2495
                    Top =1200
                    Width =2016
                    Height =315
                    TabIndex =9
                    Name ="WSDATE"
                    ControlSource ="WSDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4610
                            Top =1200
                            Width =1425
                            Height =300
                            Name ="Label38"
                            Caption ="تاريخ شروع به كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2495
                    Top =1560
                    Width =2016
                    Height =315
                    TabIndex =10
                    Name ="WEDATE"
                    ControlSource ="WEDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4610
                            Top =1560
                            Width =1185
                            Height =300
                            Name ="Label40"
                            Caption ="تاريخ ترك كار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10154
                    Top =2906
                    TabIndex =11
                    Name ="TAB56"
                    ControlSource ="TAB56"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =9392
                            Top =2820
                            Width =765
                            Height =300
                            Name ="Label43"
                            Caption ="تبصره اي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2505
                    Top =840
                    Width =2010
                    Height =315
                    TabIndex =12
                    Name ="DSW_BDATE"
                    ControlSource ="DSW_BDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4616
                            Top =840
                            Width =1215
                            Height =300
                            Name ="Label53"
                            Caption ="تاريخ تولد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1710
                    Top =120
                    Width =2805
                    Height =315
                    TabIndex =13
                    Name ="DSW_IDPLC"
                    ControlSource ="DSW_IDPLC"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.DSW_IDPLC FROM PERSONEL GROUP BY PERSONEL.DSW_IDPLC ORDER BY PER"
                        "SONEL.DSW_IDPLC;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4620
                            Top =120
                            Width =855
                            Height =300
                            Name ="Label49"
                            Caption ="محل صدور"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1710
                    Top =480
                    Width =2805
                    Height =315
                    TabIndex =14
                    Name ="DSW_NAT"
                    ControlSource ="DSW_NAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.DSW_NAT FROM PERSONEL GROUP BY PERSONEL.DSW_NAT ORDER BY PERSONE"
                        "L.DSW_NAT;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4616
                            Top =480
                            Width =1020
                            Height =300
                            Name ="Label51"
                            Caption ="مليت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    Left =1710
                    Top =1905
                    Width =2811
                    Height =315
                    TabIndex =15
                    Name ="JOB"
                    ControlSource ="JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.JOB FROM PERSONEL GROUP BY PERSONEL.JOB ORDER BY PERSONEL.JOB;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4605
                            Top =1905
                            Width =1185
                            Height =300
                            Name ="Label42"
                            Caption ="شغل"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1710
                    Top =2310
                    Width =2811
                    Height =315
                    TabIndex =16
                    Name ="PUNIT"
                    ControlSource ="PUNIT"
                    RowSourceType ="Value List"
                    RowSource ="1;\"توليد\";2;\"فروش\";3;\"خدمات\";4;\"اداري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4605
                            Top =2310
                            Width =1185
                            Height =300
                            Name ="Label65"
                            Caption ="محل  كاركرد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1725
                    Top =2745
                    Width =2796
                    Height =315
                    TabIndex =17
                    Name ="COT"
                    ControlSource ="COT"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4605
                            Top =2745
                            Width =1575
                            Height =405
                            Name ="Label67"
                            Caption ="شماره كارت ساعت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1725
                    Top =3105
                    Width =2796
                    Height =315
                    TabIndex =18
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =4605
                            Top =3105
                            Width =1575
                            Height =405
                            Name ="Label69"
                            Caption ="شماره حساب بانكي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8957
                    Top =2355
                    Width =2805
                    Height =315
                    TabIndex =19
                    Name ="MELLICOD"
                    ControlSource ="MELLICOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11863
                            Top =2355
                            Width =660
                            Height =300
                            Name ="Label71"
                            Caption ="كد ملي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10154
                    Top =3281
                    TabIndex =20
                    Name ="DIHARD"
                    ControlSource ="DIHARD"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =9392
                            Top =3195
                            Width =765
                            Height =300
                            Name ="Label74"
                            Caption ="جانباز"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "WORKING.cls"
