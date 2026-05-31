Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10942
    RowHeight =300
    ItemSuffix =37
    Left =3945
    Top =405
    Right =14880
    Bottom =7830
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd3ba6f910e8ce240
    End
    RecordSource ="PMORAKH"
    Caption ="ثبت و ويرايش مرخصي"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
        Begin FormHeader
            Height =630
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =10
                    Left =4081
                    Top =30
                    Width =1770
                    Height =600
                    FontSize =16
                    ForeColor =10092543
                    Name ="Label8"
                    Caption ="ثبت مرخصي"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =10
                    Left =4096
                    Top =15
                    Width =1770
                    Height =600
                    FontSize =16
                    Name ="Label9"
                    Caption ="ثبت مرخصي"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =6803
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
                    Left =4216
                    Top =870
                    Width =2106
                    Height =315
                    ColumnWidth =1485
                    TabIndex =1
                    Name ="MODATE"
                    ControlSource ="MODATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =870
                            Width =1185
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ درخواست"
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
                    Left =4216
                    Top =4125
                    Width =2106
                    Height =315
                    ColumnWidth =1365
                    TabIndex =12
                    BackColor =13434828
                    Name ="MORAKHDAY"
                    ControlSource ="MORAKHDAY"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =4125
                            Width =1290
                            Height =300
                            Name ="Label5"
                            Caption ="مرخصي به دقيقه"
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
                    Left =2176
                    Top =480
                    Width =4131
                    Height =315
                    ColumnWidth =3990
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =480
                            Width =915
                            Height =300
                            Name ="Label1"
                            Caption ="پرسنل"
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
                    Left =4216
                    Top =4505
                    Width =2106
                    Height =315
                    ColumnWidth =1800
                    TabIndex =13
                    BackColor =16777164
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =4500
                            Width =1725
                            Height =300
                            Name ="Label7"
                            Caption ="مانده مرخصي به دقيقه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5657
                    Top =3333
                    Width =666
                    Height =315
                    ColumnWidth =1545
                    TabIndex =8
                    BackColor =13434828
                    Name ="minu"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <60"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =3330
                            Width =1590
                            Height =300
                            Name ="Label11"
                            Caption ="مرخصي درخواستي"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5657
                    Top =3713
                    Width =666
                    Height =315
                    ColumnWidth =1335
                    TabIndex =11
                    BackColor =16777164
                    Name ="MMINU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =3713
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="مانده مرخصي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4934
                    Top =3333
                    Width =666
                    Height =315
                    TabIndex =7
                    BackColor =13434828
                    Name ="HOU"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null And <8"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4934
                    Top =3713
                    Width =666
                    Height =315
                    TabIndex =10
                    BackColor =16777164
                    Name ="MHOU"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4218
                    Top =3330
                    Width =666
                    Height =315
                    TabIndex =6
                    BackColor =13434828
                    Name ="DA"
                    Format ="Fixed"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    ControlTipText ="روز"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4218
                    Top =3710
                    Width =666
                    Height =315
                    TabIndex =9
                    BackColor =16777164
                    Name ="MDA"
                    Format ="Fixed"
                    DefaultValue ="0"
                    ControlTipText ="روز"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =4353
                    Top =2971
                    Width =390
                    Height =300
                    Name ="Label20"
                    Caption ="روز"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =4996
                    Top =2970
                    Width =510
                    Height =300
                    Name ="Label21"
                    Caption ="ساعت"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =5716
                    Top =2970
                    Width =510
                    Height =300
                    Name ="Label22"
                    Caption ="دقيقه"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =511
                    Top =1740
                    Width =5811
                    Height =315
                    TabIndex =3
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =1740
                            Width =1995
                            Height =240
                            Name ="Label26"
                            Caption ="توضيحات(ساعت ورود خروج)"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4216
                    Top =1305
                    Width =2106
                    Height =315
                    TabIndex =2
                    Name ="KINDM"
                    ControlSource ="KINDM"
                    RowSourceType ="Value List"
                    RowSource ="0;\"استحقاقي\";1;\"استراحت پزشكي\";2;\"ساعتي\";3;\"بدون حقوق\";4;\"تشويقي\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =1305
                            Width =915
                            Height =240
                            Name ="Label24"
                            Caption ="نوع مرخصي"
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
                    Left =4201
                    Top =2115
                    Width =2106
                    Height =315
                    TabIndex =4
                    Name ="MOSTDATE"
                    ControlSource ="MOSTDATE"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =2115
                            Width =855
                            Height =240
                            Name ="Label28"
                            Caption ="تاريخ شروع"
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
                    Left =4201
                    Top =2505
                    Width =2106
                    Height =315
                    TabIndex =5
                    Name ="MOENDATE"
                    ControlSource ="MOENDATE"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =2505
                            Width =720
                            Height =240
                            Name ="Label30"
                            Caption ="تاريخ پايان"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =8896
                    Top =654
                    Width =340
                    Height =218
                    TabIndex =14
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9286
                            Top =600
                            Width =405
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد "
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
                    Left =5551
                    Top =95
                    Width =756
                    Height =315
                    TabIndex =15
                    BackColor =16777164
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6406
                            Top =90
                            Width =855
                            Height =240
                            Name ="Label32"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    Left =496
                    Top =2145
                    Width =3525
                    Height =2775
                    FontSize =12
                    BackColor =255
                    ForeColor =10092543
                    Name ="Label33"
                    Caption ="نحوه محاسبه مرخصي: هر کارمند سالانه  26 روز مرخصي دارد که تقسيم بر 365 روز سال ="
                        " مرخصي يک روز کارکرد. اين عدد ضربدر روزهاي کارکرد که در ثبت و ويرايش کارکرد پرسن"
                        "ل  ثبت شده است ميشود ميزان مرخصي کارمند.جمعه ها حساب نميشوند"
                    FontName ="WeblogmaYekan"
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9272
                    Top =5045
                    TabIndex =16
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9526
                            Top =4995
                            Width =1155
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="درخواست کننده"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9242
                    Top =5329
                    Width =290
                    TabIndex =17
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9526
                            Top =5273
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="مدير واحد"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9272
                    Top =5588
                    TabIndex =18
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =9526
                            Top =5525
                            Width =960
                            Height =240
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
                    Left =2536
                    Top =5265
                    Width =3660
                    Height =330
                    TabIndex =19
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
                            Left =6256
                            Top =5265
                            Width =720
                            Height =240
                            Name ="Label34"
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
                    Left =7306
                    Top =5055
                    Width =1875
                    TabIndex =20
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7309
                    Top =5329
                    Width =1875
                    TabIndex =21
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
                    Left =7309
                    Top =5614
                    Width =1875
                    TabIndex =22
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =794
                    Top =6292
                    Width =1425
                    Height =405
                    TabIndex =23
                    Name ="Command54"
                    Caption ="چاپ برگه"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2219
                    Top =6292
                    Width =1426
                    Height =405
                    TabIndex =24
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =56
                    Top =1023
                    ColumnWidth =2070
                    TabIndex =25
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2261
                            Top =1020
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Top =1303
                    ColumnWidth =1455
                    TabIndex =26
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2205
                            Top =1305
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PMORAKH.cls"
