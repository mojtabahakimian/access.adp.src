Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15080
    RowHeight =315
    ItemSuffix =29
    Left =225
    Top =1020
    Right =14880
    Bottom =9555
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xd992b7507ac1e540
    End
    RecordSource ="ENHESAR_MOSHTARI"
    Caption ="انحصار"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            SpecialEffect =1
            Height =9014
            BackColor =9939701
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =2085
                    Width =14853
                    Height =6757
                    Name ="ENHESAR_KALA_FORM"
                    SourceObject ="Form.ENHESAR_KALA_FORM"
                    LinkChildFields ="EN_CUST_CO"
                    LinkMasterFields ="EN_CUST_CO"
                    OnExit ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =105
                    Top =390
                    Width =1170
                    Height =465
                    TabIndex =1
                    ForeColor =0
                    Name ="Command14"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =1284
                    Top =393
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3168
                    Top =953
                    Width =1236
                    Height =340
                    ColumnWidth =1500
                    ColumnOrder =2
                    TabIndex =3
                    BackColor =9995809
                    Name ="EN_STDATE"
                    ControlSource ="EN_STDATE"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4465
                            Top =953
                            Width =1395
                            Height =340
                            Name ="Label5"
                            Caption ="تاريخ شروع انحصار:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =135
                    Top =960
                    Height =340
                    ColumnOrder =3
                    TabIndex =4
                    BackColor =9995809
                    Name ="EN_ENDEN"
                    ControlSource ="EN_ENDEN"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1945
                            Top =960
                            Width =1200
                            Height =340
                            Name ="Label7"
                            Caption ="تاريخ پايان انحصار"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3896
                    Top =1690
                    Width =411
                    Height =340
                    ColumnWidth =2070
                    ColumnOrder =4
                    TabIndex =5
                    Name ="EN_CDATE"
                    ControlSource ="EN_CDATE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2771
                            Top =1690
                            Width =855
                            Height =340
                            Name ="Label9"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =90
                    Top =1530
                    Width =1431
                    Height =340
                    ColumnOrder =5
                    TabIndex =6
                    BackColor =9995809
                    Name ="USERID"
                    ControlSource ="USERID"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1581
                            Top =1530
                            Width =630
                            Height =340
                            Name ="Label11"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8940
                    Top =1534
                    Width =2331
                    Height =340
                    ColumnOrder =0
                    TabIndex =7
                    BackColor =9995809
                    Name ="EN_IYALAT"
                    ControlSource ="EN_IYALAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11359
                            Top =1530
                            Width =495
                            Height =340
                            Name ="Label1"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =8087
                    Top =983
                    Width =6171
                    Height =340
                    ColumnWidth =7305
                    ColumnOrder =1
                    TabIndex =8
                    BackColor =9995809
                    Name ="EN_CUST_CO"
                    ControlSource ="EN_CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    Tag ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14342
                            Top =983
                            Width =735
                            Height =340
                            Name ="Label23"
                            Caption =" مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5940
                    Top =975
                    Width =2091
                    Height =340
                    TabIndex =9
                    BackColor =9995809
                    Name ="EN_CUST_CO22"
                    ControlSource ="EN_CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    OnNotInList ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =2430
                    Top =390
                    Width =12600
                    Height =465
                    Name ="Label20"
                    Caption ="جهت انحصار براي هريک از مشتريان , هريک از گروه هاي تعريف شده زير را که لازم است "
                        "مشخص کنيد.درصورتيکه کالي خاصي مد نظر است در ستون نام کالا ,کالاي منظور را انتخاب"
                        " کنيد."
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5955
                    Top =1530
                    Width =2331
                    Height =340
                    TabIndex =10
                    BackColor =9995809
                    Name ="EN_CITY"
                    ControlSource ="EN_CITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8375
                            Top =1530
                            Width =405
                            Height =340
                            Name ="Label25"
                            Caption ="شهر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11970
                    Top =1530
                    Width =2331
                    Height =340
                    TabIndex =11
                    BackColor =9995809
                    Name ="EN_COUNTRY"
                    ControlSource ="EN_COUNTRY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Code, CountriesName FROM TCOD_Countries ORDER BY CountriesName"
                    ColumnWidths ="0"
                    DefaultValue ="100038"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14386
                            Top =1577
                            Width =465
                            Height =340
                            Name ="Label27"
                            Caption ="کشور"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
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
                    Left =2260
                    Top =1520
                    Width =2265
                    Height =340
                    TabIndex =12
                    BackColor =9995809
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="USERID_REM"
                    ControlSource ="USERID_REM"
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
                    DefaultValue ="77"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4569
                            Top =1577
                            Width =1320
                            Height =340
                            Name ="Label28"
                            Caption ="کاربر يادآوري کننده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ENHESAR1.cls"
