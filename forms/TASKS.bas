Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11396
    RowHeight =555
    ItemSuffix =34
    Left =405
    Top =1200
    Right =16515
    Bottom =10695
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x55e5f2583065e640
    End
    RecordSource ="SELECT TASKS.* FROM TASKS WHERE TASKS.PERSONEL = 77 AND TASKS.STATUS = 1 ORDER B"
        "Y TASKS.IDNUM"
    Caption ="وظايف"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =8107
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
                    Left =8212
                    Top =120
                    Width =471
                    ColumnWidth =630
                    ColumnOrder =1
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6398
                            Top =120
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8212
                    Top =460
                    Width =1260
                    ColumnWidth =0
                    ColumnOrder =2
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="GR"
                    ControlSource ="GR"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مديريت\";2;\"پشتيباني\";3;\"فروش\""
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =457
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="گروه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8212
                    Top =800
                    Width =1260
                    ColumnWidth =1845
                    ColumnOrder =3
                    TabIndex =2
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
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
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="77"
                    OnKeyPress ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6398
                            Top =800
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="مجري"
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
                    Left =8212
                    Top =1140
                    Width =1260
                    ColumnWidth =7365
                    ColumnOrder =5
                    TabIndex =3
                    Name ="TASK"
                    ControlSource ="TASK"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =1147
                            Width =495
                            Height =240
                            Name ="Label7"
                            Caption ="وظيفه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8212
                    Top =1481
                    Width =1260
                    ColumnWidth =990
                    ColumnOrder =7
                    TabIndex =4
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    Name ="PERIORITY"
                    ControlSource ="PERIORITY"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فوري\";2;\"معمولي\""
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6398
                            Top =1481
                            Width =885
                            Height =240
                            Name ="Label9"
                            Caption ="اولويت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8212
                    Top =1821
                    Width =1260
                    ColumnWidth =1215
                    ColumnOrder =6
                    TabIndex =5
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x320000000000
                    End
                    Name ="STATUS"
                    ControlSource ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"انجام نشده\";2;\"انجام شده\";3;\"لغو شده\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6398
                            Top =1821
                            Width =645
                            Height =240
                            Name ="Label11"
                            Caption ="وضعيت"
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
                    Left =8212
                    Top =2161
                    Width =471
                    ColumnWidth =945
                    ColumnOrder =8
                    TabIndex =6
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =2167
                            Width =840
                            Height =240
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
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
                    Left =8212
                    Top =2501
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =7
                    Name ="STTIME"
                    ControlSource ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =2497
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="زمان ارجاع"
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
                    Left =8212
                    Top =2841
                    Width =471
                    ColumnWidth =1170
                    ColumnOrder =10
                    TabIndex =8
                    Name ="ENDATE"
                    ControlSource ="ENDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =2842
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ انجام"
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
                    Left =8212
                    Top =3181
                    Width =471
                    ColumnWidth =870
                    ColumnOrder =11
                    TabIndex =9
                    Name ="ENTIME"
                    ControlSource ="ENTIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6405
                            Top =3187
                            Width =765
                            Height =240
                            Name ="Label19"
                            Caption ="زمان انجام"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8212
                    Top =3521
                    Width =1260
                    ColumnWidth =1725
                    ColumnOrder =12
                    TabIndex =10
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6398
                            Top =3521
                            Width =885
                            Height =240
                            Name ="Label21"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =1086
                    Top =4038
                    Width =8280
                    Height =2445
                    TabIndex =11
                    Name ="EVENTS"
                    SourceObject ="Form.EVENTS"
                    LinkChildFields ="IDNUM"
                    LinkMasterFields ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8165
                            Top =3798
                            Width =630
                            Height =240
                            Name ="EVENTS Label"
                            Caption ="EVENTS"
                            EventProcPrefix ="EVENTS_Label"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1644
                    Top =3061
                    Width =2331
                    ColumnWidth =4005
                    ColumnOrder =4
                    TabIndex =12
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"212-1-13\""
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4050
                            Top =3060
                            Width =1395
                            Height =240
                            Name ="Label24"
                            Caption ="پرونده/تماس گيرنده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2326
                    Top =2095
                    ColumnWidth =1417
                    ColumnOrder =13
                    TabIndex =13
                    Name ="num"
                    ControlSource ="num"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5293
                            Top =2095
                            Width =435
                            Height =240
                            Name ="Label27"
                            Caption ="num:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2326
                    Top =2435
                    ColumnWidth =1417
                    ColumnOrder =14
                    TabIndex =14
                    Name ="tg"
                    ControlSource ="tg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5443
                            Top =2435
                            Width =285
                            Height =240
                            Name ="Label28"
                            Caption ="tg:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1815
                    Top =3628
                    ColumnWidth =690
                    ColumnOrder =15
                    TabIndex =15
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4467
                            Top =3628
                            Width =750
                            Height =240
                            Name ="Label29"
                            Caption ="کد کاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    ColumnWidth =2010
                    ColumnOrder =16
                    TabIndex =16
                    Name ="skid"
                    ControlSource ="skid"
                    RowSourceType ="Value List"
                    RowSource ="0;\"سند حسابداري\";1;\"رسيد خريد\";2;\"حواله فروش\";3;\"برگشت خريد\";4;\"برگشت ف"
                        "روش\";20;\"پيش فاکتور\";100;\"درخواست پرداخت\";13;\"فاکتور فروش\";12;\"فاکتور خر"
                        "يد\";32;\"فاکتور خريد صادرات\";6;\"انتقال کالا از انبار به انبار\";90;\"پذيرش تع"
                        "ميرگاه\";25;\"فاکتور برگشت فروش آزاد\";24;\"ساير رسيد انبارها\";26;\"ساير حواله "
                        "انبار ها\";27;\"برگشت خريد ازاد\";33;\"فروش صادراتي\";34;\"خزانه داري\";35;\"سفا"
                        "رش\";36;درخواست خريد;37;\"مرخصي\";38;\"حواله خروج ساير\";39;\"حواله خروج مواد\";"
                        "40;\"تجهيزات PM\";41;\"درخواست کالا از انبار\";42;\"اعلام خرابي\";101;\"اتوماسيو"
                        "ن\""
                    ColumnWidths ="0"
                    DefaultValue ="101"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3000
                            Width =705
                            Height =240
                            Name ="Label26"
                            Caption ="نوع  سند"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1587
                    Top =1077
                    ColumnWidth =495
                    ColumnOrder =0
                    TabIndex =17
                    Name ="SEE"
                    ControlSource ="SEE"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00ff000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4584
                            Top =1077
                            Width =405
                            Height =240
                            Name ="Label31"
                            Caption ="ديدن"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1587
                    Top =1437
                    ColumnWidth =3210
                    TabIndex =18
                    Name ="SEET"
                    ControlSource ="SEET"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00ff000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4590
                            Top =1440
                            Width =1020
                            Height =240
                            Name ="Label33"
                            Caption ="زمان مشاهده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TASKS.cls"
