Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15137
    RowHeight =315
    ItemSuffix =33
    Left =3045
    Top =855
    Right =18180
    Bottom =8265
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaf7c96670f5ce440
    End
    Caption ="وظايف"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
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
            Height =7426
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13706
                    Top =56
                    Width =741
                    Height =330
                    ColumnWidth =630
                    ColumnOrder =0
                    Name ="IDNUM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14502
                            Top =56
                            Width =525
                            Height =330
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    ColumnHidden = NotDefault
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
                    Left =8449
                    Top =2317
                    Width =1260
                    ColumnWidth =0
                    ColumnOrder =1
                    TabIndex =10
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="GR"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مديريت\";2;\"پشتيباني\";3;\"فروش\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10212
                            Top =2314
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="گروه"
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
                    Left =9862
                    Top =56
                    Width =3105
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =2
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000072000000030000000000000002000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000200000003000000050000000100000000000000 ,
                        0xffff9900000000000200000006000000080000000100000000000000ccffff00 ,
                        0x310000000000320000000000330000000000
                    End
                    Name ="PERSONEL"
                    RowSourceType ="Value List"
                    RowSource ="81;\"بنفشه صادقي\";77;\"خانم برگريزان\";46;\"خانم بيات\";26;\"خانم جبينيان\";64;"
                        "\"خانم خوشنود\";59;\"خانم سجاد\";119;\"خانم سميرا ابراهيمي\";100;\"خانم سهند علي"
                        " نژاد\";123;\"خانم سيما عبادي\";19;\"خانم کرمي\";45;\"خانم کارگر\";67;\"خانم دشت"
                        "ي\";139;\"خانم تارا رادنيا\";74;\"خانم رفيعي\";121;\"خانم رقيه سادات احساني\";11"
                        "5;\"خانم زکيه عرب طاهري زاده\";132;\"خانم حساني\";84;\"خانم احمدي\";40;\"خانم عب"
                        "دالوند\";9;\"خانم صادقي\";76;\"خانم صادقي نژاد\";87;\"خانم عليميرزايي\";110;\"خا"
                        "نم فاطمه جعفري\";116;\"خانم فاطمه نادريان\";3;\"خانم قويدل\";70;\"خانم مجاهد\";1"
                        "12;\"خانم مريم صمدي\";75;\"خانم محسني\";89;\"خانم مطوري\";135;\"خانم منصوره غريب"
                        "\";124;\"خانم مهنوش مسلم نيا\";125;\"خانم نرگس معين زاده\";23;\"خانم نيره قاسمي\""
                        ";109;\"خانم نينا اميراني\";127;\"خانم هانا حيدريان\";129;\"خانم هانيه ميرعلي\";3"
                        "9;\"خانم شکيبايي\";108;\"خانم شريف\";126;\"خانم شادي قامتي مقدم\";88;\"کنترل\";1"
                        "37;\"کنترل 2\";140;\"کنترل 3\";78;\"دفتر دايموند\";104;\"دفتر تهران 3\";97;\"دفت"
                        "رتهران\";36;\"دشتبان\";117;\"آقاي سپهر سلمانيان\";90;\"آقاي دکتر حبيب زارع\";18;"
                        "\"آقاي دامكي\";122;\"آقاي حسين پورباش\";83;\"آقاي حسيني نسب\";102;\"آقاي حميد فل"
                        "اح\";92;\"آقاي احمدي\";130;\"آقاي عليرضا محسني\";118;\"آقاي محمد مهدي درويشي\";5"
                        "1;\"آقاي مطلوبي\";93;\"آقاي شرافت\";101;\"تهران انبار\";103;\"رضا فلاح\";111;\"ر"
                        "يحانه رفيعي\";138;\"Farahmand\";134;\"اقاي سادات اخوي\";113;\"اقاي احسان احمدي\""
                        ";131;\"اقاي علي دايي\";94;\"الهام ميمندي\";128;\"انبار اشکذر\";30;\"علي نژاد\";7"
                        "1;\"فروشگاه يزد\";136;\"مديرفروش ميبد-آقاي دهقاني\";91;\"مرخصي\";141;\"مهسا تقي "
                        "پور\";142;\"مهندس جعفري\";69;\"ميبد\";1;\"¬Administer\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13028
                            Top =56
                            Width =495
                            Height =330
                            Name ="Label5"
                            Caption ="گيرنده"
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
                    Left =62
                    Top =566
                    Width =14160
                    Height =1020
                    ColumnWidth =5280
                    ColumnOrder =4
                    TabIndex =3
                    BackColor =13434828
                    Name ="TASK"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14336
                            Top =737
                            Width =735
                            Height =690
                            Name ="Label7"
                            Caption ="دستور/ وظيفه/ درخواست"
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
                    Left =13095
                    Top =1934
                    Width =1260
                    Height =330
                    ColumnWidth =990
                    ColumnOrder =5
                    TabIndex =4
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    Name ="PERIORITY"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فوري\";2;\"معمولي\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14456
                            Top =1934
                            Width =510
                            Height =330
                            Name ="Label9"
                            Caption ="اولويت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11055
                    Top =1934
                    Width =1260
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =5
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x320000000000
                    End
                    Name ="STATUS"
                    RowSourceType ="Value List"
                    RowSource ="1;\"انجام نشده\";2;\"انجام شده\";3;\"لغو شده\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12454
                            Top =1934
                            Width =585
                            Height =330
                            Name ="Label11"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8775
                    Top =1927
                    Width =1206
                    Height =330
                    ColumnWidth =945
                    ColumnOrder =7
                    TabIndex =6
                    Name ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10088
                            Top =1933
                            Width =840
                            Height =330
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6729
                    Top =1931
                    Width =1041
                    Height =330
                    ColumnWidth =885
                    ColumnOrder =8
                    TabIndex =7
                    Name ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7877
                            Top =1927
                            Width =780
                            Height =330
                            Name ="Label15"
                            Caption ="زمان ارجاع"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11956
                    Top =2834
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =8
                    Name ="ENDATE"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13719
                            Top =2835
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ انجام"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11956
                    Top =3174
                    Width =471
                    ColumnWidth =870
                    ColumnOrder =10
                    TabIndex =9
                    Name ="ENTIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13719
                            Top =3180
                            Width =765
                            Height =240
                            Name ="Label19"
                            Caption ="زمان انجام"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4526
                    Top =1927
                    Width =1635
                    Height =330
                    ColumnWidth =1290
                    ColumnOrder =11
                    TabIndex =11
                    Name ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6282
                            Top =1927
                            Width =375
                            Height =330
                            Name ="Label21"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =105
                    Top =3465
                    Width =14925
                    Height =3345
                    TabIndex =12
                    Name ="taskspevent"
                    SourceObject ="Form.taskspevent"
                    LinkChildFields ="IDNUM"
                    LinkMasterFields ="IDNUM"
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
                    ColumnCount =2
                    Left =105
                    Top =60
                    Width =6696
                    Height =330
                    ColumnWidth =4950
                    ColumnOrder =3
                    TabIndex =2
                    Name ="COMP_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"211-2-13\""
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6882
                            Top =60
                            Width =990
                            Height =330
                            Name ="Label24"
                            Caption ="تماس گيرنده"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10470
                    Top =6929
                    Width =2475
                    Height =340
                    TabIndex =13
                    Name ="Command26"
                    Caption ="ذخيره و خروج F2"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7995
                    Top =6929
                    Width =2475
                    Height =340
                    TabIndex =14
                    Name ="Command27"
                    Caption ="ذخيره و جديد  F3"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3030
                    Top =6929
                    Width =2475
                    Height =340
                    TabIndex =16
                    Name ="Command28"
                    Caption ="خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5518
                    Top =6929
                    Width =2475
                    Height =340
                    TabIndex =15
                    Name ="Command29"
                    Caption ="ذخيره و اضافه كردن پي نوشت F4"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7977
                    Top =56
                    Width =1830
                    Height =340
                    TabIndex =17
                    Name ="Command30"
                    Caption ="گيرندگان بيشتر"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4081
                    Top =1587
                    Width =3075
                    Height =330
                    TabIndex =18
                    Name ="USS"
                End
            End
        End
    End
End
CodeBehindForm
' See "tasksp.cls"
