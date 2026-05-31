Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6105
    RowHeight =570
    ItemSuffix =19
    Left =1830
    Top =5865
    Right =9840
    Bottom =8040
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xf8aa9539d080e540
    End
    RecordSource ="EVENTS"
    Caption ="EVENTS"
    DatasheetFontName ="Tahoma"
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
        Begin Section
            Height =6930
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3459
                    Top =113
                    ColumnOrder =0
                    Name ="IDNUM"
                    ControlSource ="IDNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1645
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
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
                    Left =3459
                    Top =453
                    ColumnWidth =585
                    ColumnOrder =1
                    TabIndex =1
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1652
                            Top =450
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
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
                    Left =3459
                    Top =793
                    Width =2490
                    ColumnWidth =8835
                    ColumnOrder =2
                    TabIndex =2
                    Name ="EVENTs"
                    ControlSource ="EVENTS"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1645
                            Top =793
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="روي داد"
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
                    Left =3459
                    Top =1814
                    Width =2490
                    ColumnWidth =1290
                    ColumnOrder =4
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1645
                            Top =1814
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="كاربر"
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
                    Left =3679
                    Top =1077
                    Width =471
                    ColumnWidth =1020
                    ColumnOrder =5
                    TabIndex =4
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1872
                            Top =1083
                            Width =840
                            Height =240
                            Name ="Label13"
                            Caption ="تاريخ  ارجاع"
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
                    Left =3679
                    Top =1417
                    Width =471
                    ColumnWidth =885
                    ColumnOrder =6
                    TabIndex =5
                    Name ="STTIME"
                    ControlSource ="STTIME"
                    DefaultValue ="=Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1877
                            Top =1410
                            Width =870
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
                    Left =2382
                    Top =2381
                    ColumnWidth =870
                    ColumnOrder =3
                    TabIndex =6
                    Name ="SUMTIME"
                    ControlSource ="SUMTIME"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4967
                            Top =2385
                            Width =855
                            Height =240
                            Name ="Label12"
                            Caption ="ساعت كار*"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =93
                    OLETypeAllowed =1
                    Top =3174
                    ColumnOrder =7
                    TabIndex =7
                    Name ="pic"
                    ControlSource ="pic"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5760
                            Top =3180
                            Width =345
                            Height =405
                            Name ="Label14"
                            Caption ="ضميمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =6349
                    ColumnWidth =630
                    ColumnOrder =8
                    TabIndex =8
                    Name ="num"
                    ControlSource ="num"
                    DefaultValue ="Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4890
                            Top =6345
                            Width =525
                            Height =240
                            Name ="Label17"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =6689
                    ColumnWidth =735
                    ColumnOrder =10
                    TabIndex =9
                    Name ="tg"
                    ControlSource ="tg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5040
                            Top =6690
                            Width =660
                            Height =240
                            Name ="Label18"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1644
                    Top =6009
                    ColumnWidth =1830
                    ColumnOrder =9
                    TabIndex =10
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
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="101"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4644
                            Top =6009
                            Width =705
                            Height =240
                            Name ="Label26"
                            Caption ="نوع  سند"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "EVENTS.cls"
